import os
import sys
from datetime import datetime

import cv2
import numpy as np
import xlsxwriter
from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel

# Updated INTERVV to accommodate 6 digits by adding an extra position
# The positions are calculated to evenly distribute across the width
num_digits = 6  # Maximum number of digits expected
INTERVV = [(i + 0.5) / num_digits for i in range(num_digits)]
INTERVH = [0.33, 0.6]


def draw_interactive_polygon(image_path):
    # Initialize variables
    points = []
    moving = False
    current_point_index = -1

    original_image = cv2.flip(cv2.imread(image_path), 1)
    image = original_image.copy()
    window_name = "Polygon Drawer"

    def interpolate(p1, p2, t):
        """Linearly interpolate between points p1 and p2 with parameter t"""
        return [int((1 - t) * p1[0] + t * p2[0]), int((1 - t) * p1[1] + t * p2[1])]

    def draw_additional_lines(temp_image, points):
        """Draw lines between interpolated points on the first, third, second, and fourth edges"""
        if len(points) == 4:
            # Intervals for the first-third edge connection
            edge1_points = [interpolate(points[0], points[1], t) for t in INTERVV]
            edge3_points = [interpolate(points[2], points[3], 1 - t) for t in INTERVV]
            for p1, p3 in zip(edge1_points, edge3_points):
                cv2.line(temp_image, tuple(p1), tuple(p3), (0, 255, 255), 1)

            # Intervals for the second-fourth edge connection
            edge2_points = [interpolate(points[1], points[2], t) for t in INTERVH]
            edge4_points = [interpolate(points[3], points[0], 1 - t) for t in INTERVH]
            for p2, p4 in zip(edge2_points, edge4_points):
                cv2.line(temp_image, tuple(p2), tuple(p4), (255, 0, 255), 1)
        return temp_image

    def draw_polygon(temp_image, points, closed=False):
        if len(points) > 1:
            color = (0, 255, 0) if not closed else (255, 0, 0)
            cv2.polylines(temp_image, [np.array(points, np.int32)], closed, color, 2)
        for point in points:
            cv2.circle(temp_image, tuple(point), 5, (0, 0, 255), -1)
        temp_image = draw_additional_lines(temp_image, points)
        return temp_image

    def mouse_events(event, x, y, flags, param):
        nonlocal points, moving, current_point_index, image, original_image
        if event == cv2.EVENT_LBUTTONDOWN:
            for i, point in enumerate(points):
                if abs(x - point[0]) < 5 and abs(y - point[1]) < 5:
                    moving = True
                    current_point_index = i
                    return
            if len(points) < 4:
                points.append([x, y])
                image = draw_polygon(
                    original_image.copy(), points, closed=(len(points) == 4)
                )

        elif event == cv2.EVENT_MOUSEMOVE:
            if moving and current_point_index != -1:
                points[current_point_index] = [x, y]
                image = draw_polygon(
                    original_image.copy(), points, closed=(len(points) == 4)
                )

        elif event == cv2.EVENT_LBUTTONUP:
            moving = False
            current_point_index = -1

    cv2.namedWindow(window_name)
    cv2.setMouseCallback(window_name, mouse_events)

    # Main loop
    while True:
        cv2.imshow(window_name, image)
        key = cv2.waitKey(1) & 0xFF
        if key == 13:  # Enter key
            cv2.destroyAllWindows()
            return points  # Return the coordinates when Enter is pressed


def warp_polygon(image_path, polygon_points):
    """Warps the quadrilateral defined by polygon_points into a rectangle."""

    image = cv2.imread(image_path)
    image = cv2.flip(image, 1)

    pts_src = np.array(polygon_points).astype(np.float32)

    num_digits = len(INTERVV)
    segment_width = 40
    width = num_digits * segment_width  # Adjusted width to accommodate 6 digits
    height = 50  # Height of the rectangle
    pts_dst = np.array(
        [[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]]
    ).astype(np.float32)

    # Calculate the perspective transform matrix
    matrix = cv2.getPerspectiveTransform(pts_src, pts_dst)

    # Warp the image
    warped_image = cv2.warpPerspective(image, matrix, (width, height))
    return warped_image


def extract_line_pixels(warped_image, intervalsV, intervalsH):
    """Extracts pixel data along vertical and horizontal lines at given intervals."""
    height, width, channels = warped_image.shape

    pixel_data = []

    # Vertical lines
    for i in intervalsV:
        x_coord = int(width * i)
        vertical_line = warped_image[:, x_coord, 1]
        pixel_data.append(vertical_line)

    # Horizontal lines
    for i in intervalsH:
        y_coord = int(height * i)
        horizontal_line = warped_image[y_coord, :, 1]
        pixel_data.append(horizontal_line)

    return pixel_data


def peaks(data, boxes=2):
    # split data in boxes
    data = np.array_split(data, boxes)

    # get the maximum value in each box
    data = [np.max(d) for d in data]
    data = [d > 150 for d in data]
    return np.array(data).astype(int).tolist()


digits = {
    0: [[1, 0, 1], [1, 1], [1, 1]],
    1: [[0, 0, 0], [0, 1], [0, 1]],
    2: [[1, 1, 1], [0, 1], [1, 0]],
    3: [[1, 1, 1], [0, 1], [0, 1]],
    4: [[0, 1, 0], [1, 1], [0, 1]],
    5: [[1, 1, 1], [1, 0], [0, 1]],
    6: [[1, 1, 1], [1, 0], [1, 1]],
    7: [[1, 0, 0], [0, 1], [0, 1]],
    8: [[1, 1, 1], [1, 1], [1, 1]],
    9: [[1, 1, 1], [1, 1], [0, 1]],
}


def OCRdigit(vertical, horizontal1, horizontal2):
    # get times it goes above 150, remove subsequent values
    digit = [peaks(vertical, 3), peaks(horizontal1), peaks(horizontal2)]
    digit = [key for key, value in digits.items() if value == digit]
    return digit[0]


def OCRscreen(pixels_along_lines):
    num_digits = len(INTERVV)
    num_vertical_lines = num_digits
    num_horizontal_lines = len(INTERVH)
    data = pixels_along_lines[:num_vertical_lines]

    segment_length = int(len(pixels_along_lines[num_vertical_lines]) / num_digits)

    # Split the horizontal lines into segments
    for line in pixels_along_lines[num_vertical_lines:]:
        data.extend(
            [
                line[i * segment_length : (i + 1) * segment_length]
                for i in range(num_digits)
            ]
        )

    text = ""
    for i in range(num_digits):
        try:
            digit = str(
                OCRdigit(
                    data[i], data[i + num_digits], data[i + 2 * num_digits]
                )
            )
        except Exception:
            digit = "?"  # Use '?' for unrecognized digits
        text += digit
    return text


def writeDate(worksheet, row, column, date, format):
    original_format = "%Y-%m-%d %H-%M-%S-%f"
    parsed_datetime = datetime.strptime(date, original_format)

    worksheet.write_datetime(row, column, parsed_datetime, format)


def data2excel(data):
    # save the data in an excel file
    fileName = "data.xlsx"
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()

    dateFormat = workbook.add_format({"num_format": "dd/mm/yy hh:mm:ss"})

    # write the data
    row = 0
    for key, value in data.items():
        date = key.split(" ", 1)[1][:-4]
        writeDate(worksheet, row, 0, date, dateFormat)
        worksheet.write(row, 1, value)
        try:
            # Remove any leading '?' from value
            clean_value = value.lstrip("?")
            worksheet.write(row, 2, float(clean_value))
        except Exception as _:
            pass
        row += 1

    workbook.close()


def main(refImage):
    folder = "./"
    images = [f for f in os.listdir(folder) if f.endswith(".jpg")]

    region = draw_interactive_polygon(refImage)

    data = {}
    for image in images:
        try:
            warped = warp_polygon(folder + image, region)
            pixels_along_lines = extract_line_pixels(warped, INTERVV, INTERVH)
            data[image] = OCRscreen(pixels_along_lines)
        except Exception as e:
            print(f"Error processing {image}: {e}")
            continue

    data2excel(data)


# Initialize the Rich console
console = Console()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        imagePath = sys.argv[1]
        console.print(
            Markdown(
                f"""
> You have provided the following image: `{imagePath}`

In the next step:
  1. Draw a polygon around the device digits by clicking: top-left, top-right, bottom-right, bottom-left vertices.
  2. Confirm by pressing Enter.
                               
The program will process all images and save the results to `data.xlsx`."""
            )
        )
        input("\nPress Enter to start processing...")
        # *** Call the main function to process the image ***
        main(imagePath)
        console.print(
            Panel(
                "All data has been processed and saved to data.xlsx", style="bold green"
            )
        )
        input(
            """
Do not forget to rename the excel file to the correct name to avoid overwriting next time.
Press Enter to exit..."""
        )

    else:
        instructions = """
## WORKFLOW:

1. Place the executable in a folder with .jpg images from the device.
    - Ensure consistent camera orientation.

2. Drop an image on the executable to start:
    - Draw a polygon around the device digits by clicking: top-left, top-right, bottom-right, bottom-left vertices.
    - Confirm by pressing Enter.

3. The program processes all images and saves the results to `data.xlsx`.
        """
        console.print(
            Panel("No file was dragged onto the executable!", style="bold red")
        )
        console.print(Markdown(instructions))
        input("\nPress Enter to exit...")
