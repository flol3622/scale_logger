from paddleocr import PaddleOCR
import os
import cv2
import xlsxwriter


def ocr_image(image_path, ocr):
    result = ocr.ocr(image_path, cls=True)
    return result[0][0][1][0]


def cropUI(image_path):
    # small opencv window to crop the image
    image = cv2.imread(image_path)
    r = cv2.selectROI(image)
    cv2.destroyAllWindows()

    return r


def levelUI(image_path, region, initial_threshold=127):
    # Callback function for the trackbar
    def on_trackbar(val):
        _, binary_image = cv2.threshold(gray_image, val, 255, cv2.THRESH_BINARY)
        cv2.imshow("Binary Image", binary_image)

    # Load and convert the image to grayscale
    image = cv2.imread(image_path)
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray_image = gray_image[
        int(region[1]) : int(region[1] + region[3]),
        int(region[0]) : int(region[0] + region[2]),
    ]

    # Create a window and a trackbar
    cv2.namedWindow("Binary Image")
    cv2.createTrackbar("Threshold", "Binary Image", initial_threshold, 255, on_trackbar)

    # Initialize display
    on_trackbar(initial_threshold)
    cv2.waitKey(0)
    cv2.destroyAllWindows()


def cropImage(image_path, r, threshold, flip, cropped_folder):
    # crop the image and save it
    image = cv2.imread(image_path)
    cropped = image[int(r[1]) : int(r[1] + r[3]), int(r[0]) : int(r[0] + r[2])]
    gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY)[1]

    # save with new name
    if flip == "y":
        gray = cv2.flip(gray, 1)

    # save in subfolder cropped
    new_name = os.path.join(cropped_folder, image_path.split("/")[-1])
    cv2.imwrite(new_name, gray)


def data2excel(data):
    # save the data in an excel file
    fileName = "data.xlsx"
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()

    # write the data
    row = 0
    for key, value in data.items():
        worksheet.write(row, 0, key)
        worksheet.write(row, 1, value)
        try:
            worksheet.write(row, 2, float(value[:6]))
        except Exception as _:
            pass
        row += 1

    workbook.close()


def main():
    # ocr settings
    ocr = PaddleOCR(use_angle_cls=True, lang="en")

    # *** start GUIs ***
    images = [f for f in os.listdir(FOLDER) if f.endswith(".jpg")]
    region = cropUI(os.path.join(FOLDER, images[0]))
    threshold = levelUI(os.path.join(FOLDER, images[0]), region)

    flip = input("Do you want to flip the images horizontaly? (y/n): ")

    cropped_folder = os.path.join(FOLDER, "cropped")
    if not os.path.exists(cropped_folder):
        os.makedirs(cropped_folder)

    # *** start cropping ***
    for image in images:
        cropImage(os.path.join(FOLDER, image), region, threshold, flip, cropped_folder)

    # *** start OCR ***
    cropped_images = [f for f in os.listdir(cropped_folder) if f.endswith(".jpg")]
    data = {}
    for image in cropped_images:
        try:
            path = os.path.join(cropped_folder, image)
            text = ocr_image(path, ocr)
        except Exception as _:
            print("Error in cropped image")
            continue

        data[image] = text

    # *** save data in excel ***
    data2excel(data)

    print("All images cropped successfully")


if __name__ == "__main__":
    FOLDER = "."
    main()
