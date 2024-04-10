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


def cropImage(image_path, r, flip, cropped_folder):
    # crop the image and save it
    image = cv2.imread(image_path)
    cropped = image[int(r[1]) : int(r[1] + r[3]), int(r[0]) : int(r[0] + r[2])]

    # save with new name
    if flip == "y":
        cropped = cv2.flip(cropped, 1)

    # save in subfolder cropped
    new_name = os.path.join(cropped_folder, image_path.split("/")[-1])
    cv2.imwrite(new_name, cropped)


def data2excel(data):
    # save the data in an excel file
    fileName = "data.xlsx"
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()

    # write the data
    row = 0
    for key, value in data.items():
        date = key.split(" ", 1)[1][:-4]
        worksheet.write(row, 0, date)
        worksheet.write(row, 1, value)
        try:
            worksheet.write(row, 2, float(value[:6]))
        except Exception as _:
            pass
        row += 1

    workbook.close()


def main():
    # ocr settings
    os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
    ocr = PaddleOCR(use_angle_cls=True, lang="en")

    # *** start GUIs ***
    images = [f for f in os.listdir(FOLDER) if f.endswith(".jpg")]
    region = cropUI(os.path.join(FOLDER, images[0]))

    flip = input("Do you want to flip the images horizontaly? (y/n): ")

    cropped_folder = os.path.join(FOLDER, "cropped")
    if not os.path.exists(cropped_folder):
        os.makedirs(cropped_folder)

    # *** start cropping ***
    for image in images:
        cropImage(os.path.join(FOLDER, image), region, flip, cropped_folder)

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
    FOLDER = "data/"
    main()
