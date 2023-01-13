from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

excel_path = "template.xlsx"

# load excel file
excel_wb = load_workbook(excel_path)
# select a sheet
excel_worksheet = excel_wb["Hoja 1"]

categories = {"gomas": "gomas", "cascos": "cascos"}

db = [
    {
        "name": "goma manillas ",
        "category": categories["gomas"],
        "image": "images/gomas1.jpeg",
        "price": 0.90,
    },
    {
        "name": "goma pata de cambio",
        "category": categories["gomas"],
        "image": "images/gomas3.jpeg",
        "price": 0.95,
    },
    {
        "name": "goma pata de cambio pro taper",
        "category": categories["gomas"],
        "image": "images/gomas4.jpeg",
        "price": 1,
    },
    {
        "name": "casco modelo 1",
        "category": categories["gomas"],
        "image": "images/casco1.jpeg",
        "price": 30,
    },
]


def write_single_product(
    product: tuple,
    worksheet,
    name_cell: str,
    category_cell: str,
    img_cell_name: str,
    price_cell: int,
):
    def insert_image():
        img = Image(product["image"])

        img.height = 95  # insert image height in pixels as float or int (e.g. 305.5)
        img.width = 95  # insert image width in pixels as float or int (e.g. 405.8)
        img.anchor = img_cell_name  # where you want image to be anchored/start from
        worksheet.add_image(img, img_cell_name)  # adding in the image

    name_cell.value = product["name"].upper()
    category_cell.value = product["category"].upper()
    price_cell.value = round(product["price"], 2) 
    insert_image()


def write_total_cells(last_row_product: int):
    sub_total_str_cell = excel_worksheet[f"E{last_row_product}"]
    sub_total_str_cell.value = "SUB TOTAL $"
    sub_total_formula_cell = excel_worksheet[f"H{last_row_product}"]
    sub_total_formula_cell.value = f"=SUMA(H13:H{last_row_product - 1})"

    total_str_cell = excel_worksheet[f"H{last_row_product + 1}"]
    total_str_cell.value = "TOTAL $"
    total_formula_cell = excel_worksheet[f"H{last_row_product + 1}"]
    total_formula_cell.value = f"=SUMA(H13:H{last_row_product - 1})"


def run_app():
    # first row
    row_count = 13

    for product in db:
        category_cell = excel_worksheet[f"B{row_count}"]
        name_cell = excel_worksheet[f"C{row_count}"]
        img_cell = f"D{row_count}"
        price_cell = excel_worksheet[f"F{row_count}"]

        write_single_product(
            product, excel_worksheet, name_cell, category_cell, img_cell, price_cell
        )

        # next row
        row_count = row_count + 1

    # when loop finishes...
    write_total_cells(row_count)

    excel_wb.save(excel_path)
    os.startfile(excel_path)


if __name__ == "__main__":
    run_app()
