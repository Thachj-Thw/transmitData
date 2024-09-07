from openpyxl import load_workbook
import re


convert_column = {
    "B": "W",
    "C": "E",
    "D": "B",
    "E": "G",
    "F": "P",
    "G": "R",
    "H": "Q",
    "J": "M"
}

convert_vehicle = {

}
def convert_plate(plate: str, vehicle: str) -> str:
    if re.match(r"^[0-9][0-9][aA-zZ]-", plate):
        pass
    else:
        pass
    return plate

convert_method = {
    "F": lambda v: convert_vehicle[v] if v in convert_vehicle else "",
    "G": convert_plate,
}
card_reg_temp = load_workbook("Dang Ky The.xlsx")
user_reg_temp = load_workbook("Mau dang ky thang.xlsx")

user_reg = load_workbook("card.xlsx")
