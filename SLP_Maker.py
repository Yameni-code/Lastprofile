"""
Standardlastprofile calculation

"""

import os
import glob
from openpyxl import Workbook, load_workbook

# Constants
EXTENTION = ".xlsx"

def get_filenames(folder):
    xlsx_files = []
    for file in glob.glob( folder +"*.xlsx"):
        xlsx_files.append(file)
    return(xlsx_files)


def change_filename(filepath, week):
    new_filepath = "Lastprofile_" + week + EXTENTION
    os.rename(filepath, new_filepath)
    print("Changing filename Complet...")
    return new_filepath

def excel_operation(path1, path2, path3, path4, path5):
    wb1 = load_workbook(path1)
    ws1 = wb1.active

    wb2 = load_workbook(path2)
    ws2 = wb2.active

    wb3 = load_workbook(path3)
    ws3 = wb3.activ5

    wb4 = load_workbook(path4)
    ws4 = wb4.active

    wb5 = load_workbook(path5)
    ws5 = wb5.active

    for index in range(6, 102):
        z = "C" + str(index)
        ws1[z].value = ( float(ws1[z].value) + float(ws2[z].value) + float(ws3[z].value) + float(ws4[z].value) + float(ws5[z].value)   ) / 5

def main():
    week = "kw_5"

    print(get_filenames(week + "/"))

    names = [105216, 105229, 105242, 105249, 105257]

    file_path1 = week + "/" + "ExcelExport_20220329_" + str(names[0]) + EXTENTION
    file_path2 = week + "/" + "ExcelExport_20220329_" + str(names[1]) + EXTENTION
    file_path3 = week + "/" + "ExcelExport_20220329_" + str(names[2]) + EXTENTION
    file_path4 = week + "/" + "ExcelExport_20220329_" + str(names[3]) + EXTENTION
    file_path5 = week + "/" + "ExcelExport_20220329_" + str(names[4]) + EXTENTION

    #file_path1 = change_filename(file_path1, week)

    #excel_operation(file_path1, file_path2, file_path3, file_path4, file_path5)

    

main()

