"""
HR EXCEL AUTOMATION Script
- A Python script that automates HR tasks (mostly generating .xlsx files).

Created by: Nadeem Abdelkader on 9/5/2022
Last updated by Nadeem Abdelkader on 14/5/2022
"""

import sys

import pandas as pd
import pyexcel as p
import xlsxwriter

team_dict = {
    "Ayman Essawy": "Oracle",
    "Ayman Sakr": "Oracle",
    "Yasser Zakaria": "Oracle",
    "Mohamed Magdy": "Oracle",
    # "": "MSSQL",
    "Sameh Khedr": "SnrNetSec",
    "Amr AbdelRahman": "NetSec",
    "Ahmed Fouad Gendy": "NetSec",
    "Waleed ElSadek": "NetSec",
    "Ahmed I.Shalaby": "NetSec",
    "Hussien Magdy": "Sys",
    "Hassan Abdou": "Sys",
    "Ahmed Safwat Mousa": "Sys",
    "Momen Taher": "Oracle",
    # "Momen": "Oracle",
    "Lobna Alkomy": "Sys",
    "Mohamed Algohary": "MSSQL",
    "Ahmed Abdelgawad": "NetSec",
    "Ahmed saed": "NetSec",
}


def main(filename):
    """
    This is the main function that calls all the helper functions to generate the required .xlsx file based on the
    information read from the provided file
    :param filename: filename to read from
    :return: void (generates output file - "Actual Attendance Report.xlsx")
    """
    workbook = xlsxwriter.Workbook('Actual Attendance Report.xlsx')
    worksheet = workbook.add_worksheet()
    write_headings(workbook, worksheet)
    df = create_df(filename)
    write_rows(df, workbook, worksheet)
    workbook.close()
    return


def write_rows(df, workbook, worksheet):
    """
    This function writes the records from the pandas data frame to the output .xlsx file
    :param df: data frame to read data from
    :param workbook: workbook to write to
    :param worksheet: worksheet to write to
    :return: void
    """
    row_i = 1
    dates = df.Date.unique()
    for i in dates:
        for index, row in df.iterrows():
            cell_format = set_cell_format(row, workbook)
            if row['Date'] == i:
                write_records(cell_format, row, row_i, worksheet)
                row_i += 1
        row_i += 2
    return


def write_records(cell_format, row, row_i, worksheet):
    """
    This is a helper function that writes the formatted recrods to the output .xlsx file
    :param cell_format: format to use when writing the records
    :param row: contains the data for each record
    :param row_i: index to start from
    :param worksheet: worksheet to write to
    :return: void
    """
    try:
        worksheet.write(row_i, 0, row['Emp No.'], cell_format)
    except TypeError:
        worksheet.write(row_i, 5, "", cell_format)
    try:
        worksheet.write(row_i, 1, row['Name'], cell_format)
    except TypeError:
        worksheet.write(row_i, 5, "", cell_format)
    try:
        worksheet.write(row_i, 2, row['Date'], cell_format)
    except TypeError:
        worksheet.write(row_i, 5, "", cell_format)
    try:
        worksheet.write(row_i, 3, row['On duty'], cell_format)
    except TypeError:
        worksheet.write(row_i, 5, "", cell_format)
    try:
        worksheet.write(row_i, 4, row['Off duty'], cell_format)
    except TypeError:
        worksheet.write(row_i, 5, "", cell_format)
    try:
        worksheet.write(row_i, 5, row['Clock In'], cell_format)
    except TypeError:
        worksheet.write(row_i, 5, "-", cell_format)
    try:
        worksheet.write(row_i, 6, row['Clock Out'], cell_format)
    except TypeError:
        worksheet.write(row_i, 6, "-", cell_format)
    try:
        worksheet.write(row_i, 7, row['Late'], cell_format)
    except TypeError:
        worksheet.write(row_i, 7, "", cell_format)
    try:
        worksheet.write(row_i, 8, row['Early'], cell_format)
    except TypeError:
        worksheet.write(row_i, 8, "", cell_format)
    try:
        if row['Absent'] == True:
            worksheet.write(row_i, 9, 1, cell_format)
        else:
            worksheet.write(row_i, 9, '', cell_format)
    except TypeError:
        worksheet.write(row_i, 9, "", cell_format)
    try:
        worksheet.write(row_i, 10, "", cell_format)
    except TypeError:
        worksheet.write(row_i, 10, "", cell_format)
    try:
        worksheet.write(row_i, 11, "", cell_format)
    except TypeError:
        worksheet.write(row_i, 11, "", cell_format)
    return


def set_cell_format(row, workbook):
    """
    This function is responsible for setting the cell format based on the Employee's team, which is obtained from the
    team_dict
    :param row: row to set format of
    :param workbook: workbook to write to
    :return: cell_format: the format to should be applied to this particular row
    """
    cell_format = ""
    if team_dict[row['Name']] == "Oracle":
        cell_format = workbook.add_format(
            {'align': 'center', 'border': True, 'pattern': 1, 'bg_color': '#f7caac', 'font': 'Times New Roman'})
    elif team_dict[row['Name']] == "MSSQL":
        cell_format = workbook.add_format(
            {'align': 'center', 'border': True, 'pattern': 1, 'bg_color': '#c5e0b3', 'font': 'Times New Roman'})
    elif team_dict[row['Name']] == "SnrNetSec":
        cell_format = workbook.add_format(
            {'align': 'center', 'border': True, 'pattern': 1, 'bg_color': '#1e4e79', 'font': 'Times New Roman'})
    elif team_dict[row['Name']] == "NetSec":
        cell_format = workbook.add_format(
            {'align': 'center', 'border': True, 'pattern': 1, 'bg_color': '#bdd6ee', 'font': 'Times New Roman'})
    elif team_dict[row['Name']] == "Sys":
        cell_format = workbook.add_format(
            {'align': 'center', 'border': True, 'pattern': 1, 'bg_color': '#ffe598', 'font': 'Times New Roman'})
    return cell_format


def create_df(filename):
    """
    This function reads data from the input .xls or .xlsx to a pandas data frame
    :param filename: file to read from
    :return: df: data frame populated with data from the input .xls or .xlsx file
    """
    p.save_book_as(file_name=filename, dest_file_name=filename + "x")
    pd_xl_file = pd.ExcelFile(filename + "x")
    df = pd_xl_file.parse("Sheet 1")
    df = df[['Emp No.', 'Name', 'Date', 'On duty', 'Off duty', 'Clock In', 'Clock Out', 'Late', 'Early', 'Absent']]
    df = df.replace(to_replace='Momen', value='Momen Taher')
    return df


def write_headings(workbook, worksheet):
    """
    This function formats and writes the header cells
    :param workbook: workbook to write to
    :param worksheet: worksheet to write to
    :return: void
    """
    cell_format = workbook.add_format(
        {'border': True, 'align': 'center', 'pattern': 1, 'bg_color': '#d0cece', 'text_wrap': False, 'bold': True,
         'size': '14', 'font': 'Times New Roman'})
    worksheet.write('A1', 'FP Code', cell_format)
    worksheet.write('B1', 'Name', cell_format)
    worksheet.write('C1', 'Date', cell_format)
    worksheet.write('D1', 'On duty', cell_format)
    worksheet.write('E1', 'Off duty', cell_format)
    worksheet.write('F1', 'Clock In', cell_format)
    worksheet.write('G1', 'Clock Out', cell_format)
    worksheet.write('H1', 'late', cell_format)
    worksheet.write('I1', 'Early', cell_format)
    worksheet.write('J1', 'Missing Days', cell_format)
    worksheet.write('K1', 'Missing Days Notes', cell_format)
    worksheet.write('L1', 'Overtime Notes', cell_format)
    return


if __name__ == "__main__":
    """
    Gets input file name for the command line arguments and passes it to the main() function to start the application
    """
    main(sys.argv[1])
