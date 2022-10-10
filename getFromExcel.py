import time
import xlwings as xw


#
# Open the macro file
#

def getExcelValue(wb):
    # select the sheet where data is sent
    sheet = wb.sheets['SCAN_PAGE']

    # get the value from B3
    value = sheet.range("E11").value

    return str(value).replace('.0', '')


def getCompleteStatus(wb):
    # select the sheet where data is sent
    sheet = wb.sheets['SCAN_PAGE']

    # get the value from B3
    value = sheet.range("E6").value

    return value


def getGatePassID(wb):
    # select the sheet gate pass id is
    sheet = wb.sheets['SCAN_PAGE']

    # get the value from B3
    value = sheet.range("F6").value

    return str(value).replace('.0', '')


def getLocation(wb):
    # get the unloading location
    sheet = wb.sheets['SCAN_PAGE']

    # get the value from cell
    value = sheet.range("G6").value

    return str(value).replace('.0', '')


def getDoorLocation(wb):
    # get the door location
    sheet = wb.sheets['SCAN_PAGE']

    # get the value from cell
    value = sheet.range("H6").value

    return str(value).replace('.0', '')
