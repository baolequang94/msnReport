import sys
from pathlib import Path
import win32com.client as win32

from config import EXCEL_PATH, AR_TEMPLATE_PATH, SHIFTS_AR_NAMES, START_COLUMN, END_COLUMN
from constant import EXCEL_APPLICATION
from helper import getDayNames, createFileName

def handleExcel(shift):

    # Open Excel
    print("Opening Excel...")
    excel = win32.Dispatch(EXCEL_APPLICATION)
    excel.Visible = False
    excel.DisplayAlerts = False
    excelApp = excel.Workbooks

    # Open Tracker file
    print("Opening Tracker...")
    trackerFile = excelApp.Open(EXCEL_PATH)

    # Open AR Template
    print("Creating AR report...")
    approvalFile = excelApp.Open(AR_TEMPLATE_PATH)

    # Open AR Tab on Tracker and copy to template
    approvalSheet = trackerFile.Worksheets(SHIFTS_AR_NAMES[shift])
    [today, tomorrow] = getDayNames()
    start = START_COLUMN + str(approvalSheet.Columns.Find(today).Row)
    end = END_COLUMN + str(approvalSheet.Columns.Find(tomorrow).Row - 1)

    approvalSheet.Range(
        start+":"+end).Copy(Destination=approvalFile.Worksheets("Sheet1").Range("A2"))

    # Create AR report
    approvalFilePath = str(Path.cwd() / (createFileName(shift)))
    approvalFile.SaveAs(approvalFilePath)
    approvalFile.Close()
    print("Completed creating AR Report!")
    return approvalFilePath

sys.modules[__name__] = handleExcel