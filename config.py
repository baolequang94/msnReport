EXCEL_PATH = "https://microsoft-my.sharepoint.com/personal/v-nguyenhoan_microsoft_com/Documents/Tracker.xlsx"
AR_TEMPLATE_PATH = "C:\\Users\\v-baole\\Code\\Approval_File.xlsx"

ALL_SHIFTS_PERIODS = ["7am", "8am", "9am", "10am", "11am", "12pm", "1pm", "2pm", "3pm", "4pm", "5pm", "6pm", "7pm", "8pm", "9pm", "10pm"]
MORNING_SHIFT_PERIODS = ALL_SHIFTS_PERIODS[0:10]
AFTERNOON_SHIFT_PERIODS = ALL_SHIFTS_PERIODS[-10:]

OUTPUT_AR_FILE_NAME = "Approval_08_September_2022.xlsx"

MORNING_SHIFT = "MORNING SHIFT"
AFTERNOON_SHIFT = "AFTERNOON SHIFT"

SHIFTS_AR_NAMES = {MORNING_SHIFT: "AR MORNING SHIFT",
                   AFTERNOON_SHIFT: "AR AFTERNOON SHIFT"}
SHIFTS_END_TIME = {MORNING_SHIFT: "4PM", AFTERNOON_SHIFT: "10PM"}

START_COLUMN = "C"
END_COLUMN = "F"

EMAIL_FROM = "v-baole@microsoft.com"
EMAIL_TO =["v-nguyenhoan@microsoft.com", "v-thanhle@microsoft.com", "v-mintran@microsoft.com"]
EMAIL_CC = ["v-vpoolphol@microsoft.com","v-pphewklang@microsoft.com", "v-ippanda@microsoft.com", "v-nibhushan@microsoft.com", "v-vvijayan@microsoft.com", "v-ajawal@microsoft.com", "v-preksharma@microsoft.com", "shdavies@microsoft.com"]