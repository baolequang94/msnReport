from datetime import date, timedelta
from calendar import day_name

from config import SHIFTS_END_TIME


def getDayNames():
    today = date.today()
    nextDay = today + timedelta(days=1)

    return [day_name[today.weekday()],
            day_name[nextDay.weekday()]]


def getSubject(shiftName):
    today = date.today().strftime("%A %B %d")
    return SHIFTS_END_TIME[shiftName] + " VI-VN Handover: " + today


def getFakeDays():
    today = date.today() - timedelta(days=2)
    nextDay = date.today() - timedelta(days=1)

    return [day_name[today.weekday()],
            day_name[nextDay.weekday()]]
