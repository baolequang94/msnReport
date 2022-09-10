import inquirer, handleExcel, handleOutlook

from config import AFTERNOON_SHIFT, MORNING_SHIFT

# Promt
options = [
      inquirer.List("option",
                     message="Select shift: ",
                     choices=[MORNING_SHIFT, AFTERNOON_SHIFT],
          ),
]
selectedShift = inquirer.prompt(options)["option"]

attachmentPath = handleExcel(selectedShift)
handleOutlook(selectedShift, attachmentPath)
