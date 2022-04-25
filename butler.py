import xlwings as xw
import tkinter as tk
import shutil
import os
from xlwings.constants import InsertShiftDirection
from datetime import datetime, timedelta


# timesheet input box
class MyDialog:
    def __init__(self, parent):
        top = self.top = tk.Toplevel(parent)

        self.myLabel = tk.Label(
            top, text="What have you been working on?", justify=tk.LEFT
        )
        self.myLabel.grid(row=0, columnspan=2)

        # ENTRY BOX LABELS
        self.label1 = tk.Label(top, text="Job: ", justify="left")
        self.label2 = tk.Label(top, text="Hours: ", justify="left")
        self.label3 = tk.Label(top, text="Narrative: ", justify="left")
        self.label1.grid(row=1, sticky="we")
        self.label2.grid(row=2, sticky="we")
        self.label3.grid(row=3, sticky="we")

        # ENTRY BOXES
        self.job1 = tk.Entry(top, width=40)
        self.hours1 = tk.Entry(top, width=40)
        self.narrative1 = tk.Text(top, height=2, width=40)
        self.job1.grid(row=1, column=1, sticky="we", padx=5, pady=5)
        self.hours1.grid(row=2, column=1, sticky="we", padx=5, pady=5)
        self.narrative1.grid(row=3, column=1, sticky="we", padx=5, pady=5)

        self.mySubmitButton = tk.Button(top, text="Submit", command=self.send)
        self.mySubmitButton.grid(row=4, columnspan=2)

    def send(self):
        global entries
        entries.append(self.job1.get().lower())
        entries.append(self.hours1.get())
        entries.append(self.narrative1.get(1.0, "end-1c"))
        self.top.destroy()


def input_dialogue():
    """
    Initialise input dialogue on screen
    """
    root = tk.Tk(className="Timesheet Entry")
    root.withdraw()

    inputDialog = MyDialog(root)
    root.wait_window(inputDialog.top)
    root.destroy()


def print_inputs(entries):
    """
    print entries from tkinter on command line interface
    """
    print(
        f"\nJob:\t\t{entries[0]}\
            \nHours:\t\t{entries[1]}\
            \nNarrative:\t{entries[2]}"
    )


def append_timesheet_entry(folder, filename, entries):
    """
    record timesheet entries from tkinter inputs
    """
    path = os.path.join(folder, filename)
    print(f"\nThe file is saved in: {path}")

    # edit excel file
    excel_app = xw.App(visible=False)
    book = xw.Book(path)
    now = datetime.now()

    book.sheets["timesheet"].range("a2:i2").api.Insert(InsertShiftDirection.xlShiftDown)

    # Excel row entries
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    book.sheets["timesheet"].range("a2").value = timestamp

    day = now.strftime("%a")
    book.sheets["timesheet"].range("b2").value = day

    project_id = entries[0]
    book.sheets["timesheet"].range("c2").value = project_id

    current_date = now.strftime("%Y-%m-%d")
    date_obj = datetime.strptime(current_date, "%Y-%m-%d")
    start_of_week = date_obj - timedelta(days=date_obj.weekday())
    end_of_week = start_of_week + timedelta(days=6)  # Sunday
    book.sheets["timesheet"].range("d2").value = end_of_week

    date = now.strftime("%Y-%m-%d")
    book.sheets["timesheet"].range("f2").value = date

    hours = entries[1]
    book.sheets["timesheet"].range("g2").value = hours

    charge_type = "Normal Time"
    book.sheets["timesheet"].range("h2").value = charge_type

    narrative = entries[2]
    book.sheets["timesheet"].range("i2").value = narrative

    book.save()
    book.close()
    excel_app.quit()


# def backup_timesheet():
#     """
#     save backup file if Friday and after 1400
#     """
#     if now.strftime("%a") == "Fri" and int(now.strftime("%H")) > 14:
#         backup_time = now.strftime("%Y-%m-%d")
#         backup_folder = r"C:\Users\mark.de-melo\Dropbox\Arup\timesheets"
#         backup_file = f"timesheet_record-{backup_time}.xlsm"
#         backup_path = os.path.join(backup_folder, backup_file)
#         book.save(backup_path)
#         print(f"\nA backup file is saved in: {backup_path}")
    # # backup excel file
    # src = path
    # dst = r'C:\Users\mark.de-melo\OneDrive - Arup\time_booking'
    # shutil.copy(src, dst)

if __name__ == "__main__":
    # backup_folder = r"C:\Users\mark.de-melo\OneDrive - Arup\projects\!timesheets"
    timesheet_dir = r"C:\Users\mark.de-melo\OneDrive - Arup\work\timesheets"
    timesheet_file = "timesheet_record.xlsx"
    entries = []

    # open input dialogue
    input_dialogue()
    # print all inputs from dialogue
    print_inputs(entries)
    # paste all input values into
    append_timesheet_entry(timesheet_dir, timesheet_file, entries)

