"""
 -*- coding: utf-8 -*-
----------------------
Description: Tool to record work being done and save to a timesheet
             spreadsheet. Backup file to OneDrive

 Created on: 2021-01-25, 10:13
 Author: Mark de Melo
----------------------
"""

import xlwings as xw
import tkinter as tk
import shutil
import os
from xlwings.constants import InsertShiftDirection
from datetime import datetime


# timesheet input box
class MyDialog:
    def __init__(self, parent):
        top = self.top = tk.Toplevel(parent)
        self.myLabel = tk.Label(top, text="What have you been working on?")
        self.myLabel.grid(row=0, columnspan=2)

        self.label1 = tk.Label(top, text="Job 1: ").grid(row=1, sticky="we")
        self.label2 = tk.Label(top, text="Narrative 1: ").grid(row=2, sticky="we")
        self.label3 = tk.Label(top, text="Hours 1: ").grid(row=3, sticky="we")
        self.label4 = tk.Label(top, text="Journal 1: ").grid(row=4)
        self.label5 = tk.Label(top, text="Job 2: ").grid(row=5, sticky="we")
        self.label6 = tk.Label(top, text="Narrative 2: ").grid(row=6, sticky="we")
        self.label7 = tk.Label(top, text="Hours 2: ").grid(row=7, sticky="we")
        self.label8 = tk.Label(top, text="Journal 2: ").grid(row=8)

        self.job1 = tk.Entry(top)
        self.narrative1 = tk.Entry(top)
        self.hours1 = tk.Entry(top)
        self.journal1 = tk.Text(top, height=4, width=40)
        self.job2 = tk.Entry(top)
        self.narrative2 = tk.Entry(top)
        self.hours2 = tk.Entry(top)
        self.journal2 = tk.Text(top, height=4, width=40)

        self.job1.grid(row=1, column=1, sticky="we", padx=5, pady=5)
        self.narrative1.grid(row=2, column=1, sticky="we", padx=5, pady=5)
        self.hours1.grid(row=3, column=1, sticky="we", padx=5, pady=5)
        self.journal1.grid(row=4, column=1, pady=5)
        self.job2.grid(row=5, column=1, sticky="we", padx=5, pady=5)
        self.narrative2.grid(row=6, column=1, sticky="we", padx=5, pady=5)
        self.hours2.grid(row=7, column=1, sticky="we", padx=5, pady=5)
        self.journal2.grid(row=8, column=1, padx=5, pady=5)

        self.mySubmitButton = tk.Button(top, text="Submit", command=self.send)
        self.mySubmitButton.grid(row=9, columnspan=2)

    def send(self):
        global entries
        entries.append(self.job1.get().lower())
        entries.append(self.narrative1.get())
        entries.append(self.hours1.get())
        entries.append(self.journal1.get("1.0", "end-1c"))
        entries.append(self.job2.get().lower())
        entries.append(self.narrative2.get())
        entries.append(self.hours2.get())
        entries.append(self.journal2.get("1.0", "end-1c"))
        self.top.destroy()


def input_dialogue():
    """
    Initialise input dialogue on screen
    """
    root = tk.Tk()

    root.withdraw()

    inputDialog = MyDialog(root)
    root.wait_window(inputDialog.top)
    root.destroy()


def print_inputs(entries):
    """
    print entries from tkinter on command line interface
    """
    print(f"\n{entries[0]} \t{entries[1]} \t{entries[2]} \n{entries[3]}")
    print(f"\n{entries[4]} \t{entries[5]} \t{entries[6]} \n{entries[7]}")


def record_inputs_xlsx(folder, filename, entries):
    """
    record timesheet entries from tkinter inputs
    """
    path = os.path.join(folder, filename)
    print(f"\nThe file is saved in: {path}")

    # edit excel file
    excel_app = xw.App(visible=False)
    book = xw.Book(path)
    now = datetime.now()

    # job 1 entries
    total_entries = 2
    entry_counter = 0
    while entry_counter < total_entries:
        entry_counter = entry_counter * 4
        # timesheet entry
        if entries[entry_counter]:
            book.sheets[2].range("a3:o3").api.Insert(InsertShiftDirection.xlShiftDown)
            # timestamp
            book.sheets[2].range("b3").value = now.strftime(
                "%Y-%m-%d %H:%M:%S")
            # Day
            day = now.strftime("%a")
            book.sheets[2].range("c3").value = day
            # Project id
            # Day
            date = now.strftime("%Y-%m-%d")
            book.sheets[2].range("g3").value = date
            # Hours            
            book.sheets[2].range("h3").value = entries[entry_counter + 2].strip()
            book.sheets[2].range("d3").value = entries[entry_counter].strip()  # Narrative
            book.sheets[2].range("j3").value = entries[entry_counter + 1].strip()
            # Ignore?
            book.sheets[2].range("k3").value = "No"
            
        # journal entry
        if entries[entry_counter + 3]:
            book.sheets[1].range("a2:f2").api.Insert(InsertShiftDirection.xlShiftDown)
            book.sheets[1].range("a2").value = now.strftime("%Y-%m-%d %H:%M:%S")
            book.sheets[1].range("b2").value = now.strftime("%a")  # day
            book.sheets[1].range("c2").value = now.strftime("%Y-%m-%d")  # date
            book.sheets[1].range("d2").value = entries[
                entry_counter
            ].strip()  # project code
            book.sheets[1].range("f2").value = entries[
                entry_counter + 3
            ].strip()  # journal entry

        entry_counter += 1

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


if __name__ == "__main__":
    # backup_folder = r"C:\Users\mark.de-melo\OneDrive - Arup\projects\!timesheets"
    folder = r"C:\Users\mark.de-melo\OneDrive - Arup\work\timesheets"
    filename = "timesheet_record.xlsx"
    entries = []

    # open input dialogue
    input_dialogue()
    # print all inputs from dialogue
    print_inputs(entries)
    # paste all input values into
    record_inputs_xlsx(folder, filename, entries)


# # backup excel file
# src = path
# dst = r'C:\Users\mark.de-melo\OneDrive - Arup\time_booking'
# shutil.copy(src, dst)
