from tracemalloc import start
import xlwings as xw
import shutil
import os
from xlwings.constants import InsertShiftDirection
from datetime import datetime, timedelta


def create_timesheet_export(timesheet_dir, timesheet):
    """
    record timesheet entries from tkinter inputs
    """
    # timesheet template
    template_file = r'Import_TS_items_Hourly.xls'
    template_dir = rf'{timesheet_dir}\templates'
    template_path = os.path.join(template_dir, template_file)
    
    # location of timesheet record - source file
    record_path = os.path.join(timesheet_dir, timesheet)
    
    now = datetime.now()
    current_date = now.strftime("%Y-%m-%dT%H-%M-%S")
    date_obj = datetime.strptime(current_date, "%Y-%m-%dT%H-%M-%S")
    start_of_week = date_obj - timedelta(days=date_obj.weekday())
    end_of_week = start_of_week + timedelta(days=6)  # Sunday
    end_of_week_str = end_of_week.strftime("%Y-%m-%dT%H-%M-%S")
    # location for timesheet export - destination file
    export = rf'Import_TS_Items_Hourly - {end_of_week_str}.xls'
    export_path = os.path.join(timesheet_dir, export)

    # create copy of template
    shutil.copyfile(template_path,export_path)
    
# def copy_entries
    # TODO - Copy entries from 
    
    # # edit excel file
    # excel_app = xw.App(visible=False)
    
    # record_wb = xw.Book(record_path)
    # record_sht = record_wb.sheets['timesheet']
    
    # export_wb = xw.Book(export_path)
    # export_sht = export_wb.sheets['Sheet1']

    # export_sht.range("b2").value = end_of_week

    # #copy row
    # row = 2
    # # print(f'e{row}')
    # # print(record_sht.range('f2').value)
    # # print(start_of_week)
          
    # while record_sht.range(f'f{row}').value >= start_of_week:
    #     print(row, export_sht.range(f'f{row}').value)
    #     row += 1
        
    # # date = record_sht.range
    # # export_sht.range("a6:f6").api.Insert(
    # #     InsertShiftDirection.xlShiftDown)
    # # record_sht.range('e2:i2').api.Copy()
    # # export_sht.range('a6').api.Select()
    # # # export_sht.api.Paste()

    # export_wb.save()
    # export_wb.close()
    # record_wb.close()
    # excel_app.quit()



if __name__ == "__main__":
    directory = r"C:\Users\mark.de-melo\OneDrive - Arup\work\timesheets"
    filename = "timesheet_record.xlsx"   
    create_timesheet_export(directory, filename)


# # backup excel file
# src = path

# dst = r'C:\Users\mark.de-melo\OneDrive - Arup\time_booking'
# shutil.copy(src, dst)
