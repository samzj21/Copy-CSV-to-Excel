import os
import win32com.client as wincl
import time
import pandas as pd

def copy_csv_to_excel(csv_file, excel_template, excel_sheet, initial_range, final_column):
    start_time = time.time()

    if not os.path.isfile(csv_file):
        print(f"The CSV file {csv_file} does not exist.")
        return

    excel = wincl.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:

        workbook_csv = excel.Workbooks.Open(csv_file)
        sheet_csv = workbook_csv.Sheets(1)

        last_row = sheet_csv.UsedRange.Rows.Count
        csv_range = f"{initial_range}:{final_column}{last_row}"
        print(f"Copying from range: {csv_range}")

        # - define copy range #
        copy_range = sheet_csv.Range(csv_range)

        copy_range.Copy()

        workbook_excel = excel.Workbooks.Open(excel_template)
        sheet_excel = workbook_excel.Sheets(excel_sheet)

        paste_range_start = initial_range
        paste_range_end = f"{final_column}{last_row}"
        print(f"Pasting on Excel to: {paste_range_start}:{paste_range_end}")

        # - define paste range #
        paste_range = sheet_excel.Range(paste_range_start)

        # - paste only values #
        paste_range.PasteSpecial(Paste=-4163)  # -4163 corresponds to xlPasteValues # 

        workbook_excel.Save()
        workbook_excel.Close()

        workbook_csv.Close(SaveChanges=False)

    except Exception as e:
        print(f"Error during the process: {e}")

    finally:
        excel.Quit()

    end_time = time.time()
    print(f"Data copied from {csv_file}, pasted on {excel_template} in the sheet '{excel_sheet}' inside the range '{csv_range}'.")
    print(f"Execution time: {end_time - start_time:.2f} seconds.")

copy_csv_to_excel(csv_file='', initial_range='', excel_sheet='', initial_range='', final_column='')
