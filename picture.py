import win32com.client
import os
from PIL import ImageGrab
import time
def export_excel_range_to_image(excel_path, range_address, output_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        workbook = excel.Workbooks.Open(excel_path)
        sheet = workbook.Sheets(1)
        sheet.Range(range_address).CopyPicture(Format=2)
        time.sleep(1)
        img = ImageGrab.grabclipboard()
        if not os.path.exists(os.path.dirname(output_path)):
            os.makedirs(os.path.dirname(output_path))
        img.save(output_path, "PNG")
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()
excel_file = r"C:\Users\admin\Desktop\agenda.xlsx"
range_to_export = "A1:H16"
output_file = r"C:\Users\admin\Desktop\agenda.png"
export_excel_range_to_image(excel_file, range_to_export, output_file)

