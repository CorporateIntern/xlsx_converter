# pip install pywin32
import win32com.client as win32
import os

filepath = r"C:\Users\User\Desktop\Internship_Projects\xls_converter"
os.mkdir(filepath+"\\downloaded_files")
for dirname, _, filenames in os.walk(filepath):
    for filename in filenames:
        fname = os.path.join(dirname, filename)
        if(fname.endswith(".xls")):
            download_path = filepath+"\\downloaded_files"+"\\"
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)
            print(download_path + filename+"x")

            #FileFormat = 51 is for .xlsx extension
            #FileFormat = 56 is for .xls extension
    
            wb.SaveAs(download_path + filename+"x", FileFormat = 51)              
            wb.Close()                               
            excel.Application.Quit()
