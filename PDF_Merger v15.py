#This is a Python PDF Merger Application I have developed for collegues at work.
#It will compress images, convert images and Excel Spreadsheets to PDF file format and merge them into one PDF.
#This is an super easy user interface application.
#This was created in order to greatly reduce time for approx. 100 engineers nationwide time taken to submit expenses each month and save money for our company.

import os
import img2pdf
import time
import os.path
import sys
import win32com.client
import win32api
import subprocess
import PIL
import tempfile
from tkinter.ttk import Progressbar
from PIL import Image
from pathlib import Path
from PyPDF2 import PdfFileMerger
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import ttk
from tkinter import *

#Initiate main function
def main():
    
    #Set progress bar to 0
    bar["value"] = 0
    bar.update()

    #Start timer
    start_time = time.time()

    #Kill Acrobat Reader
    try:
        subprocess.Popen("taskkill /F /im AcroRd32.exe", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except:
        print("Acrobat Reader could not be terminated")
    
    #Open file dialog to select file
    dirname = filedialog.askopenfilename(title=' Please select a folder containing .JPG, .PDF or .XLSX files ')
    dirname = os.path.dirname(dirname)

    #Always appear on top of screen
    root.lift()
    
    bar["value"] = bar["value"] + 1
    bar.update()

    #Check if file Converted_images.pdf is present and delete it
    filePath = dirname + "/Converted_images.pdf"
    try:
        os.remove(filePath)
    except:
        time.sleep(0)

    bar["value"] = bar["value"] + 1
    bar.update()

    #Check if file !Final_merged_PDF.pdf is present and delete it
    filePath1 = dirname + "/!Final_merged_PDF.pdf"
    try:
        os.remove(filePath1)
    except:
        time.sleep(0)

    bar["value"] = bar["value"] + 1
    bar.update()

    #Throw an error if file dialog box was closed/file not selected
    try:
        imgdir = os.chdir(dirname)
    except:
        v.set("Directory not selected.")
        root.update_idletasks()
        style.configure("black.Horizontal.TProgressbar", background='red')

    #Convert EXCEL to PDF
    #Acceptable Excel file types
    excelext = (".xlsx", ".xls", ".xlsm", ".csv")
    imgdir = os.chdir(dirname)
    for file in os.listdir(dirname):
        style.configure("black.Horizontal.TProgressbar", background='SpringGreen4')
        bar.update()
        v.set("")
        root.update_idletasks()

        if file.endswith(tuple(excelext)):
	    
            #If filename starts with ~$, ignore it as it is Excel temporary file
            if file.startswith("~$"):
                time.sleep(0)
            else:
		#Check if any Excel instances with same filename are open
                try:
                    xl = win32com.client.Dispatch("Excel.Application")
                    xl.Visible = False
                    xl.DisplayAlerts = False
                    xl.EnableEvents = False
                    wb = xl.Workbooks.Open(os.getcwd() + "/" + file)

		#Throw an exception if Excel Workbook is open to stop data loss
                except:
                    v.set("Found an Open Unsaved Excel Workbook, Save and Close any Excel instances.")
                    root.update_idletasks()
                    style.configure("black.Horizontal.TProgressbar", background='red')
                    bar.update()
	
		#Convert Excel Spreadsheet to PDF
                v.set("Excel file being converted: " + file)
                root.update_idletasks()
                ws = wb.Worksheets(1)
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesTall = 1
                ws.PageSetup.FitToPagesWide = 1
                filename = os.path.splitext(file)
                filename = filename[0]

                wb.ExportAsFixedFormat(0, os.getcwd() + "\!" + str(filename) + ".pdf")

                style.configure("black.Horizontal.TProgressbar", background='SpringGreen4')
                bar["value"] = bar["value"] + 5
                bar.update()
                wb.Close(True)
                xl.Quit()


    #PNG Image to JPG Converter
    imgpng = (".png", ".PNG")
    for png in os.listdir(dirname):
        if png.endswith(tuple(imgpng)):
                pngimg = os.path.splitext(png)
                pngimg = pngimg[0]
                im = PIL.Image.open(png)
                rgb_im = im.convert('RGB').save(pngimg + ".jpg" ,"JPEG")
                v.set("PNG Image converted: " + png)
                root.update_idletasks()
                bar["value"] = bar["value"] + 1
                bar.update()

    #Resize images as some images from e.g. iPhones are over 5MB and it takes 5x of these and final PDF will not be send over email
    #Below are acceptable image file types
    imgres = (".jpg", ".jpeg", ".JPG", ".JPEG")
    res = 1024
    for images in os.listdir(os.getcwd()):
        if images.endswith(tuple(imgres)):
            im = PIL.Image.open(images)
            width, height = im.size
            if width <= res:
                time.sleep(0)
            elif height <= res:
                time.sleep(0)
            else:
                basewidth = 1024
                img = PIL.Image.open(images)
                wpercent = (basewidth/float(img.size[0]))
                hsize = int((float(img.size[1])*float(wpercent)))
                img = img.resize((basewidth,hsize), PIL.Image.ANTIALIAS)
                img.save(images)
                v.set("Image resized: " + images)
                root.update_idletasks()
                bar["value"] = bar["value"] + 1
                bar.update()

    #Convert Images to PDF
    ext = (".jpg", ".jpeg", ".JPG", ".JPEG")
    for fname in os.listdir(os.getcwd()):
        if fname.endswith(tuple(ext)):
            with open("Converted_images.pdf", "wb") as f:
                f.write(img2pdf.convert([i for i in os.listdir(os.getcwd()) if i.endswith(tuple(ext))]))
                time.sleep(0.1)
                v.set("Images converted: " + fname)
                root.update_idletasks()
                bar["value"] = bar["value"] + 1
                bar.update()

    # Merge all PDFs in same directory
    def list_files(dirname, extension):
        return (y for y in (sorted(os.listdir(dirname), key=str.lower)) if y.endswith('.' + extension))
    pdfs = list_files(dirname, "pdf")
    merger = PdfFileMerger(strict = False)
    for pdf in pdfs:
        merger.append(pdf)
        v.set("File merged: " + pdf)
        root.update_idletasks()
        bar["value"] = bar["value"] + 1
        bar.update()
        time.sleep(0.5)

    try :
        merger.write('!Final_merged_PDF.pdf')
        merger.close()
        bar["value"] = 45
        bar.update()

    except Exception as error:
        v.set("Close Adobe Acrobat Reader Session and restart PDF Merger Application")
        root.update_idletasks()
        time.sleep(15)
        style.configure("black.Horizontal.TProgressbar", background='red')
        bar["value"] = 50
        bar.update()

    #Remove temporary converted PDF file
    for file in os.listdir(dirname):
        if file.endswith(tuple(excelext)):
                filename = os.path.splitext(file)
                filename = filename[0]
                if os.path.isfile(os.getcwd() + "\!" + filename + ".pdf"):
                    os.remove("!" + filename + ".pdf")
                    style.configure("black.Horizontal.TProgressbar", background='SpringGreen4')
                    bar["value"] = bar["value"] + 1
                    bar.update()


    #Remove merged images PDF
    for file in os.listdir(dirname):
        if os.path.isfile(os.getcwd() + "\Converted_images.pdf"):
            os.remove("Converted_images.pdf")
            bar["value"] = bar["value"] + 1
            bar.update()

    #Update status bar
    v.set("")
    root.update_idletasks()

    #Counter stopped, calculate how many second it took to complete all tasks
    counter = time.time() - start_time
    counter = round(counter)

    style.configure("black.Horizontal.TProgressbar", background='SpringGreen4')
    v.set("Success! It has taken " + str(counter) + " seconds to merge all files!")
    root.update_idletasks()

    bar["value"] = 50
    bar.update()

    #Give it a break
    time.sleep(1)

    #Checks if file is present and opens it
    filePath1 = dirname + "/!Final_merged_PDF.pdf"
    try:
        os.startfile(filePath1)
        bar["value"] = bar["value"] + 1
        bar.update()
    except:
        print()



####################################################################################################


#This code will make sure application starts in the center of windows screen
def center(win):
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()


#Initiate Tkinter GUI screen before starting the application
if __name__ == "__main__":

    #Transparent application icon
    ICON = (b'\x00\x00\x01\x00\x01\x00\x10\x10\x00\x00\x01\x00\x08\x00h\x05\x00\x00'
        b'\x16\x00\x00\x00(\x00\x00\x00\x10\x00\x00\x00 \x00\x00\x00\x01\x00'
        b'\x08\x00\x00\x00\x00\x00@\x05\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00'
        b'\x00\x01\x00\x00\x00\x01') + b'\x00'*1282 + b'\xff'*64

    _, ICON_PATH = tempfile.mkstemp()
    with open(ICON_PATH, 'wb') as icon_file:
        icon_file.write(ICON)

    #Start Tkinter GUI
    root = Tk()
    root.resizable(width=False, height=False)
    root.title("PDF Merger v15")
    root.attributes('-alpha', 0.0)
    root.iconbitmap(default=ICON_PATH)

    #Initiate status bar updates
    v = StringVar()
    v.set("-")
    root.update_idletasks()

    #Instructions text box
    instructions = ("""
    MAKE SURE TO SAVE AND CLOSE ANY OPEN EXCEL SESSIONS!
    YOU COULD LOSE UNSAVED PROGRESS IN YOUR EXCEL EXPENSES FORMS!

    1.   Create a folder for merging the files.
            > e.g. /Documents/Expenses/January Expenses 2020
            > name this folder Cash and Fuel Expenses
    2.   Copy Expenses/Timesheets January 2020 Excel file (.xlxs/.xls) into above folder.
    3.   Copy all needed PDF receipt files (.pdf) into above folder.
    4.   Copy all needed pictures of receipts files (.jpg/.jpeg/.png) into above folder.
    5.   Click "Run PDF Merger".
    6.   Dialog Window will pop up to select the folder for merging.
    7.   Navigate to the inside of the folder you have created in Step 1.
    8.   Select Expenses/Timesheets Excel file which you have copied in Step 2.
    9.   Click "Open" button or Double-Click on the Excel file.
    10. Application will convert all Excel, PDF and Image files into one PDF file.
    11. Once merger has completed, file !Final_merged_PDF.pdf will appear in above folder.
    12. Merged file will open automatically after completion.
    """)

    #Set a frame to instructions text box
    text_frame = Frame(root)
    text_frame.grid(column=0, row=0)
    text = Label(text_frame, text=instructions, font=("MS Sans Serif", 7, 'bold'), justify="left", relief = GROOVE, bg = "white")
    text.grid(column = 0, row = 0, pady = 5, padx = 5)

    #Set up frame and locations of status and progress bars
    status_frame = Frame(root)
    status_frame.grid(column = 0, row = 2, sticky = W)
    status_bar = Label(status_frame, textvariable = v, font = ("MS Sans Serif", 7, 'bold'))
    status_bar.grid(column = 0, row = 0, padx = 5, sticky = W)
    bar_frame = Frame(root)
    bar_frame.grid(column = 0, row = 3)

    #Set style of the progress bar
    style = ttk.Style()
    style.theme_use('alt')
    style.configure("black.Horizontal.TProgressbar", background='SpringGreen4', thickness = 5)
    bar = Progressbar(bar_frame, length=525, style="black.Horizontal.TProgressbar", mode ="determinate", phase = 10)
    bar.grid(column=0, row=0)
    bar['maximum'] = 50
    bar["value"] = 0
    bar.update()

    #Set up frame and location of the Run button
    button_frame = Frame(root)
    button_frame.grid(column = 0, row = 4)
    run_button = Button(button_frame, text = "Run PDF Merger", command = lambda : main(), width = 25, font=("MS Sans Serif", 7, "bold"),bg="gray88")
    run_button.grid(column = 2, row = 0, pady = 3, padx = 5, sticky = E)

    #Display application in the center of the screen
    center(root)
    root.attributes('-alpha', 1.0)
    root.mainloop()
    pass
