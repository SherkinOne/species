import openpyxl
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Entry, OptionMenu, Button
from tkinter.font import Font
import hi

root=Tk()
root.title("Database parser")
root.geometry("500x200+200+200")
dataBaseList=["Marine", "Land"]

# Variables

labelFont=Font(family="Arial", size=14)
noOfFractals=StringVar()
lengthOfFractals=StringVar()
listChoice=StringVar()
sheetChoice=StringVar()
theFile=" "
sheets= []
global wb

#functions

def clearOptions() :
    # empty the entry boxes
    noLabel=Label(root, text="", font=labelFont)
    
def openFile():   
    root.filename =filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xlsx files","*.xlsx"), ("All files","*.*")))
    # show the file name
    theFile = root.filename
    if theFile is not () :
        noLabel.config(text=theFile)
        # show the sheets 
        wb = openpyxl.load_workbook(filename=theFile , read_only=False)
        sheets=wb.get_sheet_names()
        sheetBox=OptionMenu(root, sheetChoice, sheets[0], *sheets)
        sheetBox.grid(row=7,column=1,columnspan=1)
    
def search() :
        # These are the parameters that will be passed to the functions
        try: 
            hi.searchDatabase(noLabel["text"],sheetChoice.get(), dataBaseList[dataBaseList.index(listChoice.get())])
        except :
            print()


# tkinter
mainLabel=Label(root, text="Please choose from the following :", font=("arial","16"))
mainLabel.grid(row=0,column=0, columnspan=2)

typeLabel=Label(root, text="Which database :", font=labelFont)
typeLabel.grid(row=1, column=1)

noLabel=Label(root, text=theFile, font=labelFont)
noLabel.grid(row=2, column=0,  columnspan=5)

listBox=OptionMenu(root, listChoice, dataBaseList[0], *dataBaseList)
listBox.grid(row=3,column=1)

clearButton=Button(root, text="Clear", command=clearOptions)
clearButton.grid(row=4, column=2)

openButton=Button(root, text="Open File", command=openFile)
openButton.grid(row=4, column=1)

searchButton=Button(root, text="Check File", command=search)
searchButton.grid(row=5, column=1)

details=Label(root, text="Please choose - sheet data must be in column A ", font=labelFont)
details.grid(row=6, column=0,  columnspan=5)

mainloop()