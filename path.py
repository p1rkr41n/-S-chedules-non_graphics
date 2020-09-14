#Get path file
from tkinter import *
from tkinter import filedialog
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def getpath(): 
    #Tk().geometry('1200x1100')
    Tk().wm_iconbitmap('icon.ico') #Change icon
    #Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename( initialfile =  "/", title = "Choose source file excel (non-support Word)", filetype = (("excel files","*.xlsx *.xls" ),("all files","*.*")) ) # show an "Open" dialog box and return the path to the selected file
    return filename


#Check end
#print ('===Successfull getpath.py===')