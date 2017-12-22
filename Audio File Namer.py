import pandas as pd
import os
from Tkinter import *
from tkFileDialog import askopenfilename
import tkMessageBox
from tkFileDialog import askdirectory


class fileConverter():

    def __init__(self):
        self.fileName = ""

        self.dirName = ""

        self.requiredWords = []
        self.requiredFiles = []

        self.root = Tk()
        self.root.title("MP3 Name Assigner...")

        self.frame1 = Frame(self.root, width=300, height=300)
        self.frame1.grid(row=0, column=0)

        self.addFrame1Elements(self.frame1)

        self.root.mainloop()

    def readExcelFie(self, excelFileName):
        df = pd.read_excel(excelFileName, header=None)
        words = list(df.iloc[:,0])
        return words

    def renameFiles(self):
        count = 0
        for file in self.requiredFiles:
                print file, self.requiredWords[count]
                os.rename(self.dirName + "/" + file,self.dirName + "/" + self.requiredWords[count]+".mp3")
                count += 1

    def selectFile(self):
        self.fileName = askopenfilename(title = "Select file",filetypes = (("excel files","*.xls"),("excel files","*.xlsx")))
        if self.fileName != "":
            Label(self.frame1, text = "File Selected: "+self.fileName).grid(row = 2, column = 1,  columnspan = 2)
            self.requiredWords = self.readExcelFie(self.fileName)
            self.requiredWords = [x for x in self.requiredWords if x != ""]
            Label(self.frame1, text = "Words found: "+str(len(self.requiredWords))).grid(row = 3, column = 1, columnspan = 2)

    def readFiles(self, dirPath):
        return sorted(os.listdir(dirPath))

    def selectDir(self):
        self.dirName = askdirectory(title="Select Directory")
        if self.dirName != "":
            Label(self.frame1, text="Directory Selected: " + self.dirName).grid(row=5, column=1, columnspan=2)
            self.requiredFiles = self.readFiles(self.dirName)
            self.requiredFiles = [x for x in self.requiredFiles if x != "" and x.endswith(".mp3")]
            Label(self.frame1, text="Files found: " + str(len(self.requiredFiles))).grid(row=6, column=1,
                                                                                             columnspan=2)

    def start(self):
        if not self.requiredWords:
            tkMessageBox.showwarning("Warning", "Please select an Excel file!")
            return

        if not self.requiredFiles:
            tkMessageBox.showwarning("Warning", "Please select a directory")
            return

        if len(self.requiredWords) != len(self.requiredFiles):
            tkMessageBox.showwarning("Warning", "Number of words in excel and Number of files in the directory are different numbers!")
        else:
            self.renameFiles()




    def addFrame1Elements(self, frame):
        # Creating the select file widget
        Label(frame, text = "\t").grid(row = 1, column = 0)
        Label(frame, text = "Select the Excel file:").grid(row = 1, column = 1, sticky = "E")
        Button(frame, text = "Select File", command = self.selectFile).grid(row=1, column = 2, sticky = "W", pady = 5, padx = 10)

        Label(frame, text = "\t").grid(row = 1, column = 0)
        Label(frame, text = "Select the Directory:").grid(row = 4, column = 1, sticky = "E")
        Button(frame, text = "Select Directory", command = self.selectDir).grid(row=4, column = 2, sticky = "W", pady = 5, padx = 10)

        Button(frame, text="Start Now", command=self.start).grid(row=9, column=3, sticky="E", ipadx=17, ipady=5)


run = fileConverter()


