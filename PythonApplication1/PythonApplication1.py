import xlsxwriter
import pandas
import numpy
import tkinter as tk
from tkinter import filedialog
from difflib import SequenceMatcher
output=xlsxwriter.Workbook('output.xlsx')
sheet1=output.add_worksheet()
read1=pandas.read_excel(r'C:\Users\DT123\OneDrive\Documents\Book1.xlsx',usecols=[0,3])#insertexcel1
read2=pandas.read_excel(r'C:\Users\DT123\OneDrive\Documents\Book2.xlsx',usecols=[3,6])#insertexcel2
array1=read1.to_numpy(dtype=None, copy=False)
array2=read2.to_numpy(dtype=None, copy=False)
row=1
col=0
sheet1.write(0,col,'Isolation_Point_ID')
sheet1.write(0,col+1,'Description')
for each,every in array1:
    sheet1.write(row,col,each)
    sheet1.write(row,col+1,every)
    col=1
    for one, two in array2:
        field1=str(every)
        field2=str(two)
        perc=SequenceMatcher(None,field1,field2).ratio()
        if(perc>=0.66): #Adjust this percentile for level of accuracy.
            col+=1
            sheet1.write(0,col,'FGC')
            sheet1.write(row,col,one)
            col+=1
            sheet1.write(0,col,'Functional Description')
            sheet1.write(row,col,two)
    row+=1
    col=0
output.close()

