import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showinfo
from docx2pdf import convert

win=tk.Tk()
win.title("Word to Pdf Converter App")

def openfile():
  file = askopenfile(filetypes=[('Word Files', '*.docx')])
  convert(file.name, 'converted.pdf')   
  showinfo("Done", "File successfully converted ")

label=tk.Label(win,text="Choose a file!")
label.grid(row=10,column=5,padx=5,pady=5)  


button=ttk.Button(win,text="Select",width=30,command=openfile())
button.grid(row=20,column=5,padx=5,pady=5)

win.mainloop()