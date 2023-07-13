import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showinfo
from docx2pdf import convert

win = tk.Tk()
win.title("Word to Pdf Converter App")

def openfile():
    file = askopenfile(filetypes=[('Word Files', '*.docx')])
    if file:
        progress_label.config(text="Converting...")
        convert(file.name, 'converted.pdf')
        showinfo("Done", "File successfully converted")
        progress_label.config(text="Conversion completed")


label = tk.Label(win, text="Choose a file!")
label.grid(row=10, column=5, padx=5, pady=5)

button = ttk.Button(win, text="Select", width=30, command=openfile)
button.grid(row=20, column=5, padx=5, pady=5)

progress_label = tk.Label(win, text="")
progress_label.grid(row=30, column=5, padx=5, pady=5)

win.mainloop()
