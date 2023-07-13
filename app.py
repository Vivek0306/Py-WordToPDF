import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showinfo
from docx2pdf import convert

win = tk.Tk()
win.title("Word to PDF Converter App")
win.geometry("400x200")
win.resizable(False, False)

def openfile():
    file = askopenfile(filetypes=[('Word Files', '*.docx')])
    if file:
        progress_label.config(text="Converting...")
        convert(file.name, 'converted.pdf')
        showinfo("Done", "File successfully converted")
        progress_label.config(text="Conversion completed")

label = ttk.Label(win, text="Choose a Word file:")
label.grid(row=0, column=0, padx=10, pady=10, sticky="w", columnspan=2)

button = ttk.Button(win, text="Select", command=openfile)
button.grid(row=1, column=0, padx=10, pady=10, sticky="w", columnspan=2)

progress_label = ttk.Label(win, text="")
progress_label.grid(row=2, column=0, padx=10, pady=10, sticky="w", columnspan=2)

# Centering the widgets vertically and horizontally
win.grid_columnconfigure(0, weight=1)
win.grid_columnconfigure(1, weight=1)
win.grid_rowconfigure(0, weight=1)
win.grid_rowconfigure(3, weight=1)

win.mainloop()
