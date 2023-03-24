import tkinter as tk
from tkinter import filedialog
import os

cwd = os.getcwd()

window = tk.Tk()
window.withdraw()

filetypes = (
    ("Excel files", "*.xlsx"),
    ("All files", "*.*")    
)
file_path = filedialog.askopenfilename(title="Select a material xlsx file", initialdir=cwd, filetypes=filetypes)
window.deiconify()

if not file_path:
    print("no file selected")    

print(file_path)