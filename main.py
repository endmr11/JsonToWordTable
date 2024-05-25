import tkinter as tk
from tkinter import filedialog, messagebox
import json
from docx import Document

def addTable(table, json_data, level=0):
    for key, value in json_data.items():
        if isinstance(value, dict):
            row_cells = table.add_row().cells
            row_cells[0].text = '  ' * level + key
            row_cells[1].text = ''
            addTable(table, value, level + 1)
        elif isinstance(value, list):
            row_cells = table.add_row().cells
            row_cells[0].text = '  ' * level + key
            row_cells[1].text = ''
            for item in value:
                if isinstance(item, dict):
                    addTable(table, item, level + 1)
                else:
                    row_cells = table.add_row().cells
                    row_cells[0].text = '  ' * (level + 1)
                    row_cells[1].text = str(item)
        else:
            row_cells = table.add_row().cells
            row_cells[0].text = '  ' * level + key
            row_cells[1].text = str(value)

def jsonToWord(json_path, word_path):
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        doc = Document()
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Key'
        hdr_cells[1].text = 'Value'

        addTable(table, json_data)
        doc.save(word_path)
        messagebox.showinfo("Success", "")
    except Exception as e:
        messagebox.showerror("Error", f"{e}")

def selectJsonFile():
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if file_path:
        json_entry.delete(0, tk.END)
        json_entry.insert(0, file_path)

def selectOutputFile():
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if file_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, file_path)

def convert():
    json_path = json_entry.get()
    output_path = output_entry.get()
    if json_path and output_path:
        jsonToWord(json_path, output_path)
    else:
        messagebox.showwarning("Warning", "Please select all file paths.")

root = tk.Tk()
root.title("JSON to Word Converter")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

json_label = tk.Label(frame, text="JSON:")
json_label.grid(row=0, column=0, sticky=tk.W)
json_entry = tk.Entry(frame, width=50)
json_entry.grid(row=0, column=1, padx=5)
json_button = tk.Button(frame, text="Select", command=selectJsonFile)
json_button.grid(row=0, column=2, padx=5)

output_label = tk.Label(frame, text="Output Word:")
output_label.grid(row=1, column=0, sticky=tk.W)
output_entry = tk.Entry(frame, width=50)
output_entry.grid(row=1, column=1, padx=5)
output_button = tk.Button(frame, text="Select", command=selectOutputFile)
output_button.grid(row=1, column=2, padx=5)

convert_button = tk.Button(frame, text="Convert", command=convert)
convert_button.grid(row=2, columnspan=3, pady=10)

root.mainloop()
