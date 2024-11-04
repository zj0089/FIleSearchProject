import os
import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import docx
import pandas as pd
from typing import List

# Function to search for names in files
def search_names_in_files(folder_path, names_list):
    results = {}
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith('.pdf'):
                with open(file_path, 'rb') as pdf_file:
                    reader = PyPDF2.PdfReader(pdf_file)
                    text = ''.join(page.extract_text() for page in reader.pages)
            elif file.endswith('.docx'):
                doc = docx.Document(file_path)
                text = ' '.join(para.text for para in doc.paragraphs)
            elif file.endswith('.xlsx'):
                df = pd.read_excel(file_path)
                text = ' '.join(df.to_string().split())
            else:
                continue

            for name in names_list:
                if name in text:
                    if name not in results:
                        results[name] = []
                    results[name].append(file_path)
    return results

# GUI application
def run_gui_app():
    def browse_folder():
        folder = filedialog.askdirectory()
        folder_path_entry.delete(0, tk.END)
        folder_path_entry.insert(0, folder)

    def load_names():
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'r') as file:
                names = file.read().splitlines()
                names_text.delete('1.0', tk.END)
                names_text.insert(tk.END, "\n".join(names))

    def search_files():
        folder_path = folder_path_entry.get()
        names_list = names_text.get("1.0", tk.END).strip().splitlines()

        if not folder_path or not names_list:
            messagebox.showwarning("Input Error", "Please provide a folder path and a list of names.")
            return

        results = search_names_in_files(folder_path, names_list)
        if results:
            result_text.delete("1.0", tk.END)
            for name, files in results.items():
                result_text.insert(tk.END, f"{name} found in:\n")
                for file in files:
                    result_text.insert(tk.END, f"  - {file}\n")
                result_text.insert(tk.END, "\n")
        else:
            result_text.delete("1.0", tk.END)
            result_text.insert(tk.END, "No matches found.")

    # Set up the main window
    window = tk.Tk()
    window.title("File Search Application")
    window.geometry("600x500")

    # Folder selection
    tk.Label(window, text="Select Folder Path:").pack(anchor="w", padx=10)
    folder_path_entry = tk.Entry(window, width=50)
    folder_path_entry.pack(anchor="w", padx=10, pady=5)
    tk.Button(window, text="Browse", command=browse_folder).pack(anchor="w", padx=10)

    # Names input
    tk.Label(window, text="Enter Names (one per line) or Load from File:").pack(anchor="w", padx=10)
    names_text = tk.Text(window, height=10, width=60)
    names_text.pack(anchor="w", padx=10, pady=5)
    tk.Button(window, text="Load Names from File", command=load_names).pack(anchor="w", padx=10)

    # Search button
    tk.Button(window, text="Search Files", command=search_files, bg="green", fg="white").pack(pady=10)

    # Results display
    tk.Label(window, text="Results:").pack(anchor="w", padx=10)
    result_text = tk.Text(window, height=10, width=60)
    result_text.pack(anchor="w", padx=10, pady=5)

    window.mainloop()

# Run the GUI application
if __name__ == "__main__":
    run_gui_app()
