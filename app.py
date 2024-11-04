import os
import tkinter as tk
from tkinter import filedialog, messagebox, Scrollbar
import PyPDF2
import docx
import pandas as pd
from openpyxl import Workbook
from typing import List, Dict

# Function to search for names in files
def search_names_in_files(folder_path: str, names_list: List[str]) -> Dict[str, List[Dict]]:
    results = {}
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            folder_name = os.path.basename(root)

            # Skip temporary files
            if file.startswith('~$'):
                continue
            
            if file.endswith('.pdf'):
                with open(file_path, 'rb') as pdf_file:
                    reader = PyPDF2.PdfReader(pdf_file)
                    for page_number in range(len(reader.pages)):
                        page = reader.pages[page_number]
                        text = page.extract_text()
                        if text:  # Check if text extraction was successful
                            for name in names_list:
                                occurrences = text.lower().count(name.lower())
                                if occurrences > 0:
                                    if name not in results:
                                        results[name] = []
                                    results[name].append({
                                        'folder_name': folder_name,
                                        'folder_path': root,
                                        'file': file,
                                        'type': 'PDF',
                                        'page': page_number + 1,
                                        'occurrences': occurrences
                                    })

            elif file.endswith('.docx'):
                doc = docx.Document(file_path)
                section_number = 1
                text = ''
                for para in doc.paragraphs:
                    text += para.text + ' '
                    if para.text == '':  # Assuming an empty paragraph indicates a section break
                        # Process text collected so far
                        for name in names_list:
                            occurrences = text.lower().count(name.lower())
                            if occurrences > 0:
                                if name not in results:
                                    results[name] = []
                                results[name].append({
                                    'folder_name': folder_name,
                                    'folder_path': root,
                                    'file': file,
                                    'type': 'DOCX',
                                    'section': section_number,
                                    'occurrences': occurrences
                                })
                        text = ''  # Reset text for the next section
                        section_number += 1  # Increment section number

                # Process any remaining text after the last section
                if text:
                    for name in names_list:
                        occurrences = text.lower().count(name.lower())
                        if occurrences > 0:
                            if name not in results:
                                results[name] = []
                            results[name].append({
                                'folder_name': folder_name,
                                'folder_path': root,
                                'file': file,
                                'type': 'DOCX',
                                'section': section_number,
                                'occurrences': occurrences
                            })

            elif file.endswith('.xlsx'):
                df = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
                for sheet_name, sheet_data in df.items():
                    for row_idx, row in sheet_data.iterrows():
                        for col_idx, (col_name, value) in enumerate(row.items()):
                            for name in names_list:
                                if isinstance(value, str):
                                    occurrences = value.lower().count(name.lower())
                                    if occurrences > 0:
                                        if name not in results:
                                            results[name] = []
                                        results[name].append({
                                            'folder_name': folder_name,
                                            'folder_path': root,
                                            'file': file,
                                            'type': 'Excel',
                                            'sheet': sheet_name,
                                            'row': row_idx + 1,
                                            'row_name': sheet_data.index[row_idx] if hasattr(sheet_data.index, 'values') else '',  # Row label
                                            'column': col_idx + 1,
                                            'column_name': col_name,  # Column label
                                            'occurrences': occurrences
                                        })

            else:
                continue

    return results

# Function to save results to Excel using openpyxl
def save_results_to_excel(results: Dict[str, List[Dict]], output_file: str):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Search Results"

    # Add headers to the sheet
    headers = ['Name', 'Folder Name', 'Folder Path', 'File', 'Type', 'Sheet', 'Page', 'Section', 'Row', 'Row Name', 'Column', 'Column Name', 'Occurrences', 'Total Occurrences']
    sheet.append(headers)

    for name, entries in results.items():
        total_occurrences = sum(entry['occurrences'] for entry in entries)
        for entry in entries:
            row = [
                name,
                entry.get('folder_name'),
                entry.get('folder_path'),
                entry.get('file'),
                entry.get('type'),
                entry.get('sheet', ''),
                entry.get('page', ''),
                entry.get('section', ''),
                entry.get('row', ''),
                entry.get('row_name', ''),
                entry.get('column', ''),
                entry.get('column_name', ''),
                entry['occurrences'],
                total_occurrences
            ]
            sheet.append(row)

    workbook.save(output_file)

# GUI application
def run_gui_app():
    def browse_folder():
        folder = filedialog.askdirectory()
        folder_path_entry.delete(0, tk.END)
        folder_path_entry.insert(0, folder)

    def load_names():
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("Excel files", "*.xlsx")])
        if file_path:
            if file_path.endswith('.txt'):
                with open(file_path, 'r') as file:
                    names = file.read().splitlines()
                    names_text.delete('1.0', tk.END)
                    names_text.insert(tk.END, "\n".join(names))
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
                names = df.iloc[:, 0].tolist()  # Assuming names are in the first column
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
            for name, entries in results.items():
                result_text.insert(tk.END, f"{name} found in:\n")
                for entry in entries:
                    result_text.insert(
                        tk.END,
                        f"  - Folder: {entry['folder_name']} (Path: {entry['folder_path']}), File: {entry['file']} "
                        f"(Type: {entry['type']}, Sheet: {entry.get('sheet', '')}, Page: {entry.get('page', '')}, "
                        f"Section: {entry.get('section', '')}, Row: {entry.get('row', '')}, Row Name: {entry.get('row_name', '')}, "
                        f"Column: {entry.get('column', '')}, Column Name: {entry.get('column_name', '')}): {entry['occurrences']} occurrence(s)\n"
                    )
                result_text.insert(tk.END, "\n")
            save_results_to_excel(results, 'search_results.xlsx')  # Save results to Excel
        else:
            result_text.delete("1.0", tk.END)
            result_text.insert(tk.END, "No matches found.")

    # Set up the main window
    window = tk.Tk()
    window.title("File Search Application")
    window.geometry("800x600")

    # Folder selection
    tk.Label(window, text="Select Folder Path:").pack(anchor="w", padx=10)
    folder_frame = tk.Frame(window)
    folder_frame.pack(anchor="w", padx=10, pady=5)
    folder_path_entry = tk.Entry(folder_frame, width=50)
    folder_path_entry.pack(side="left", fill="x", expand=True)
    tk.Button(folder_frame, text="Browse", command=browse_folder).pack(anchor="w", padx=5, pady=5, side="left")

    # Names input with vertical scroll bar
    tk.Label(window, text="Enter Names (one per line) or Load from File:").pack(anchor="w", padx=10)
    names_frame = tk.Frame(window)
    names_frame.pack(anchor="w", padx=10, pady=5)
    names_text = tk.Text(names_frame, height=10, width=60, wrap="word")
    names_text.pack(side="left", fill="both", expand=True)
    names_v_scroll = Scrollbar(names_frame, orient="vertical", command=names_text.yview)
    names_v_scroll.pack(side="right", fill="y")
    names_text.config(yscrollcommand=names_v_scroll.set)

    # Load names button
    tk.Button(window, text="Load Names from File", command=load_names).pack(anchor="w", padx=10, pady=5)

    # Search button
    tk.Button(window, text="Search Files", command=search_files, bg="green", fg="white").pack(anchor="w", padx=10, pady=10)

    # Results display with vertical scroll bar
    tk.Label(window, text="Results:").pack(anchor="w", padx=10)
    results_frame = tk.Frame(window)
    results_frame.pack(anchor="w", padx=10, pady=5, fill="both", expand=True)
    result_text = tk.Text(results_frame, wrap="word")
    result_text.pack(side="left", fill="both", expand=True)
    results_v_scroll = Scrollbar(results_frame, orient="vertical", command=result_text.yview)
    results_v_scroll.pack(side="right", fill="y")
    result_text.config(yscrollcommand=results_v_scroll.set)

    window.mainloop()

# Run the application
run_gui_app()
