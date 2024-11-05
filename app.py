import os
import tkinter as tk
from tkinter import filedialog, messagebox, Scrollbar
import PyPDF2
import docx
import pandas as pd
from openpyxl import Workbook
from typing import List, Dict
import re

# Function to generate name variations, ensuring spaces between names and accounting for commas
def generate_name_variations(name: str) -> List[str]:
    # Remove any spaces around commas and split on both spaces and commas
    parts = [part.strip() for part in name.replace(',', ' ').split()]
    variations = {name}  # Start with the original name

    # Generate variations based on the number of parts
    if len(parts) == 2:
        first, last = parts
        variations.add(f"{last} {first}")  # Last First
        variations.add(f"{first[0]} {last}")  # Initial Last
        variations.add(f"{last} {first[0]}")  # Last Initial
        variations.add(f"{last}, {first[0]}")  # Last, Initial

    elif len(parts) == 3:
        first, middle, last = parts
        variations.add(f"{last} {first} {middle}")  # Last First Middle
        variations.add(f"{first} {middle} {last}")  # First Middle Last
        variations.add(f"{middle} {last} {first}")  # Middle Last First
        variations.add(f"{last} {middle} {first}")  # Last Middle First
        variations.add(f"{first[0]} {middle} {last}")  # Initial Middle Last
        variations.add(f"{middle} {first[0]} {last}")  # Middle Initial Last
        variations.add(f"{last} {first[0]} {middle}")  # Last Initial Middle
        variations.add(f"{first[0]} {middle} {last}")  # Initial Middle Last
        variations.add(f"{middle[0]} {last} {first}")  # Middle Initial Last
        variations.add(f"{middle} {first} {last}")  # Middle First Last
        variations.add(f"{first} {last} {middle}")  # First Last Middle
        variations.add(f"{last}, {first[0]} {middle}")  # Last, Initial Middle
        variations.add(f"{last}, {middle} {first[0]}")  # Last, Middle Initial

    return list(variations)

# Modify search_names_in_files to include variations
def search_names_in_files(folder_path: str, names_list: List[str]) -> Dict[str, List[Dict]]:
    results = {}
    # Generate variations for each name
    names_variations = {name: generate_name_variations(name) for name in names_list}
    
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            folder_name = os.path.basename(root)

            # Skip temporary files
            if file.startswith('~$'):
                continue
            
            # Process PDF files
            if file.endswith('.pdf'):
                try:
                    with open(file_path, 'rb') as pdf_file:
                        reader = PyPDF2.PdfReader(pdf_file)
                        for page_number in range(len(reader.pages)):
                            page = reader.pages[page_number]
                            text = page.extract_text()
                            if text:  # Check if text extraction was successful
                                for name, variations in names_variations.items():
                                    for variation in variations:
                                        pattern = r'\b' + re.escape(variation) + r'\b'
                                        occurrences = len(re.findall(pattern, text, flags=re.IGNORECASE))
                                        if occurrences > 0:
                                            if name not in results:
                                                results[name] = []
                                            results[name].append({
                                                'folder_name': folder_name,
                                                'folder_path': root,
                                                'file': file,
                                                'type': 'PDF',
                                                'page': page_number + 1,
                                                'occurrences': occurrences,
                                                'variation': variation
                                            })
                except PyPDF2.errors.PdfReadError:
                    print(f"Error: Unable to read encrypted or unreadable PDF file: {file_path}")
                except Exception as e:
                    print(f"Error: Unable to process PDF file {file_path} due to {str(e)}")

            # Process DOCX files
            elif file.endswith('.docx'):
                try:
                    doc = docx.Document(file_path)
                    text = ''
                    for para in doc.paragraphs:
                        text += para.text + ' '
                    for name, variations in names_variations.items():
                        for variation in variations:
                            pattern = r'\b' + re.escape(variation) + r'\b'
                            occurrences = len(re.findall(pattern, text, flags=re.IGNORECASE))
                            if occurrences > 0:
                                if name not in results:
                                    results[name] = []
                                results[name].append({
                                    'folder_name': folder_name,
                                    'folder_path': root,
                                    'file': file,
                                    'type': 'DOCX',
                                    'occurrences': occurrences,
                                    'variation': variation
                                })
                except Exception as e:
                    print(f"Error: Unable to process DOCX file {file_path} due to {str(e)}")

            # Process XLSX files
            elif file.endswith('.xlsx'):
                try:
                    df = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
                    for sheet_name, sheet_data in df.items():
                        for row_idx, row in sheet_data.iterrows():
                            for col_idx, (col_name, value) in enumerate(row.items()):
                                if isinstance(value, str):
                                    for name, variations in names_variations.items():
                                        for variation in variations:
                                            pattern = r'\b' + re.escape(variation) + r'\b'
                                            occurrences = len(re.findall(pattern, value, flags=re.IGNORECASE))
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
                                                    'column': col_idx + 1,
                                                    'column_name': col_name,
                                                    'occurrences': occurrences,
                                                    'variation': variation
                                                })
                except Exception as e:
                    print(f"Error: Unable to process XLSX file {file_path} due to {str(e)}")

            else:
                # Skip unsupported file types without error messages
                continue

    return results

# Function to save results to Excel using openpyxl
def save_results_to_excel(results: Dict[str, List[Dict]], output_file: str):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Search Results"

    # Add headers to the sheet
    headers = ['Name', 'Name Variation', 'Folder Name', 'Folder Path', 'File', 'Type', 'Sheet', 'Page', 'Row', 'Column', 'Occurrences', 'Total Occurrences']
    sheet.append(headers)

    for name, entries in results.items():
        total_occurrences = sum(entry['occurrences'] for entry in entries)
        for entry in entries:
            row = [
                name,
                entry['variation'],  # Include the specific variation found
                entry.get('folder_name'),
                entry.get('folder_path'),
                entry.get('file'),
                entry.get('type'),
                entry.get('sheet', ''),
                entry.get('page', ''),
                entry.get('row', ''),
                entry.get('column', ''),
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
        names_list = names_text.get("1.0", tk.END).strip()
        
        # Split names by both newlines and commas, then strip whitespace
        names_list = [name.strip() for name in names_list.replace(',', '\n').splitlines() if name.strip()]

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
                        f"Row: {entry.get('row', '')}, Column: {entry.get('column', '')}): {entry['occurrences']} occurrence(s)\n"
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
