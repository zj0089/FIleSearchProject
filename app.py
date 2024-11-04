import os
import pandas as pd
from PyPDF2 import PdfReader
from docx import Document

def search_files(folder_path, names):
    matches = []

    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)

        # Search in PDF files
        if file.endswith(".pdf"):
            try:
                pdf_reader = PdfReader(file_path)
                for i, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    if text and any(name in text for name in names):
                        matches.append((file_path, i + 1))
            except:
                continue

        # Search in Word files
        elif file.endswith(".docx"):
            try:
                doc = Document(file_path)
                text = "\n".join([para.text for para in doc.paragraphs])
                if any(name in text for name in names):
                    matches.append(file_path)
            except:
                continue

        # Search in Excel files
        elif file.endswith(".xlsx"):
            try:
                excel_data = pd.read_excel(file_path, engine='openpyxl')
                if any(name in str(excel_data.values) for name in names):
                    matches.append(file_path)
            except:
                continue

    return matches

if __name__ == "__main__":
    folder_path = input("Enter the folder path: ")
    names_input = input("Enter names (comma-separated): ")
    names = [name.strip() for name in names_input.split(",")]
    results = search_files(folder_path, names)

    if results:
        print("Matches found:")
        for result in results:
            print(result)
    else:
        print("No matches found.")
