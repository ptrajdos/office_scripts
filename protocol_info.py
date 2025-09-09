#!/usr/bin/env python3
import os
import re
import csv
import argparse
from docx import Document
import pandas as pd

def process_prtocols(directory_path, output_csv_path):
    
    regexes = {
        "IDX": r"Numer albumu: (\d+)",
        "NAME": r"Imię/imiona i nazwisko dyplomanta:\s+([A-Za-zżźćńółęąśŻŹĆĄŚĘŁÓŃ\s]+)",
        "Start year": r"Studiował[a]* w latach:\s+([0-9\/]+)",
        "Poziom": r"Poziom studiów:\s+Studia\s+([I]+)",
        "Ocena Opiekuna": r"Ocena opiekuna:\s+([0-9,\.]+)",
        "Ocena Recenzenta": r"Ocena recenzenta:\s+([0-9,\.]+)",
        "Ocena Recenzenta 2": r"Ocena dodatkowego recenzenta:\s+([0-9,\.]+)",
    }

    colnames = ["ID"]
    for name in regexes.keys():
        colnames.append(name)

    out_df = pd.DataFrame(columns=colnames)

    for root, _, files in os.walk(directory_path):
        for  file_idx,file in enumerate(files):
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                row = dict()
                row["ID"] = file_idx 
                try:
                    doc = Document(file_path)
                    for i, paragraph in enumerate(doc.paragraphs, start=1):
                        for name, pattern in regexes.items():
                            match = re.search(pattern, paragraph.text)
                            if match:
                                # Store all groups joined by space
                                row[name] = match.group(1)
                            else:
                                row[name] = ""      
                    for i, table in enumerate(doc.tables, start=1):
                        for row_idx, table_row in enumerate(table.rows, start=1):
                            for cell_idx, cell in enumerate(table_row.cells, start=1):
                                for name, pattern in regexes.items():
                                    match = re.search(pattern, cell.text)
                                    if match:
                                        row[name] = match.group(1)
                                    else:
                                        if name not in row:
                                            row[name] = ""
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")

                out_df = pd.concat([out_df, pd.DataFrame([row])], ignore_index=True)

    out_df.to_csv(output_csv_path, index=False,sep=';')

if __name__ == "__main__":
    # --- Parse command-line arguments ---
    parser = argparse.ArgumentParser(description="Search multiple regexes in DOCX files and save results to CSV.")
    parser.add_argument("input_dir", help="Directory containing DOCX files")
    parser.add_argument("output_csv", help="Output CSV file path")
    args = parser.parse_args()
    directory_path =  args.input_dir
    output_csv_path = args.output_csv
    process_prtocols(directory_path, output_csv_path)
    