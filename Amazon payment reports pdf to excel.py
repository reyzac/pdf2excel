# %%
import pdfplumber
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# %%
root = tk.Tk() #This creates a root window object—essentially the main GUI window.Even though you won’t display it, it’s required to initialize the GUI environment.
root.withdraw() #This hides the root window so it doesn’t pop up awkwardly.
root.attributes('-topmost', True)  # Force dialog to front
file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")], title="Select pdf file to convert", parent=root)
folder_path = os.path.dirname(file_path)

# %%
#Test the tables detected
with pdfplumber.open(file_path) as pdf:
    for i, page in enumerate(pdf.pages, start=1):
        text = page.extract_table({
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "intersection_tolerance": 5
        })
        print(f"Page {i} text:\n", text)

# %%
#PDFPLUMBER not working
with pdfplumber.open(file_path) as pdf:
    all_tables = []
    for page_num, page in enumerate(pdf.pages, start=1):
        tables = page.extract_tables()
        for table_num, table in enumerate(tables, start=1):
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append((page_num, table_num, df))
print(file_path)
print(all_tables)

# Combine all tables into one Excel file

# Save each table as a separate CSV file
for page_num, table_num, df in all_tables:
    file_name = f"page{page_num}_table{table_num}.csv"
    print(file_name)
    output_file = os.path.join(folder_path, file_name)
    df.to_csv(output_file, index=False)
    print(f"file saved in {output_file}")

# %%
#CAMELOT NOT WORKING
import camelot

file_name = "camelot_output.csv"
output_file = os.path.join(folder_path, file_name)
tables = camelot.read_pdf(file_path, pages="all")
print(tables)
tables.export(output_file, f="csv", compress=False)
print(output_file)



# %%
import tabula
import os

'''
To avoid RuntimeError: Can't find org.jpype.jar support library
Downloaded the tabula dependency jar file from https://github.com/tabulapdf/tabula-java/releases
renamed it to tabula.jar and then specify the path below
'''
os.environ["TABULA_USE_SUBPROCESS"] = "true"
os.environ["TABULA_JAR_PATH"] = "C:/Users/Cbolasco/Downloads/tabula.jar"


# Extract tables from all pages
try:
    tables_tabula = tabula.read_pdf(file_path, pages="all", multiple_tables=True)
except ImportError:
    print("error")

print(len(tables_tabula))
for table in tables_tabula:
    print(table)

# Save each table to CSV
for i, df in enumerate(tables_tabula):
    orig_pdf = os.path.basename(file_path)
    orig_pdf = orig_pdf.replace(".pdf", "")
    file_name = f"{orig_pdf}_table_{i+1}.csv"
    output_file = os.path.join(folder_path, file_name)
    df.to_csv(output_file, index=False)

# %% [markdown]
# <h1> This is for Payment reports

# %%
import tabula
import os

'''
To avoid RuntimeError: Can't find org.jpype.jar support library
Downloaded the tabula dependency jar file from https://github.com/tabulapdf/tabula-java/releases
renamed it to tabula.jar and then specify the path below
'''
os.environ["TABULA_USE_SUBPROCESS"] = "true"
os.environ["TABULA_JAR_PATH"] = "C:/Users/Cbolasco/Downloads/tabula.jar"

root = tk.Tk() #This creates a root window object—essentially the main GUI window.Even though you won’t display it, it’s required to initialize the GUI environment.
root.withdraw() #This hides the root window so it doesn’t pop up awkwardly.
root.attributes('-topmost', True)  # Force dialog to front
file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")], title="Select pdf file to convert", parent=root)
folder_path = os.path.dirname(file_path)

US = [
    [203.445, 21.384, 423.225, 392.634],
    [205.425, 401.544, 416.295, 773.784]
]

CA = [
    [204.435, 18.81, 385.605, 390.06],
    [205.425, 403.92, 409.365, 773.1899999999999]
]

MX = [
    [206.415, 15.84, 393.525, 396.99],
    [205.425, 402.93, 411.34499999999997, 779.13]
]

others = []
while True:
    country = input("What country? (US, CA, MX)")
    if country == "US":
        areas = US
        break
    elif country == "CA":
        areas = CA
        break
    elif country == "MX":
        areas = MX
        break

# Extract tables from all pages
try:
    tables_tabula = []
    for area in areas:
        tables = tabula.read_pdf(file_path, pages="all", area=area)
        tables_tabula.extend(tables)
except ImportError:
    print("error")

# Save each table to CSV
for i, df in enumerate(tables_tabula):
    orig_pdf = os.path.basename(file_path)
    orig_pdf = orig_pdf.replace(".pdf", "")
    file_name = f"{orig_pdf}_table_{i+1}.csv"
    output_file = os.path.join(folder_path, file_name)
    df.to_csv(output_file, index=False)


