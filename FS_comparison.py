
import os
import tkinter as tk
from tkinter import filedialog
import xlwings as xw
from pathlib import Path
import shutil, tempfile
import pandas as pd

def select_file(file_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(filetypes=[("Excel file", "*.xlsb *.xlsx")],title=f"Select {file_name} file", parent=root)
    return file_path

def open_astemporary_local(original_path: str) -> Path:
    original = Path(original_path)
    temp_dir = Path(tempfile.gettempdir())
    temp_copy = temp_dir / original.name
    shutil.copy2(original, temp_copy)
    return temp_copy


def main():
    #select wb1, new PL, new BS
    file_names = ["WB1 excel", "New Profit and Loss", "New Balance Sheet"]
    wb1_path = select_file(file_names[0])
    new_pl_path = select_file(file_names[1])
    #new_bs_path = select_file(file_names[2])

    #print(wb1_path)
    #print(new_pl_path)
    #print(new_bs_path)

    #Check if app instance is running. If yes, then close it
    for app in xw.apps:
        print(app)
        for book in app.books:
            print(book.fullname)
            if book.fullname.lower() == str(wb1_path).lower():
                book.close(save_changes=False)
        app.quit


    # Open workbook
    try:
        app = xw.App(visible=False)
        wb1 = app.books.open(open_astemporary_local(wb1_path))
        last_row = input("Enter last row number for data range in WB1: ")
        header_range = f"C5:AM5"
        value_range = f"C8:AM{last_row}"
        pl_header = wb1.sheets["1. PL Import"].range(header_range).value
        pl_values = wb1.sheets["1. PL Import"].range(value_range).value
        
        # rename the first column to 'Account'
        pl_header[0] = "Account"

        #create a dataframe from the PL data
        old_pl = pd.DataFrame(pl_values, columns=pl_header)
        
        #use index location in strip and lower the account names
        old_pl.iloc[:, 0] = old_pl.iloc[:,0].str.strip().str.lower()
        #print(old_pl.iloc[:,0])

        old_pl_unpivot = old_pl.melt(
            id_vars=["Account"],
            var_name="Month",
            value_name="Amount"
        )
        
        old_pl_unpivot["Month"] = pd.to_datetime(old_pl_unpivot["Month"], format="%Y-%m-%d %H:%M:%S").dt.strftime("%B-%Y")
        old_pl_unpivot["SortKey"] = range(len(old_pl_unpivot))

        account_and_type = pd.DataFrame(columns=["account name", "type"])
        #get account type from working PL
        try:
            account_types = {
                "revenue" : "D8:D57",
                "revenue_adj" : "D75:D102",
                "cogs" : "D117:D141",
                "cogs_others" : "D147:D182",
                "expenses_media" : "D189:D213",
                "expenes_general" : "D218:D283",
                "expenses_staffing" : "D286:D309"
            }
            

        except Exception as e:
            print(f"Error in getting account type: {e}")
        
        for item in account_types.values:
            accounts = wb1.sheets['3. Working P&L'].range(item)
        print(accounts)
        #show the amounts that are not null and not zero
        #print(old_pl_unpivot[old_pl_unpivot['Amount'].notnull() & old_pl_unpivot['Amount']!=0])
        
        #show the data types for each column
        #print(old_pl_unpivot.dtypes)
        
    except Exception as e:
        print(f"Error opening workbook: {e}")
    
    finally:
        #print(old_pl_unpivot)
        wb1.close()
        app.quit()
    
    #open new PL file from QB or Xero
    try:
        app = xw.App(visible=False)
        new_pl_wb = app.books.open(open_astemporary_local(new_pl_path))
        sheet = new_pl_wb.sheets[0]
        
        #Data starts from cell A5
        start_cell = sheet.range("A5")

        #Expand is like Ctrl+Shift+arrow down then arrow right
        used_range = start_cell.expand()
        new_pl = used_range.options(pd.DataFrame, header=1, index=False).value
        new_pl = new_pl.rename(columns={new_pl.columns[0]: "Account"})
        
        #use specific column name = "Total" to drop the total column if exists
        new_pl = new_pl.drop(columns="Total", errors='ignore')
        
        #use specific column name = "Account" to strip and lower the account names
        new_pl["Account"] = new_pl["Account"].str.strip().str.lower()

        new_pl_unpivot = new_pl.melt(
            id_vars=["Account"],
            var_name="Month",
            value_name="Amount"
        )

        new_pl_unpivot["Month"] = pd.to_datetime(new_pl_unpivot["Month"], format="%B %Y").dt.strftime("%B-%Y")
        new_pl_unpivot["SortKey"] = range(len(new_pl_unpivot))

    
    except Exception as e:
        print(f"Error opening new PL file: {e}")
    finally:
        #print(new_pl_unpivot)
        new_pl_wb.close()
        app.quit()
    
    #Compare old and new PL by merging it
    try:
        comparison_pl = pd.merge(
            old_pl_unpivot,
            new_pl_unpivot,
            on=["Account", "Month"],
            how="outer",
            suffixes=("_OldPL", "_NewPL")
        )

        comparison_pl["Month"] = pd.to_datetime(comparison_pl["Month"], format="%B-%Y")
        comparison_pl["Month_formatted"] = pd.to_datetime(comparison_pl["Month"], format="%B-%Y").dt.strftime("%B-%Y")
        comparison_pl["difference"] = comparison_pl["Amount_NewPL"] - comparison_pl["Amount_OldPL"]
        comparison_pl["Account type"] = None

        #reorder columns
        sorted_columns = [
            "Account",
            "Account type",
            "Month_formatted",
            "Amount_NewPL",
            "Amount_OldPL",
            "difference"
        ]
        not_sorted_columns = [col for col in comparison_pl.columns if col not in sorted_columns]
        comparison_pl = comparison_pl[sorted_columns + not_sorted_columns]

        #sort by SortKey and Month
        comparison_pl = comparison_pl.sort_values(by=["SortKey_NewPL", "Month"]) 

    except TypeError:
        print("Error merging dataframes. Please check the data types of 'Account' and 'Month' columns.")
    finally:
        #print(comparison_pl)
        #print(comparison_pl.dtypes)
        print("Saving comparison to CSV file...")
    
    #save to csv file
    output_file = os.path.join(os.path.dirname(wb1_path), "PL_Comparison.csv")
    comparison_pl.to_csv(output_file, index=False)
    print(f"Comparison saved to {output_file}")
    print(comparison_pl)

main()


