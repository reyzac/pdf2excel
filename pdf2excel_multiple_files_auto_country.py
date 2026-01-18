import tabula
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from deep_translator import GoogleTranslator
import glob

def translate_text(text):
    translator = GoogleTranslator(source='auto', target='en')
    try:
        translation = translator.translate(str(text))
        return translation
    except Exception as e:
        print(f"Translation failed {e}:  {text}")
        return text  # Return original text if translation fails
    
def main():
    os.environ["TABULA_USE_SUBPROCESS"] = "true"
    os.environ["TABULA_JAR_PATH"] = "C:/Users/Cbolasco/Downloads/tabula.jar"

    #file dialog to select folder
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # Force dialog to front
    folder_path = filedialog.askdirectory(title="Select folder with pdf files to convert", parent=root)
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))

    #Coordinates for table extraction based on country
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

    # Extract tables from all pages, auto detect country for each file
    try:
        tables_tabula = []
        
        #Detect country based on file name prefix
        for file in pdf_files:
            tables = []
            filename = os.path.basename(file)
            if filename.lower().startswith('us'):
                areas = US
                print(f"Country detected: US of file {filename}")
            elif filename.lower().startswith('ca'):
                areas = CA
                print(f"Country detected: CA of file {filename}")
            elif filename.lower().startswith('mx'):
                areas = MX
                print(f"Country detected: MX of file {filename}")
            else:
                print(f"Country not recognized for file {file}, skipping.")
                continue

            for area in areas:
                extracted_table = tabula.read_pdf(file, pages='all', area=area)
                tables.extend(extracted_table)

            for table in tables:
                #add source file column in each dataframe in list
                source_file = os.path.basename(file)
                table['source_file'] = source_file
                #print(table.head())

                #add total column if debit and credit are float types
                debits = table.columns[1]
                credits = table.columns[2]
                if pd.api.types.is_float_dtype(debits) and pd.api.types.is_float_dtype(credits):
                    table['total'] = table[debits] + table[credits]
                
                #rename columns Debit and Credit in case different language
                try:
                    new_names = {table.columns[1]: 'Debits', table.columns[2]: 'Credits'}
                    #print(f"old Column names {table.columns[1]} and {table.columns[2]}")
                    table = table.rename(columns=new_names)
                    #print(f"New column names {table.columns[1]} and {table.columns[2]}")
                except Exception as e:
                    print(f"Error renaming columns: {e}")

                #translate first column to english
                try:
                    table.iloc[:, 0] = table.iloc[:, 0].apply(translate_text)
                except Exception as e:
                    print(f"Error translating text: {e}")

                try:
                    #move FBA liquidation proceeds adjustment to bottom row
                    condition = table.iloc[:, 0] == "FBA liquidation proceeds adjustments"
                    table = pd.concat([table[-condition], table[condition]], ignore_index=True)
                except Exception as e:
                    print(f"Error moving FBA liquidation proceeds adjustment row: {e}")

                #print(table.head())
            tables_tabula.extend(tables)
    except Exception as e:
        print(f"Error during table extraction: {e}")

    #combine all dataframes 
    combined_df = pd.concat(tables_tabula, ignore_index=True)
    #print(combined_df)

    #Add total column
    combined_df["Debits"] = combined_df["Debits"].str.replace(",", "")
    combined_df["Debits"] = pd.to_numeric(combined_df["Debits"], errors="coerce").fillna(0)
    combined_df["Credits"] = combined_df["Credits"].str.replace(",", "")
    combined_df["Credits"] = pd.to_numeric(combined_df["Credits"], errors="coerce").fillna(0)

    combined_df["Total"] = combined_df["Debits"] + combined_df["Credits"]

    #Rename first column to account
    combined_df = combined_df.rename(columns={combined_df.columns[0]: "Account"})
    print(combined_df.dtypes)

    #aggregate by first column and sum the rest
    try:
        combined_df_pivot = combined_df.pivot(index="Account", columns="source_file", values="Total")
        combined_df_pivot = combined_df_pivot.reset_index()
        combined_df_groupby = combined_df.groupby(["Account", "source_file"])["Total"].sum().unstack(fill_value=0)
        combined_df_groupby = combined_df_groupby.reset_index()
        
        #Custom sort 
        sort_list = [ 'Product sales (non-FBA)','Product sale refunds (non-FBA)','FBA product sales','FBA product sale refunds','FBA inventory credit','FBA liquidation proceeds','Shipping credits','Shipping credit refunds','Gift wrap credits','Gift wrap credit refunds','Promotional rebates','Promotional rebate refunds','A-to-z Guarantee claims','Chargebacks','Amazon Shipping Reimbursement Adjustments','SAFE-T reimbursement','Seller fulfilled selling fees','FBA selling fees','Selling fee refunds','FBA transaction fees','FBA transaction fee refunds','Other transaction fees','Other transaction fee refunds','FBA inventory and inbound services fees','Shipping label purchases','Shipping label refunds','Carrier shipping label adjustments','Service fees','Refund administration fees','Adjustments','Cost of Advertising','Refund for Advertiser','Miscellaneous COGS']
        account_list_all = combined_df_groupby['Account'].unique()
        sort_list_complete = list(sort_list) + [acc for acc in account_list_all if acc not in sort_list]

        combined_df_groupby["Account"] = pd.Categorical(combined_df_groupby["Account"],
                                                        categories=sort_list_complete,
                                                        ordered=True)
        combined_df_groupby = combined_df_groupby.sort_values(['Account'])
    except Exception as e:
        print(f"Error during pivoting: {e}")
    finally:
        print(combined_df_pivot)

    

    #save combined df to one csv
    output_file1 = os.path.join(folder_path, "combined_payment_reports.csv")
    output_file2 = os.path.join(folder_path, "combined_payment_reports_pivot.csv")
    output_file3 = os.path.join(folder_path, "combined_payment_reports_groupby.csv")
    combined_df.to_csv(output_file1, index=False)
    combined_df_pivot.to_csv(output_file2, index=False)
    combined_df_groupby.to_csv(output_file3, index=False)

    

main()