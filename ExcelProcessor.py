import os
import pandas as pd
from PolioCleaner import PolioCleaner
import re
import openpyxl
import zipfile
import lxml.etree as ET
from Dataform import DataForm
import warnings

query_list = ["SIAs-0" "SIAs-0 ",
            "SIAs-1", "SIAs-1 ", "SIAs-2", "SIAs-2 ", "SIAs-3",
            "Additional Request (RD1)", "Additional Request (RD2)", "Round 0", "Round 1"]

warnings.filterwarnings("ignore", category=UserWarning, message="Cannot parse header or footer so it will be ignored")


class ExcelProcessor:
    def __init__(self, data_directory):
        self.directory = os.path.normpath(data_directory)
        self.clean_directory = os.path.join(self.directory, 'clean')
        if not os.path.exists(self.clean_directory):
            os.makedirs(self.clean_directory)
        self.campaign_id = self.extract_campaign_id()
        self.xl_files = [f for f in os.listdir(self.directory) if f.endswith('.xlsx')]
    
        # Retrieves campaign ID from directory path
    def extract_campaign_id(self):
        folder_name = os.path.basename(os.path.normpath(self.directory))
        # Extract the 4 digits after 'CE' from the folder name
        match = re.search(r'CE(\d{4})', folder_name)
        if match:
            digits = match.group(1)
            return int(digits)
        else:
            raise ValueError("The folder name does not contain 'CE' followed by four digits")
            return 0000

    def process_file(self, filename):
            #INSERT DECRYPTION HERE
            dataframe = []
            filepath = os.path.join(self.directory, filename)
            # Ensure the file extensions are correct
            if not filename.endswith('.xlsx'):
                raise ValueError("must have a .xlsx extension.")
            # Load the Excel file
            try:
                with pd.ExcelFile(filepath, engine='openpyxl') as xls:
                    for sheet_name in xls.sheet_names:
                            if sheet_name in query_list:
                                print(f'sheet name: {sheet_name}')
                                df = xls.parse(sheet_name=sheet_name)
                                sheet = DataForm(df, sheet_name, self.directory, filename)
                                clean_sheet= PolioCleaner(sheet)
                                clean_sheet.clean()
                                dataframe.append(clean_sheet.sheet.df)
            except zipfile.BadZipFile:
                print(f"File {filename} contains sensitivity tag or is opened elsewhere")
                print(f'ERROR: {filename} EXCLUDED FROM DATASET')
                return
            except Exception as e:
                print(f"Failed to save {filename} due to {e}")
                return
            self.save_excel(dataframe, filename)
            
    def save_excel(self, df_list, filename):
        # Creates clean directory if it doesn't already exist
        if not os.path.exists(self.clean_directory):
            os.makedirs(self.clean_directory)
            print(f"'clean' directory created at: {self.clean_directory}")
        output = os.path.join(self.clean_directory, filename)
        combined_df = pd.concat(df_list, ignore_index=True)

        # Ensures phase column is saved as int type
        combined_df['Phase'] = combined_df['Phase'].astype(int)

        # Save the combined dataframe to an excel file
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name=f'temp', index=False)
        print(f"{filename} data saved with sheet name 'temp'")

    # Concatenates Province level data into final xlsx
    def concat_xlsx_files(self):
        df_list = []
        # List all files in the directory
        files = [f for f in os.listdir(self.clean_directory) if f.endswith('.xlsx')]

        # Loop through the files and read them into pandas dataframes
        for file in files:
            file_path = os.path.join(self.clean_directory, file)
            df = pd.read_excel(file_path, engine='openpyxl')
            df_list.append(df)

        # Concatenate all dataframes into a single dataframe
        combined_df = pd.concat(df_list, ignore_index=True)

        # Save the combined dataframe to a new Excel file
        output_path = os.path.join(self.clean_directory, f"{self.campaign_id}_clean.xlsx")

        # save final Excel file and clean directory
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name=f"{self.campaign_id}", index=False)
        print(f"{self.campaign_id} excel files have been combined and saved to {output_path}")
        self.remove_original_xlsx_files()

    # Directory clean-up    
    def remove_original_xlsx_files(self):
        # List all files in the directory
        files = [f for f in os.listdir(self.clean_directory) if f.endswith('.xlsx')]
        # Loop through the files and remove them, excluding the final combined file
        for file in files:
            if not file.endswith('_clean.xlsx'):
                file_path = os.path.join(self.clean_directory, file)
                os.remove(file_path)
                print(f"Removed {file}")
            print(f"Temporaray files removed from {self.clean_directory}")

    def extract_additional_budget_request(self, xls_file):
        pass
        # Load the Excel file into a DataFrame
        # Find the sheet that contains the word "update"
        update_sheet_name = None
        for sheet_name in xls_file.sheet_names:
            if 'update' in sheet_name.lower():
                update_sheet_name = sheet_name
                break
        # If an update sheet is found, extract it
        if update_sheet_name:
            df = xls_file.parse(update_sheet_name)
            return df
        else:
            print("No sheet containing the word 'update' was found.")
            return None
        
    def clean_additional_budget_request(self, df):
        # Locate the starting point (first row in column 1 with the cell value 'A')
        start_row = df[df.iloc[:, 0] == 'A'].index[0]
        
        # Iterate down the column to find the third consecutive empty cell
        empty_count = 0
        end_row = start_row
        for i in range(start_row + 1, len(df)):
            if pd.isna(df.iloc[i, 0]):
                empty_count += 1
                if empty_count == 3:
                    end_row = i
                    break
            else:
                empty_count = 0
        
        # Extract the desired range of rows and the first 12 columns
        extracted_df = df.iloc[start_row:end_row, :12]



