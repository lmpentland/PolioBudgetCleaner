import pandas as pd
import re
import os
from pprint import pprint



class PolioCleaner:
    def __init__(self, sheet):
        self.sheet = sheet
        self.dataframes = []
        # Moves Cost Category

    def column_1_fix(self, val):
        first_char = str(val)[0]
        if first_char.isnumeric():
            try:
                first_int = int(first_char)
                if float(first_int) < 2:
                    return "1. Finger Markers"
                elif float(first_int) < 3:
                    return "2. Human resources and incentives"
                elif float(first_int) < 4:
                    return "3. Training and Meetings"
                elif float(first_int) < 5:
                    return "4. Supplies and Equipment"
                elif float(first_int) < 6:
                    return "5. Transportation"
                elif float(first_int) < 7:
                    return "6. Social mobilisation and communication"
                elif float(first_int) < 8:
                    return "7. Vaccine Management" 
                elif float(first_int) < 9:
                    return "8. Other Operational Costs"
                elif float(first_int) >= 9:
                    return "9. PPE Operational Costs"
            except (ValueError, TypeError, IndexError):
                return val
        else:
            try:
                if first_char == "A":
                    return "2. Human resources and incentives"
                elif first_char == "B":
                    return "3. Training and Meetings"
                elif first_char == "C":
                    return "4. Supplies and Equipment"
                elif first_char == "D":
                    return "5. Transportation"
                elif first_char == "E":
                    return "6. Social mobilisation and communication"
                elif first_char == "F":
                    return "7. Vaccine Management" 
                elif first_char == "G":
                    return "8. Other Operational Costs"
            except (ValueError, TypeError, IndexError):   
                return val

    # Function to check if the value is numeric or contains only a dash
    def is_valid(self, value):
        if pd.isna(value):
            return False
        # Convert to string for regex check
        value_str = str(value)
        return bool(re.match(r'^[\d-]+$', value_str))

    # Use extracted data to initalize 3 columns
    def add_columns(self):
        temp_df = self.sheet.df
        num_rows = temp_df.shape[0]
        new_columns = {
            'Campaign ID': [self.sheet.campaign_id] * num_rows,
            'Phase': [self.sheet.phase] * num_rows,
            'Country': [self.sheet.country] * num_rows,
            'State': [self.sheet.state] * num_rows,
            'Round': [self.sheet.round_number] * num_rows
        }
        columns_df = pd.DataFrame(new_columns)
        # pprint(temp_df.iloc[:2])

        # Concatenate the new DataFrame onto the beginning of the existing 'extracted' DataFrame
        result_df = pd.concat([columns_df, temp_df], join='inner', axis=1, ignore_index=True)
        pprint(result_df.iloc[:5])
        self.sheet.df = result_df

   
    
    # Function to apply multiple checks and modify data accordingly
    def apply_checks_and_modify(self, row):
        def is_empty(value):
            if pd.isna(value) or value =="":
                return True
        # Check if the third column is empty and the second column has any data
        if is_empty(row.iloc[2]) and not is_empty(row.iloc[1]):
            row.iloc[2] = '-'
        # Keep the row unless both the second and third columns are empty
        return not (is_empty(row.iloc[1]) and is_empty(row.iloc[2]))


    # Assign proper column labels
    def relabel_columns(self):
        temp_df = self.sheet.df
        new_labels = ["Campaign ID", "Phase", "Country",
                    "State", "Round", "Cost Category", "Line Item",	
                    "Quantity", "Unit Description", "Number of Days",
                    "Cost/Unit", "Total Cost", "Currency",
                    "Exchange Rate", "Total Costs (USD)"]
        
        # Check if the number of columns matches the number of new labels
        if len(temp_df.columns) != len(new_labels):
            raise ValueError("The number of columns in the DataFrame does not match the number of new labels.")
        
        # Relabel the columns
        temp_df.columns = new_labels
        self.sheet.df = temp_df


    def fill_currency_columns(self):

        #if int(self.sheet.exchange_rate) == 1:
        if True:  #workaround FIX LATER
            # sheet.currency = "USD"
            self.sheet.df["Exchange Rate"] = 1
            self.sheet.df['Currency'] = "USD"
        else:
            self.sheet.df["Exchange Rate"] = self.sheet.exchange_rate
            self.sheet.currency = "???"


        # accepts both str and int search_string
    def find_row_index(self, col_index, search_string):
        search_string = str(search_string)  # Ensure everything is treated as a string for regex search
        try:
            condition = self.sheet.df.iloc[:, col_index].astype(str).str.contains(search_string, na=False, regex=True)
            if condition.any():
                return int(condition.idxmax())  # Cast to standard Python integer
            else:
                return None
        except Exception as e:
            print(f"Error during search: {e}")
            return None


    def remove_category_row(self):
        pprint(self.sheet.df.iloc[:6])
        char_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
        try:
            column = self.sheet.df.iloc[:, 0]
            # Use list comprehension to find indices of rows with whole numbers (Category rows)
            removal_indices = [
                index for index, value in column.items() 
                if isinstance(value, (int, float)) and value % 1 == 0 or value in char_list
                ]
            # Drop the rows with these indices from the dataframe
            pprint(f'dropping rows:  {removal_indices}')
            self.sheet.df.drop(index=removal_indices, inplace=True)
        except Exception as e:
                print(f"Error during dropping rows: {e}")
        


    # Filters dataset in-place
    def extract_data(self):
            # Extract data starting from cell A8, drop columns after 10
            temp_df = self.sheet.df.iloc[7:] 
            #temp_df = self.sheet.df.iloc[8:] 
            temp_df = temp_df.iloc[:, :10]        
            
            print(temp_df.iloc[:0])
            # Drop rows where column C has no value
            self.sheet.df = temp_df.dropna(subset=[temp_df.columns[1]])

            # Drop OG row 12 by index
            removal_index = self.find_row_index(0, "Line")
            removal_index_2 = self.find_row_index(0, "No")  #Tsha update
            #self.sheet.df = self.sheet.df.drop([removal_index])
            #self.sheet.df = self.sheet.df.drop([removal_index_2])

            self.remove_category_row()

            # Replace values in column 1
            self.sheet.df.iloc[:, 0] = self.sheet.df.iloc[:, 0].apply(self.column_1_fix)
            temp_df = self.sheet.df

            #pprint(self.sheet.df.iloc[:3])

            # Drop rows where the third column has no value or invalid value
            mod_df = temp_df[temp_df.apply(self.apply_checks_and_modify, axis=1)]

            self.sheet.df = mod_df.reset_index(drop=True)
        
    def remove_dashes(self):
        # Define indices for the columns to clean
        columns_to_clean = [7, 10]  # Indices for 'Quantity' and 'Cost/Unit'

        # Replace '-' with NaN to unify the types of missing values
        for index in columns_to_clean:
            self.sheet.df.iloc[:, index] = self.sheet.df.iloc[:, index].replace('-', pd.NA)
            
        # Fill NaN values with 0 and correct data types
        for index in columns_to_clean:
            self.sheet.df.iloc[:, index] = self.sheet.df.iloc[:, index].fillna(0).infer_objects()

    def fix_days(self):
        # Index for 'Number of Days'
        days_index = 9
        # Replace '-' and '0' with NaN
        self.sheet.df.iloc[:, days_index] = self.sheet.df.iloc[:, days_index].replace(['-', '0'], pd.NA)
        # Fill NaN values with 1
        self.sheet.df.iloc[:, days_index] = self.sheet.df.iloc[:, days_index].fillna(1).infer_objects()



    # Main process loop
    def clean(self):
        self.extract_data()
        pprint(self.sheet.df.iloc[:4])
        self.add_columns()
        self.relabel_columns()
        self.fill_currency_columns()
        self.remove_dashes()
        self.fix_days()


