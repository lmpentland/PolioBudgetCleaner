import os
import re

class DataForm:
    def __init__(self, df, sheet_name, filepath, filename):
        self.df = df
        self.phase = self.extract_phase_num(sheet_name)
        self.country = self.extract_cell_value('B2')
        self.state = self.extract_province(filename)
        self.campaign_id = self.extract_campaign_id(filepath)
        self.currency = None   # FIX
        self.exchange_rate= self.extract_cell_value('C6')
        #self.round_number = self.extract_cell_value('B3')
        self.round_number = self.phase
        #self.province = self.extract_province(filepath)


    def extract_phase_num(self, name):
        trimmed_name = name.rstrip()
        # Check if the last character is a digit
        if trimmed_name and trimmed_name[-1].isdigit():
            return int(trimmed_name[-1])
        # Check if the second-to-last character is a digit
        elif len(trimmed_name) > 1 and trimmed_name[-2].isdigit():
            return int(trimmed_name[-2])
        # Return None if neither the last nor the second-to-last character is a digit
        else:
            return None
     
    # Returns cell data using Excel coordinates
    def extract_cell_value(self, cell):            
        # Convert the cell reference to row and column indices
        col_index = ord(cell[0].upper()) - ord('A') 
        row_index = int(cell[1:]) - 2  # formerly 1.  WHY
        # Extract the cell value
        cell_value = self.df.iat[row_index, col_index]
        return cell_value
    

    def extract_campaign_id(self, directory):
        folder_name = os.path.basename(os.path.normpath(directory))
        # Extract the 4 digits after 'CE' from the folder name
        match = re.search(r'CE(\d{4})', folder_name)
        if match:
            digits = match.group(1)
            return int(digits)
        else:
            raise ValueError("The folder name does not contain 'CE' followed by four digits")
            return 0000
        

    def extract_province(self, filename):
        province = None
        try:
            filename_str = str(filename)

            # Remove "Copy of " from the filename if it exists
            filename_str = filename_str.replace("Copy of ", "")

            # Debug: Print the modified filename
            print("Modified filename:", filename_str)

            # Find the index of '_template'
            index = filename_str.find('_Template')
            print("Index of '_Template':", index)  # Debug print the index

            # Extract and return everything before '_template'
            if index != -1:
                print("Index is not -1, extracting substring")  # Debug statement
                return filename_str[:index]
            else:
                last_underscore_index = filename_str.rfind('_')
                # if no _template, get end of filename
                if last_underscore_index != -1:
                    province_str = filename_str[last_underscore_index + 1:]
                    period_index = province_str.find('.')
                    if period_index != -1:
                        province = province_str[:period_index]
            return province

        except Exception as e:
            print(f"Error during extraction: {e}")
            return None