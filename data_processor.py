import pandas as pd
import re,os

class Sheet:
    def __init__(self, excel_path : str, input : str, output : str):
        self.excel_path = excel_path
        self.input_filename = input
        self.output_filename = output
    
    def process_sheets(self):
        try:
            # Read the excel file and create the df
            sheet_df = pd.read_excel(
                os.path.join(self.excel_path, self.input_filename), 
                sheet_name='All'    
            )
            # insert pincode column
            insertion_index = sheet_df.columns.get_loc('Products')
            # insert 3 address columns
            additional_columns = ["Pincode", "Phone", "Address 1", "Address 2", "Address 3"]
            for column in additional_columns[::-1]:
                sheet_df.insert(insertion_index, column, value = None)
                insertion_index = sheet_df.columns.get_loc(column)
            
            for address in sheet_df['Address']:
                if type(address) == str:
                    print(address)
                    # pincode 
                    pincode_pattern = r"[1-9][0-9]{5}"
                    pincode_match = re.findall(pincode_pattern, address)
                    # phone number
                    phone_pattern = r"[6-9]\d{9}"
                    phone_number_match = re.findall(phone_pattern,address)
                    
                    if pincode_match and phone_number_match:
                        print(pincode_match,phone_number_match)
                print("-"*20)
            print(sheet_df)
            
            # save output
            sheet_df.to_excel(
                os.path.join(self.excel_path,self.output_filename),
                index='False'             
            )
        except Exception as e:
            print(e)


sheet_inst = Sheet(
    excel_path='/home/hari/Desktop/Postal Direct',
    input='Direct parcel.xlsx',
    output='out.xlsx'
)
sheet_inst.process_sheets()



