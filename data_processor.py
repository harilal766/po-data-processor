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
                sheet_name='PO Data'    
            )
            # insert pincode column
            insertion_index = sheet_df.columns.get_loc('Products')
            # insert 3 address columns
            additional_columns = [
                "Barcode", "Name", "City", "Pincode", "Phone", "Address 1", "Address 2", "Address 3"
            ]
            for column in additional_columns[::-1]:
                sheet_df.insert(insertion_index, column, value = None)
                insertion_index = sheet_df.columns.get_loc(column)
                
            split_syntax = r"\n|,\n|,"
            patterns = {
                "Name" : split_syntax,
                "Pincode" : r"([1-9]\d{4,5}|[1-9]\d{2}\s+\d{3}){1}",
                "Phone" : r"([6-9]\d{9}|[6-9]\d{4} \d{5})"
            }
            for idx, row in sheet_df.iterrows():
                address = row["Address"]
                if type(address) == str:
                    for column_name, pattern_syntax in patterns.items():
                        if column_name != "Name":
                            pattern_matches = re.findall(pattern_syntax, address,re.IGNORECASE)
                            pattern_match = re.sub(r"\s", "", pattern_matches[0])
                        else:
                            pattern_matches = re.split(pattern_syntax, address,re.IGNORECASE)
                            pattern_match = pattern_matches[0]
                        
                        if pattern_matches and type(pattern_matches) == list:
                            # assign value to the column
                            
                            #print(f"{column_name} : {pattern_match}")
                            sheet_df.loc[idx, column_name] = pattern_match
                            # remove the match from the address cell 
                            things_to_remove = (pattern_match, r"Pin|PIN|pin|Mob|MOB|mob|Phone|phone|/.+|/,")
                            for thing in things_to_remove:
                                sheet_df.loc[idx, "Address"] = re.sub(
                                    thing, "", sheet_df.loc[idx, "Address"],re.IGNORECASE
                                )
                            
                    # split address in to three
                    address_lines = re.split(
                        split_syntax,sheet_df.loc[idx, "Address"]
                    )
                    
                    leftover_address = []
                    for line in address_lines:
                        if line not in ("", " ",","):
                            leftover_address.append(line)
                    
                    leftover_address_length = len(leftover_address)
                    dividing = int(leftover_address_length/3)
                    index_count = 0
                    address_redistribution_columns = additional_columns[-3::]
                    for col in address_redistribution_columns:
                        starting = index_count
                        index_count += dividing
                        sheet_df.loc[idx,col] = ','.join(leftover_address[starting:index_count])
            # save output
            print(sheet_df[["Address", "Pincode","Phone"]])
            sheet_df.to_excel(
                os.path.join(self.excel_path,self.output_filename),
                index=False         
            )
        except Exception as e:
            print(e)


sheet_inst = Sheet(
    excel_path='/home/hari/Desktop/Postal Direct',
    input='Direct parcel.xlsx',
    output='out.xlsx'
)
sheet_inst.process_sheets()



