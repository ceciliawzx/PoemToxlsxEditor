import pandas as pd

"""
This is a program to append content to an existed excel file with specific sheet name.
It will automatically calculate the start row by looking at the max_row of cur sheet.
"""

# The example file name is test.xlsx here, should be replaced by the real file.
writer = pd.ExcelWriter('test.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
# '诗词' is an example sheet name.
start_row = writer.sheets['诗词'].max_row

# df is an example DataFrame.
df = pd.DataFrame(['a', 'b', 'c'])
df.index += start_row
df.to_excel(writer, sheet_name='诗词', startrow=start_row, header=False)
writer.close()
