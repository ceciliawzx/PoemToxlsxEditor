import pandas as pd

"""
This is a program to append content to an existed excel file with specific sheet name.
It will automatically calculate the start row by looking at the max_row of cur sheet.
"""

def AskInput():
    stop = False
    append_list = []
    while not stop:
        new_poem: str = input("Please input new poem and author, divided by space.\n")
        alist = new_poem.split()
        append_list.append(alist)
        stop = input("Add more?\n0 for yes, 1 for no\n") == "1"
    return append_list

def Append():
    writer = pd.ExcelWriter('result.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
    # '诗词' is an example sheet name.
    start_row = writer.sheets['诗词'].max_row
    the_list = AskInput()
    # df is an example DataFrame.
    df = pd.DataFrame(the_list)
    df.index += start_row
    df.to_excel(writer, sheet_name='诗词', startrow=start_row, header=False)
    writer.close()

Append()