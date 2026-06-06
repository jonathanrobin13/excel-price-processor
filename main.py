import functions as fc

file = input('What excel file should be edited? (<file>.xlsx) ')
sheet = input('What sheet should be edited? (ex. Sheet1) ')

print("********EXCEL-PRICE-PROCESSOR***********")
min_row = int(input(
    'What is the minimum row of your data that should be edited? (ex. 2) '))
max_row = int(input(
    'What is the maximum row of you data that should be edited? (ex. 4) '))
old_col = int(input(
    'What is the old column of your data that has the original prices? (ex. 2) '))
new_col = int(input(
    'What is the new column of your data that should have the new prices? (ex. 3) '))
percent = float(input(
    'By what percent do you want to discount the value? (type as a decimal) '))

fc.update_price(filename=file, sheet_name=sheet, min_row=min_row, max_row=max_row,
                old_column=old_col, new_column=new_col, percent_discount=percent)

print("****************************************")
