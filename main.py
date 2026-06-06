import functions as fc

print("********EXCEL-PRICE-PROCESSOR***********")

print("Welcome to the Excel-Price-Processor!")
input("Please make sure that your excel file is in this folder (press enter to continue)")
input("Please make sure that your excel file is not opened anywhere on your computer currently")

decision = int(input(
    "Type in 0 if you want to create a fixed discount or 1 for percent discount: "))


file = input('What excel file should be edited? (<file>.xlsx) ')
sheet = input('What sheet should be edited? (ex. Sheet1) ')

min_row = int(input(
    'What is the minimum row of your data that has the prices? '))
max_row = int(input(
    'What is the maximum row of you data that has the prices? '))
old_col = input(
    'What is the letter of the old column of your data that has the original prices? ')
new_col = input(
    'What is the letter of the new column of your data that should have the new discounted prices? ')


old_col = fc.column_finder(old_col)
new_col = fc.column_finder(new_col)


# Create a new instance of dicount
discount_file = fc.Discount(filename=file, sheet_name=sheet, min_row=min_row,
                            max_row=max_row, old_column=old_col, new_column=new_col)

if decision == 0:
    fixed_discount = float(input(
        'By what amount do you want to discount the value? (type as a decimal) '))
    discount_file.fixed_discount(fixed_discount)


elif decision == 1:
    percent = float(input(
        'By what percent do you want to discount the value? (type as a decimal) '))
    discount_file.percentage_discount(percent)


print("Done!")
print("****************************************")
