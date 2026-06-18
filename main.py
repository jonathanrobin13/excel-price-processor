import excelFunctions as fc

print("\n********EXCEL-PRICE-PROCESSOR***********")
print("Welcome to the Excel-Price-Processor!")
input("\nPlease make sure that your excel file is in this folder (press enter to continue)")
input("Please make sure that your excel file is not opened anywhere on your computer currently")

decision = int(input(
    "\nType in 0 if you want to create a fixed discount or 1 for percent discount: "))


file = input('\nWhat excel file should be edited? (<file>.xlsx) ')
sheet = input('What sheet should be edited? (ex. Sheet1) ')

min_row = int(input(
    '\nWhat is the minimum row of your data that has the prices? '))
max_row = int(input(
    'What is the maximum row of you data that has the prices? '))
original_price_col = input(
    'What is the letter of the old column of your data that has the original prices? ')
discount_price_col = input(
    'What is the letter of the new column of your data that should have the new discounted prices? ')
product_name_col = input("What column contains the product names? ")

product_name_col = fc.column_finder(product_name_col)
original_price_col = fc.column_finder(original_price_col)
discount_price_col = fc.column_finder(discount_price_col)


# Create a new instance of dicount
discount_file = fc.Price(filename=file, sheet_name=sheet, min_row=min_row,
                         max_row=max_row, original_price_col=original_price_col, discount_price_col=discount_price_col, product_name_col=product_name_col)

if decision == 0:
    fixed_discount = float(input(
        'By what amount do you want to discount the value? (type as a decimal) '))
    discount_file.fixed_discount(fixed_discount)


elif decision == 1:
    percent = float(input(
        'By what percent do you want to discount the value? (type as a decimal) '))
    discount_file.percentage_discount(percent)


min_max_decision = input(
    "Type 0 if you want to find the min or press 1 to find the max or press enter to end the program ")

if min_max_decision == '0' or min_max_decision == '1':
    # min_max, min_max_product = discount_file.min_max_finder(min_max_decision)
    min_max_decision = int(min_max_decision)
    debug = discount_file.min_max_finder(min_max_decision)
    print(debug)
    # print(f"Product: {min_max_product} \nPrice: ${min_max}")

print("Done! You can now check your file for the updated prices!")
print("****************************************\n")
