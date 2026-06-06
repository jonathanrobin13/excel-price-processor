import openpyxl as xl
from openpyxl.chart import BarChart, Reference


class Discount:

    def __init__(self, filename, sheet_name, min_row, max_row, old_column, new_column):
        self.filename = filename
        self.sheet_name = sheet_name
        self.min_row = min_row
        self.max_row = max_row
        self.old_column = old_column
        self.new_column = new_column

    def percentage_discount(self, percent_discount):

        wb = xl.load_workbook(self.filename)
        sheet = wb[self.sheet_name]

        subtracted_price = 1.0 - float(percent_discount)

        for row in range(self.min_row, self.max_row + 1):
            cell = sheet.cell(row, self.old_column)
            corrected_price = float(cell.value) * subtracted_price
            corrected_price_cell = sheet.cell(row, self.new_column)
            corrected_price_cell.value = corrected_price

        wb.save(self.filename)

    def fixed_discount(self, fixed_discount):

        wb = xl.load_workbook(self.filename)
        sheet = wb[self.sheet_name]

        for row in range(self.min_row, self.max_row + 1):
            cell = sheet.cell(row, self.old_column)
            corrected_price = float(cell.value) - fixed_discount
            new_cell = sheet.cell(row, self.new_column)
            new_cell.value = corrected_price

        wb.save(self.filename)


# a min price and a max price finder for later


# Different Function for later
#    values = Reference(sheet,
#                       min_row=2,
#                       max_row=sheet.max_row,
#                       min_col=4,
#                       max_col=4)
#
#    chart = BarChart()
#    chart.add_data(values)
#    sheet.add_chart(chart, 'e2')
