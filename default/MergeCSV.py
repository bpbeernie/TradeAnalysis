import os
import glob
import csv
from xlsxwriter.workbook import Workbook
from datetime import date
from default import Constants as const
import string
import math

def convertIndexToXLRow(index):
    return index + 1

def convertIndexToXLCol(index):
    return index + 1

monthlySummaries = []
stockList = []

folders = os.listdir(const.PATH)
folders.remove("Archive")

for folder in folders:
    filename = f'{folder}.xlsx'
    monthlySummaries.append(filename)
    workbook = Workbook(const.PATH + filename)
    summaryWorksheet = workbook.add_worksheet("Summary")
    summaryWorksheet.write(0, 0, "Stock")
    summaryWorksheet.write(0, 1, "Total")
    summaryRow = 1
    for csvfile in glob.glob(const.PATH + f'{folder}/*.csv'):
        stock = os.path.basename(csvfile).removesuffix('.csv')
        stockList.append(stock)
        worksheet = workbook.add_worksheet(name=stock)
        with open(csvfile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    try:
                        value = float(col)
                        worksheet.write_number(r, c, value)
                    except ValueError:
                        worksheet.write(r, c, col)
                if r != 0:
                    worksheet.write_formula(r, 3, f'=(B{convertIndexToXLRow(r)} -C{convertIndexToXLRow(r)})*{const.RISK}')
                    worksheet.write_formula(r, 4, f'=SUM(D2:D{convertIndexToXLRow(r)})')
        summaryWorksheet.write_formula(summaryRow, 0, f'=HYPERLINK("#{stock}!H5", "{stock}")')
        summaryWorksheet.write_formula(summaryRow, 1, f'={stock}!H5')
        summaryRow+=1
        worksheet.write(0, 0, "Date")
        worksheet.write(0, 3, "Subtotal")
        worksheet.write(0, 4, "Running Total")
        worksheet.write(1, 6, "Total Win")
        worksheet.write(2, 6, "Total Lose")
        worksheet.write(4, 6, "Total")
    
        worksheet.write_formula(1, 7, "=SUM(B2:B1000)")
        worksheet.write_formula(2, 7, "=SUM(C2:C1000)")
        worksheet.write_formula(4, 7, "=SUM(D2:D1000)")
        
    workbook.close()
    
today = date.today()
todayString = today.strftime("%Y%m%d")

numRows = len(set(stockList)) + 1
alphabetList = list(string.ascii_uppercase)

with Workbook(const.PATH + f'{todayString}-Summary.xlsx') as workbook:
    annualWorksheet = workbook.add_worksheet("AnnualSummary")
    
    numColumns = len(monthlySummaries) * 2
    
    for col in range(numColumns):
        for i in range(numRows):
            index = math.floor(col / 2)
            letter = alphabetList[col % 2]
            annualWorksheet.write_formula(i, col, f'=\'[{monthlySummaries[index]}]Summary\'!${letter}${i+1}')