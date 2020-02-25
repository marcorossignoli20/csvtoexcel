import xlsxwriter, csv, os

csv_path = input('Please enter the path of the CSV file: ')

if os.path.exists(csv_path): #if csv file path is correct
    workbook_name = input('Please enter the name of the EXCEL file: ')

    workbook = xlsxwriter.Workbook('./'+workbook_name) #creating workbook
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True}) #setting bold format
    money = workbook.add_format({'num_format': '€#,#0'}) #change the € symbol with the one you prefer
    incomes = 0
    expenses = 0

    with open(csv_path) as csv_file: #opening csv file
        csv_reader = csv.reader(csv_file, delimiter=',')
        row = 0
        
        for line in csv_reader:
            if row == 0:
                worksheet.write(row, 0, 'Date', bold)
                worksheet.write(row, 1, 'Category', bold)
                worksheet.write(row, 2, 'Expenses', bold)
                worksheet.write(row, 3, 'Incomes', bold)
                worksheet.write(row, 4, 'Notes', bold)
                row+=1
            else:
                for col in range(4):
                    if col==2:
                        num = float(line[col].lstrip())
                        if num < 0:
                            num = abs(num)
                            expenses += num
                            worksheet.write(row, col, num, money)
                        else:
                            num = abs(num)
                            incomes += num
                            worksheet.write(row, col+1, num, money)
                    elif col == 3:
                        worksheet.write(row, col+1, line[col])
                    else:
                        worksheet.write(row, col, line[col])
                row+=1

    worksheet.write(row, 0, 'Total amount:', bold)
    worksheet.write(row, 2, expenses, bold)
    worksheet.write(row, 3, incomes, bold)
    row+=2
    worksheet.write(row, 0, 'Totale missing: ', bold)
    worksheet.write(row, 1, expenses-incomes, bold)
    workbook.close()
else:
    print('** ERROR: CSV file not found **')
