import xlsxwriter, csv

workbook = xlsxwriter.Workbook('./Spese01.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
money = workbook.add_format({'num_format': '€#,#0'})
entrate = 0
uscite = 0

with open('./spese.csv') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    row = 0
    for line in csv_reader:
        if row == 0:
            worksheet.write(row, 0, 'Data', bold)
            worksheet.write(row, 1, 'Categoria', bold)
            worksheet.write(row, 2, 'Uscite', bold)
            worksheet.write(row, 3, 'Entrate', bold)
            worksheet.write(row, 4, 'Note', bold)
            row+=1
        else:
            if(line[1].lstrip() == 'Gifts'):
                pass
            else:
                for col in range(4):
                    if col==2:
                        num = float(line[col].lstrip())
                        if num < 0:
                            num = abs(num)
                            uscite += num
                            worksheet.write(row, col, num, money)
                        else:
                            num = abs(num)
                            entrate += num
                            worksheet.write(row, col+1, num, money)
                    elif col == 3:
                        worksheet.write(row, col+1, line[col])
                    else:
                        worksheet.write(row, col, line[col])
                row+=1
worksheet.write(row, 0, 'Totale:', bold)
worksheet.write(row, 2, uscite, bold)
worksheet.write(row, 3, entrate, bold)
row+=2
worksheet.write(row, 0, 'Totale mancante: ', bold)
worksheet.write(row, 1, uscite-entrate, bold)
workbook.close()
        
