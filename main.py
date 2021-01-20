import tabula
# Read pdf into list of DataFrame
from PyPDF2 import PdfFileReader
import csv
import pandas as pd
import xlsxwriter
import math
import tabula

dates = []

# convert PDF into CSV file
tabula.convert_into("68403879_1610959324905.pdf", "output1.csv", output_format="csv", pages='all')
file = 'output1.csv'
df = pd.read_csv(file, skipfooter=5, engine = 'python')
print(df)
date = df["Date"]
# depo = df["Deposit Amt."].fillna(0)
for j in range(0, len(date)):
    # if date.iloc[j] != nan:
        dates.append(date.iloc[j])
        if type(dates[j]) == float:
            dates[j] = ""

####################### NARRATION TEST ################################
narration_final = []
# with open('output.csv', newline='') as csvfile:
#     data = csv.DictReader(csvfile)
#     for row in data:
##################### for narration ######################
narration = []
# narration_final = []

single_narration = df["Narration"]
for j in range(0, len(single_narration)):
    # if date.iloc[j] != nan:
        narration.append(single_narration.iloc[j])
        if type(narration[j]) == float:
            narration[j] = ""

# print(len(narration))
k=int()
for j in range(0, len(single_narration)):
    if "STATEMENT" in str(single_narration[j]):
        k = j
    else: k= len(narration)
#
for j in range(0, k):

    if j == k-1 or j == k-2:
        narration_final.append(narration[j])
        continue
    if j < k:
        if j != len(narration)+1:
            if len(dates[j]) != 0:
                if len(dates[j+1]) == 0:

                    if len(dates[j+2]) == 0 :
                        narration [j] = narration[j]+ narration[j+1] + narration[j+2]
                        narration_final.append(narration[j])
                        narration[j+1] = ""
                        narration[j+2] =""
                        narration_final.append(narration[j+1])
                        narration_final.append(narration[j+2])



                    else:
                        narration[j] = narration[j] + narration[j+1]
                        narration_final.append(narration[j])
                        narration[j+1] = ''
                        narration_final.append(narration[j+1])
                        pass
                else:
                    narration_final.append(narration[j])


############# for Chq./Ref.No. #########################

cheque_nos = []

cheque = df["Chq./Ref.No."]
for j in range(0, len(df.index)):
    # if date.iloc[j] != nan:
        cheque_nos.append(cheque.iloc[j])
        if type(cheque_nos[j]) == float:
            cheque_nos[j] = " "
# ############### for Value Dt ###############################
#
value_dates = []
v_date = df["Value Dt"]
for j in range(0, len(df.index)):
    # if date.iloc[j] != nan:
        value_dates.append(v_date.iloc[j])
        if type(value_dates[j]) == float:
            value_dates[j] = " "
############## for withdrawal #################
df_1 = pd.read_csv(file,skipfooter=len(narration) - 3)

wd_amount = []
wd = df["Withdrawal Amt."].str.replace(",","").astype(float)

for j in range(0, len(df.index)):
    # if date.iloc[j] != nan:
        wd_amount.append(wd.iloc[j])
        if math.isnan(wd_amount[j]):
            wd_amount[j] = " "
        if math.isnan(wd.iloc[j]) :
            wd.iloc[j] = " "
# ############### for deposit #################

depo_amount = []
x = df["Deposit Amt."].str.replace(",","").astype(float)
# depo = x.fillna(0)
for j in range(0, len(x)):
    # if date.iloc[j] != nan:
        depo_amount.append(x.iloc[j])
        if math.isnan(depo_amount[j]):
            depo_amount[j] = " "

# print(len(depo_amount))
############### for closing #################

closing_balance = []
closing = df["Closing Balance"].str.replace(",","").astype(float)
# print(closing)
for j in range(0, len(closing)):
    # if date.iloc[j] != nan:
        closing_balance.append(closing.iloc[j])
        if math.isnan(closing_balance[j]):
            closing_balance[j] = " "
        if math.isnan(closing.iloc[j]) :
            closing.iloc[j] = " "
#
date_column = []
narration_column = []
cheque_column = []
value_column = []
wd_column = []
depo_column = []
closing_column = []
for j in range(0, len(dates)):
    if j != len(dates)-1:
        if len(dates[j]) != 0:
            date_column.append(dates[j])
            narration_column.append(narration[j])
            cheque_column.append(cheque_nos[j])
            closing_column.append(closing_balance[j])
            wd_column.append(wd_amount[j])
            value_column.append(value_dates[j])
            depo_column.append(depo_amount[j])

outworkbook = xlsxwriter.Workbook("bank_statment_to_excel_6.xlsx")
outsheet = outworkbook.add_worksheet()

outsheet.write("A1", "Date")
outsheet.write("B1", "Narration")
outsheet.write("C1", "Chq./Ref.No.")
outsheet.write("D1", "Value Dt")
outsheet.write("E1", "Withdrawal Amt.")
outsheet.write("F1", "Deposit Amt.")
outsheet.write("G1", "Closing balance.")
#
#
#
#
for j in range (0, len(date_column)):
        outsheet.write(j+1, 0, date_column[j])
        outsheet.write(j+1, 1, narration_column[j])
        outsheet.write(j+1,2,cheque_column[j])
        outsheet.write(j+1,3,value_column[j])
        outsheet.write(j+1,4,wd_column[j])
        outsheet.write(j+1,5,depo_column[j])
        outsheet.write(j+1,6,closing_column[j])
        j = j + 1


outworkbook.close()
# #
#

#
# import win32com.client
#
# filename = 'C:\\Users\\91898\\.PyCharmCE2019.3\\config\\scratches\\python compliance\\bank_statment_to_excel_3.xlsx'
# sheetname = 'Sheet1'
# xl = win32com.client.DispatchEx('Excel.Application')
# wb = xl.Workbooks.Open(Filename=filename)
# ws = wb.Sheets(sheetname)
#
# begrow = 1
# endrow = ws.UsedRange.Rows.Count
# for row in range(begrow,endrow+1): # just an example
#   if ws.Range('A{}'.format(row)).Value is None:
#     ws.Range('A{}'.format(row)).EntireRow.Delete(Shift=-4162) # shift up
#
# wb.Save()
# wb.Close()
# xl.Quit()
