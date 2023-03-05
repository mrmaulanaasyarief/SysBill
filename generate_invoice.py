from openpyxl import load_workbook
from openpyxl import drawing 
from pprint import pprint
import pandas as pd
import pdfkit
from terbilang import Terbilang
from num2words import num2words
import glob
import os

# get all csv files
path =  os.getcwd()
folder_name = "csv"
read_files = glob.glob(path+"\\" + folder_name + "\\*.csv")

# loop all csv files
for read_file in read_files:
    # read csv
    data = pd.read_csv(read_file, dayfirst=True, parse_dates=[4])

    # get csv file name
    file_name = read_file.split("\\")[-1][:-4]

    cnt = 1
    
    if(data.values[0][3]=="UT"): 
        # open excel template
        try:
            workbook = load_workbook(filename="template/template.xlsx")
        except:
            print("File 'template/template.xlsx' doesn't exist")
            exit()

        for billing in data.values:
            workbook.copy_worksheet(workbook["UT"]).title = "(UT) "+billing[0]

            sheet = workbook["(UT) "+billing[0]]

            #modify header
            sheet["B4"] = "INV/UT/DRK/"+str(billing[4].year)+"/"+billing[4].strftime('%m')+"/{:04d}".format(cnt) # inv number
            sheet["B5"] = billing[1]    # Tenant
            sheet["B6"] = billing[2]    # Alamat Tenant
            sheet["AF5"] = billing[4]   # Invoice Date
            sheet["AF6"] = (billing[4] + pd.DateOffset(days=19)).strftime('%d/%m/%Y')   # Due Date
            sheet["E11"] = billing[0]   # Unit
            sheet["Q10"] = billing[4].month_name()+" "+str(billing[4].year)    # Tagihan Bulan
            sheet["Y10"] = billing[25]    # Jumlah Tagihan
            sheet["AD10"] = billing[26]   # VA

            #modify electricity
            sheet["L16"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-2)).strftime('%d/%m/%Y')   # Periode Start
            sheet["S16"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode End
            sheet["K17"] = billing[9]   # Consumption Start
            sheet["Q17"] = billing[10]   # Consumption End
            sheet["K18"] = billing[8]   # Meter Factor
            sheet["K19"] = billing[11]   # Usage
            sheet["K20"] = billing[7]   # Minimum Charge
            sheet["H21"] = billing[5]   # Daya kVa
            sheet["AA19"] = billing[6]    # Rates 1
            sheet["AA20"] = billing[6]    # Rates 2
            if isinstance(billing[7], float):
                min_chrg = 0
                usg_chrg = billing[11].replace(".","").replace(",",".")
            else:
                min_chrg = billing[7].replace(".","").replace(",",".")
                usg_chrg = billing[11].replace(".","").replace(",",".")
            if float(usg_chrg)<float(min_chrg):
                sheet["AE19"] = "0,00"   # Amount Usage
                sheet["AE21"] = billing[12]   # Amount Minimum
            else:
                sheet["AE19"] = billing[12]   # Amount Usage
                sheet["AE21"] = "0,00"   # Amount Minimum
            sheet["AE22"] = billing[13]   # Ppju
            sheet["AE23"] = billing[14]   # Total Amount

            #modify water
            sheet["L24"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-2)).strftime('%d/%m/%Y')   # Periode Start
            sheet["S24"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode End
            sheet["K25"] = billing[16]   # Consumption Start
            sheet["Q25"] = billing[17]   # Consumption End
            sheet["K26"] = billing[18]   # Usage
            sheet["K27"] = billing[20]   # Fix Charge
            sheet["K28"] = billing[21]   # Maintenance
            sheet["AA26"] = billing[15]    # Rates
            sheet["AE26"] = billing[19]   # Amount Usage
            sheet["AE27"] = billing[20]   # Amount Fixed Charge
            sheet["AE28"] = billing[21]   # Amount Maintenance
            sheet["AE29"] = billing[22]   # Total Amount

            #modify admin
            sheet["AE31"] = billing[23]   # Materaui
            sheet["AE32"] = billing[24]   # Admin Fee
            sheet["AE33"] = billing[25]   # Total Billing

            sheet["D35"] = (num2words(billing[25][:-3].replace(".",""))+" rupiah").title()   # Terbilang
            
            print(len(data.values)-cnt)
            cnt += 1
    # elif(data.values[0][3]="SC-SF"):
    # exit()

    # del workbook["UT"]
    # del workbook["SC-SF"]
    # for sheet in workbook:
    #         img = drawing.image.Image('img/logo.png') 
    #         ttd = drawing.image.Image('img/ttd.png')
    #         sheet.add_image(img, 'B2')
    #         sheet.add_image(ttd, 'AC43')
    # #save the file
    # workbook.save(filename="result/Maret 2023 - UT.xlsx")


data = pd.read_csv('csv/Maret 2023 - UT.csv', dayfirst=True, parse_dates=[4])

UT = 1
workbook = load_workbook(filename="template/template.xlsx")
for billing in data.values:
    workbook.copy_worksheet(workbook["UT"]).title = "(UT) "+billing[0]

    sheet = workbook["(UT) "+billing[0]]

    #modify header
    sheet["B4"] = "INV/UT/DRK/"+str(billing[4].year)+"/"+billing[4].strftime('%m')+"/{:04d}".format(UT) # inv number
    sheet["B5"] = billing[1]    # Tenant
    sheet["B6"] = billing[2]    # Alamat Tenant
    sheet["AF5"] = billing[4]   # Invoice Date
    sheet["AF6"] = (billing[4] + pd.DateOffset(days=19)).strftime('%d/%m/%Y')   # Due Date
    sheet["E11"] = billing[0]   # Unit
    sheet["Q10"] = billing[4].month_name()+" "+str(billing[4].year)    # Tagihan Bulan
    sheet["Y10"] = billing[25]    # Jumlah Tagihan
    sheet["AD10"] = billing[26]   # VA

    #modify electricity
    sheet["L16"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-2)).strftime('%d/%m/%Y')   # Periode Start
    sheet["S16"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode End
    sheet["K17"] = billing[9]   # Consumption Start
    sheet["Q17"] = billing[10]   # Consumption End
    sheet["K18"] = billing[8]   # Meter Factor
    sheet["K19"] = billing[11]   # Usage
    sheet["K20"] = billing[7]   # Minimum Charge
    sheet["H21"] = billing[5]   # Daya kVa
    sheet["AA19"] = billing[6]    # Rates 1
    sheet["AA20"] = billing[6]    # Rates 2
    if isinstance(billing[7], float):
        min_chrg = 0
        usg_chrg = billing[11].replace(".","").replace(",",".")
    else:
        min_chrg = billing[7].replace(".","").replace(",",".")
        usg_chrg = billing[11].replace(".","").replace(",",".")
    if float(usg_chrg)<float(min_chrg):
        sheet["AE19"] = "0,00"   # Amount Usage
        sheet["AE21"] = billing[12]   # Amount Minimum
    else:
        sheet["AE19"] = billing[12]   # Amount Usage
        sheet["AE21"] = "0,00"   # Amount Minimum
    sheet["AE22"] = billing[13]   # Ppju
    sheet["AE23"] = billing[14]   # Total Amount

    #modify water
    sheet["L24"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-2)).strftime('%d/%m/%Y')   # Periode Start
    sheet["S24"] = (billing[4] + pd.DateOffset(days=19) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode End
    sheet["K25"] = billing[16]   # Consumption Start
    sheet["Q25"] = billing[17]   # Consumption End
    sheet["K26"] = billing[18]   # Usage
    sheet["K27"] = billing[20]   # Fix Charge
    sheet["K28"] = billing[21]   # Maintenance
    sheet["AA26"] = billing[15]    # Rates
    sheet["AE26"] = billing[19]   # Amount Usage
    sheet["AE27"] = billing[20]   # Amount Fixed Charge
    sheet["AE28"] = billing[21]   # Amount Maintenance
    sheet["AE29"] = billing[22]   # Total Amount

    #modify admin
    sheet["AE31"] = billing[23]   # Materaui
    sheet["AE32"] = billing[24]   # Admin Fee
    sheet["AE33"] = billing[25]   # Total Billing

    sheet["D35"] = (num2words(billing[25][:-3].replace(".",""))+" rupiah").title()   # Terbilang
    
    print(len(data.values)-UT)
    UT += 1
del workbook["UT"]
del workbook["SC-SF"]
for sheet in workbook:
        img = drawing.image.Image('img/logo.png') 
        ttd = drawing.image.Image('img/ttd.png')
        sheet.add_image(img, 'B2')
        sheet.add_image(ttd, 'AC43')
#save the file
workbook.save(filename="result/Maret 2023 - UT.xlsx")
