from openpyxl import load_workbook
from openpyxl import drawing 
from pprint import pprint
import pandas as pd
from num2words import num2words
import glob
import os

def generate_UT(sheet, billing, cnt):
    billing[3] = pd.to_datetime(billing[3], format='%d/%m/%Y')
    #modify header
    sheet["B4"] = "INV/UT/DRK/"+str(billing[3].year)+"/"+billing[3].strftime('%m')+"/{:04d}".format(cnt) # inv number
    sheet["B5"] = billing[1]    # Tenant
    sheet["B6"] = billing[2]    # Alamat Tenant
    sheet["AF5"] = billing[3]   # Invoice Date
    sheet["AF6"] = (billing[3] + pd.DateOffset(days=19)).strftime('%d/%m/%Y')   # Due Date
    sheet["E11"] = billing[0]   # Unit
    sheet["Q10"] = billing[3].month_name()+" "+str(billing[3].year)    # Tagihan Bulan
    sheet["Y10"] = billing[25]    # Jumlah Tagihan
    sheet["AD10"] = billing[40]   # VA

    #modify electricity
    sheet["L16"] = (billing[3] + pd.DateOffset(days=19) + pd.DateOffset(months=-2)).strftime('%d/%m/%Y')   # Periode Start
    sheet["S16"] = (billing[3] + pd.DateOffset(days=19) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode End
    sheet["K17"] = billing[9]   # Consumption Start
    sheet["Q17"] = billing[10]   # Consumption End
    sheet["K18"] = billing[8]   # Meter Factor
    sheet["K19"] = billing[11]   # Usage
    sheet["K20"] = billing[7]   # Minimum Charge
    sheet["H21"] = billing[5] + " kVa"  # Daya kVa
    sheet["AA19"] = billing[6]    # Rates 1
    sheet["AA20"] = billing[6]    # Rates 2
    if isinstance(billing[7], float):
        min_chrg = 0
        usg_chrg = billing[11].replace(".","").replace(",",".").replace(")","").replace("(","")
    else:
        min_chrg = billing[7].replace(".","").replace(",",".").replace(")","").replace("(","")
        usg_chrg = billing[11].replace(".","").replace(",",".").replace(")","").replace("(","")
    if float(usg_chrg)<float(min_chrg):
        sheet["X19"] = ""   # X
        sheet["AA19"] = ""   # rates
        sheet["AD19"] = ""   # IDR
        sheet["AE19"] = ""   # Amount Usage
        sheet["AE20"] = billing[12]   # Amount Minimum
    else:
        sheet["AE19"] = billing[12]   # Amount Usage
        sheet["AE20"] = ""   # Amount Minimum
        sheet["X20"] = ""   # X
        sheet["AA20"] = ""   # rates
        sheet["AD20"] = ""   # IDR
        sheet["AE20"] = ""   # Amount Usage
    sheet["AE22"] = billing[13]   # Ppju
    sheet["AE23"] = billing[14]   # Total Amount

    #modify water
    sheet["L24"] = (billing[3] + pd.DateOffset(days=19) + pd.DateOffset(months=-2)).strftime('%d/%m/%Y')   # Periode Start
    sheet["S24"] = (billing[3] + pd.DateOffset(days=19) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode End
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
    
    sheet["G36"] = (num2words(billing[25].replace(")","").replace("(","")[:-3].replace(".",""))+" rupiah").title()   # Terbilang

def generate_SCSF(sheet, billing, cnt):
    billing[3] = pd.to_datetime(billing[3], format='%d/%m/%Y')
    adjust = 22
    #modify header
    sheet["B4"] = "INV/SC-SF/DRK/"+str(billing[3].year)+"/"+billing[3].strftime('%m')+"/{:04d}".format(cnt) # inv number
    sheet["B5"] = billing[1]    # Tenant
    sheet["B6"] = billing[2]    # Alamat Tenant
    sheet["AF5"] = billing[3]   # Invoice Date
    sheet["AF6"] = (billing[3] + pd.DateOffset(days=19)).strftime('%d/%m/%Y')   # Due Date
    sheet["E11"] = billing[0]   # Unit
    sheet["Q10"] = billing[3].month_name()+" "+str(billing[3].year)    # Tagihan Bulan
    sheet["Y10"] = billing[adjust+16]    # Jumlah Tagihan
    sheet["AD10"] = billing[40]   # VA

    # set date start and end
    dt_start = (billing[3]).strftime('%d/%m/%Y')
    dt_end = (billing[3] + pd.tseries.offsets.MonthEnd(0)).strftime('%d/%m/%Y')
    
    #modify SC
    sheet["K16"] = dt_start + " - " + dt_end # Periode
    sheet["K18"] = billing[adjust+5]   # Rate SC
    sheet["S18"] = billing[adjust+6]   # Area SC
    sheet["K19"] = billing[adjust+8]   # Vat SC
    sheet["AE17"] = billing[adjust+7]    # Net Amount SC
    sheet["AE19"] = billing[adjust+8]    # Vat Amount SC
    sheet["AE20"] = billing[adjust+9]    # Total SC

    #modify SF
    sheet["K21"] = dt_start + " - " + dt_end # Periode
    sheet["K23"] = billing[adjust+10]   # Rate SF
    sheet["S23"] = billing[adjust+6]   # Area SF
    sheet["K24"] = billing[adjust+12]   # Vat SF
    sheet["AE22"] = billing[adjust+11]    # Net Amount SF
    sheet["AE24"] = billing[adjust+12]    # Vat Amount SF
    sheet["AE25"] = billing[adjust+13]    # Total SF

    #modify admin
    sheet["AE27"] = billing[adjust+14]   # Materai
    sheet["AE28"] = billing[adjust+15]   # Admin Fee
    sheet["AE29"] = billing[adjust+16]   # Total Billing

    sheet["G32"] = (num2words(billing[adjust+16].replace(")","").replace("(","")[:-3].replace(".",""))+" rupiah").title()   # Terbilang

# get all csv files
path =  os.path.dirname(os.path.realpath(__file__))
folder_name = "csv"
read_files = glob.glob(path+"/" + folder_name + "/*.csv")

# loop all csv files
for read_file in read_files:
    # read csv
    data = pd.read_csv(read_file, dayfirst=True, parse_dates=[4], dtype={'VA': str})

    cnt_ut = 1
    cnt_scsf = 1
    print(data.values[0][3])

    # get csv file name
    file_name = read_file.split("\\")[-1].split("/")[-1][:-4]

    # open excel template
    try:
        workbook = load_workbook(filename= path + "/template/template.xlsx")
    except:
        print("File 'template/template.xlsx' doesn't exist")
        exit()
    banyak = len(data.values)
    for billing in data.values: 
        print(banyak)
        banyak = banyak-1
        if billing[25] != "0,00":
            workbook.copy_worksheet(workbook["UT"]).title = billing[0]+ " UT"
            sheet = workbook[billing[0]+ " UT"]
            generate_UT(sheet, billing, cnt_ut)
            cnt_ut += 1

        if billing[38] != "0,00":   
            workbook.copy_worksheet(workbook["SC-SF"]).title = billing[0]+ " SC-SF"
            sheet = workbook[billing[0]+ " SC-SF"]
            generate_SCSF(sheet, billing, cnt_scsf)
            cnt_scsf +=1
            
    del workbook["UT"]
    del workbook["SC-SF"]
    for sheet in workbook:
        img = drawing.image.Image(path + "/" + 'img/logo.png') 
        ttd = drawing.image.Image(path + "/" + 'img/ttd.png')
        sheet.add_image(img, 'B2')
        if "UT" in str(sheet):
            sheet.add_image(ttd, 'AC43')
        if "SC-SF" in str(sheet):
            sheet.add_image(ttd, 'AC39')

    #save the file
    # checking if the directory exist or not.
    if not os.path.exists(path + "/result/"):
        # then create it.
        os.makedirs(path + "/result/")

    workbook.save(filename= path + "/result/" + file_name + ".xlsx")

print("DONE!")