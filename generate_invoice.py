from openpyxl import load_workbook
from openpyxl import drawing 
from pprint import pprint
import pandas as pd
from num2words import num2words
import os
import pickle as pkl
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import tkinter.messagebox

def generate_UT(sheet, billing, cnt):
    billing[3] = pd.to_datetime(billing[3], format='%d/%m/%Y')
    #modify header
    sheet["B4"] = "INV/UT/GKN/"+str(billing[3].year)+"/"+billing[3].strftime('%m')+"/{:04d}".format(cnt) # inv number
    sheet["B6"] = billing[1]    # Tenant
    sheet["B7"] = billing[2]    # Alamat Tenant
    sheet["AF6"] = billing[3]   # Invoice Date
    sheet["AF7"] = (billing[3] + pd.DateOffset(days=15)).strftime('%d/%m/%Y')   # Due Date
    sheet["E11"] = billing[0]   # Unit
    sheet["Y10"] = billing[3].month_name()+" "+str(billing[3].year)    # Tagihan Bulan
    sheet["AD10"] = billing[25]    # Jumlah Tagihan
    # sheet["AD10"] = billing[40]   # VA

    #modify electricity
    sheet["L16"] = (billing[3] + pd.DateOffset(days=-9) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode Start
    sheet["S16"] = (billing[3] + pd.DateOffset(days=-9)).strftime('%d/%m/%Y')   # Periode End
    sheet["K17"] = billing[9]   # Consumption Start
    sheet["Q17"] = billing[10]   # Consumption End
    sheet["K18"] = billing[8]   # Meter Factor
    sheet["K19"] = billing[11]   # Usage
    sheet["K20"] = billing[7]   # Minimum Charge
    sheet["H21"] = billing[5] + " kVa"  # Daya kVa
    sheet['K22'] = "2,4%" # PPJU
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
        sheet["X20"] = ""   # X
        sheet["AA20"] = ""   # rates
        sheet["AD20"] = ""   # IDR
        sheet["AE20"] = ""   # Amount Minimum
        sheet["AE19"] = billing[12]   # Amount Usage
    sheet["AE22"] = billing[13]   # Ppju
    sheet["AE23"] = billing[14]   # Total Amount

    # modify water
    sheet["L24"] = (billing[3] + pd.DateOffset(days=-2) + pd.DateOffset(months=-1)).strftime('%d/%m/%Y')   # Periode Start
    sheet["S24"] = (billing[3] + pd.DateOffset(days=-2)).strftime('%d/%m/%Y')   # Periode End
    sheet["K25"] = billing[16]   # Consumption Start
    sheet["Q25"] = billing[17]   # Consumption End
    sheet["K26"] = billing[18]   # Usage
    sheet["K27"] = billing[20]   # Fix Charge
    sheet["K28"] = billing[21]   # Maintenance
    sheet["AA26"] = billing[15]   # Rates
    sheet["AE26"] = billing[19]   # Amount Usage / Usage Charge
    sheet["AE27"] = billing[20]   # Amount Fixed Charge
    sheet["AE28"] = billing[21]   # Amount Maintenance
    sheet["AE29"] = billing[22]   # Total Amount

    #modify admin
    sheet["AE31"] = billing[23]   # Common Area
    # sheet["AE31"] = billing[42] if isinstance(billing[42], str) else "0,00"  # Late Charges
    sheet["AE32"] = "0,00"  # Late Charges
    sheet["AE33"] = billing[24]   # Materai
    sheet["AE34"] = billing[25]   # Total Billing
    
    sheet["G37"] = (num2words(billing[25].replace(")","").replace("(","")[:-3].replace(".",""))+" rupiah").title()   # Terbilang

def generate_SCSF(sheet, billing, cnt):
    billing[3] = pd.to_datetime(billing[3], format='%d/%m/%Y')
    #modify header
    sheet["B4"] = "INV/SC-SF/GKN/"+str(billing[3].year)+"/"+billing[3].strftime('%m')+"/{:04d}".format(cnt) # inv number
    sheet["B6"] = billing[1]    # Tenant
    sheet["B7"] = billing[2]    # Alamat Tenant
    sheet["AF6"] = billing[3]   # Invoice Date
    sheet["AF7"] = (billing[3] + pd.DateOffset(days=15)).strftime('%d/%m/%Y')   # Due Date
    sheet["E11"] = billing[0]   # Unit
    sheet["Y10"] = billing[3].month_name()+" "+str(billing[3].year)    # Tagihan Bulan
    sheet["AD10"] = billing[38]    # Jumlah Tagihan
    # sheet["AD10"] = billing[41]   # VA SC-SF

    # set date start and end
    dt_start = (billing[3]).strftime('%d/%m/%Y')
    dt_end = (billing[3] + pd.tseries.offsets.MonthEnd(0)).strftime('%d/%m/%Y')
    
    #modify SC
    sheet["K16"] = dt_start + " - " + dt_end # Periode
    sheet["K18"] = billing[27]   # Rate SC
    sheet["S18"] = billing[28]   # Area SC
    sheet["K19"] = billing[30]   # PPH
    sheet["AE17"] = billing[29]    # Net Amount SC
    sheet["AE19"] = billing[30]    # PPH
    sheet["AE20"] = billing[31]    # Total SC

    #modify SF
    sheet["K21"] = dt_start + " - " + dt_end # Periode
    sheet["K23"] = billing[32]   # Rate SF
    sheet["S23"] = billing[28]   # Area SF
    sheet["K24"] = billing[34]   # PPH
    sheet["AE22"] = billing[33]    # Net Amount SF
    sheet["AE24"] = billing[34]    # PPH
    sheet["AE25"] = billing[35]    # Total SF

    #modify admin
    sheet["AE27"] = billing[36]   # Admin Fee
    # sheet["AE28"] = billing[43] if isinstance(billing[43], str) else "0,00"   # Late Charges
    sheet["AE28"] = "0,00"   # Late Charges
    sheet["AE29"] = billing[37]   # Materai
    sheet["AE30"] = billing[38]   # Total Billing

    sheet["G33"] = (num2words(billing[38].replace(")","").replace("(","")[:-3].replace(".",""))+" rupiah").title()   # Terbilang

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()

def main():
    

    print("DONE!")
    tkinter.messagebox.showinfo("Done", "Billing Generated")

if __name__ == '__main__':
    main()