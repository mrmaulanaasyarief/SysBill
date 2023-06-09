import os
from pikepdf import Pdf
from pprint import pprint
import pickle as pkl
import glob
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import tkinter.messagebox

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

def set_file2pages(pkl_file):
    # arr = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 1, 2, 2, 2, 1, 2, 1, 1]
    arr = {}
    with open(pkl_file, 'rb') as f:
        arr = pkl.load(f)

    file2pages = {}

    n = 0
    idx = 0
    for i in arr["seq"]:
        a = n
        b = n+i
        n = b
        file2pages[idx] = [a,b]
        idx += 1

    return file2pages, arr["unit"], arr["type"]

def main():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    while True:
        filename = askopenfilename(title="Select PDF File") # show an "Open" dialog box and return the path to the selected file
        if(filename==""):
            if tkinter.messagebox.askretrycancel("Error",  "No PDF file selected"):
                pass
            else:
                exit()
        else:
            if(filename.lower().endswith(".pdf")):
                break
            else:
                if tkinter.messagebox.askretrycancel("Error",  "Selected file must be in PDF format"):
                    pass
                else:
                    exit()

    # the target PDF document to split
    # filename = "result/pdf/Final Invoice April 23.pdf"

    print("Preparing...")
    # get all csv files
    path =  os.path.dirname(os.path.realpath(__file__))
    # get csv file name
    file_name = filename.split("\\")[-1].split("/")[-1][:-4]
    folder_name = "result/"+file_name+"/"
    # read_files = glob.glob(path+"/" + folder_name + "/*.pdf")
    read_files = [filename]
    
    # a dictionary mapping PDF file to original PDF's page range
    pkl_file = filename.replace("pdf","pkl")
    file2pages, unit, type = set_file2pages(pkl_file)

    #save the file
    # checking if the directory exist or not.
    if not os.path.exists(path +"/" + folder_name + "/pdf splitted/"):
        # then create it.
        os.makedirs(path + "/" + folder_name + "/pdf splitted/")

    # loop all pdf files
    for read_file in read_files:
        # load the PDF file
        pdf = Pdf.open(read_file)

        # make the new splitted PDF files
        new_pdf_files = [ Pdf.new() for i in file2pages ]
        # the current pdf file index
        new_pdf_index = 0

        print("Splitting PDF")

        # iterate over all PDF pages
        for n, page in enumerate(pdf.pages):  
            if n in list(range(*file2pages[new_pdf_index])):
                # add the `n` page to the `new_pdf_index` file
                new_pdf_files[new_pdf_index].pages.append(page)
                # print(f"[*] Assigning Page {n} to the file {unit[new_pdf_index]}")
            else:
                # make a unique filename based on original file name plus the index
                name, ext = os.path.splitext(read_file)
                names = name.split("pdf")
                names.pop()
                names.append(path+"/"+folder_name+"pdf splitted")
                name = "/".join(names)
                output_filename = f"{name}/{unit[new_pdf_index]}.pdf"

                # save the PDF file
                new_pdf_files[new_pdf_index].save(output_filename)
                # print(f"[+] File: {output_filename} saved.")
                
                for n_sub, page_sub in enumerate(new_pdf_files[new_pdf_index].pages):
                    dst = Pdf.new()
                    dst.pages.append(page_sub)
                    if n_sub == 0:
                        if type[new_pdf_index] != "SC-SF UT":
                            dst.save(f'{name}/{unit[new_pdf_index]} {type[new_pdf_index]}.pdf')
                            # print(f"[+] File: {name}/{unit[new_pdf_index]} {type[new_pdf_index]}.pdf saved.")
                        else:
                            dst.save(f'{name}/{unit[new_pdf_index]} UT.pdf')
                            # print(f"[+] File: {name}/{unit[new_pdf_index]} UT.pdf saved.")
                    elif n_sub == 1:
                        dst.save(f'{name}/{unit[new_pdf_index]} SC-SF.pdf')
                        # print(f"[+] File: {name}/{unit[new_pdf_index]} SC-SF.pdf saved.")

                # go to the next file
                new_pdf_index += 1
                # add the `n` page to the `new_pdf_index` file
                new_pdf_files[new_pdf_index].pages.append(page)
                # print(f"[*] Assigning Page {n} to the file {unit[new_pdf_index]}")
            
            printProgressBar(n, len(pdf.pages), prefix = 'Progress:', suffix = 'Complete', length = 50)
        # save the last PDF file
        name, ext = os.path.splitext(read_file)
        names = name.split("pdf")
        names.pop()
        names.append(path+"/"+folder_name+"pdf splitted")
        name = "/".join(names)
        output_filename = f"{name}/{unit[new_pdf_index]}.pdf"
        new_pdf_files[new_pdf_index].save(output_filename)
        # print(f"[+] File: {output_filename} saved.")

        for n_sub, page_sub in enumerate(new_pdf_files[new_pdf_index].pages):
            dst = Pdf.new()
            dst.pages.append(page_sub)
            if n_sub == 0:
                if type[new_pdf_index] != "SC-SF UT":
                    dst.save(f'{name}/{unit[new_pdf_index]} {type[new_pdf_index]}.pdf')
                    # print(f"[+] File: {name}/{unit[new_pdf_index]} {type[new_pdf_index]}.pdf saved.")
                else:
                    dst.save(f'{name}/{unit[new_pdf_index]} UT.pdf')
                    # print(f"[+] File: {name}/{unit[new_pdf_index]} UT.pdf saved.")
            elif n_sub == 1:
                dst.save(f'{name}/{unit[new_pdf_index]} SC-SF.pdf')
                # print(f"[+] File: {name}/{unit[new_pdf_index]} SC-SF.pdf saved.")
            
        printProgressBar(n+1, len(pdf.pages), prefix = 'Progress:', suffix = 'Complete', length = 50)
    
    print("DONE!")
    tkinter.messagebox.showinfo("Done", "PDF Splitted")

if __name__ == '__main__':
    main()