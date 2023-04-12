import os
from pikepdf import Pdf
from pprint import pprint
import pickle as pkl
import glob
import os

def set_file2pages():
    # arr = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 1, 2, 2, 2, 1, 2, 1, 1]
    arr = {}
    with open('sequence.pkl', 'rb') as f:
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

    return file2pages, arr["unit"]

# a dictionary mapping PDF file to original PDF's page range
file2pages, unit = set_file2pages()

# the target PDF document to split
# filename = "result/pdf/Final Invoice April 23.pdf"

# get all csv files
path =  os.path.dirname(os.path.realpath(__file__))
folder_name = "result/pdf/"
read_files = glob.glob(path+"/" + folder_name + "/*.pdf")

#save the file
# checking if the directory exist or not.
if not os.path.exists(path +"/" + folder_name + "/split/"):
    # then create it.
    os.makedirs(path + "/" + folder_name + "/split/")

# loop all pdf files
for read_file in read_files:
    # load the PDF file
    pdf = Pdf.open(read_file)

    # make the new splitted PDF files
    new_pdf_files = [ Pdf.new() for i in file2pages ]
    # the current pdf file index
    new_pdf_index = 0

    # iterate over all PDF pages
    for n, page in enumerate(pdf.pages):
        if n in list(range(*file2pages[new_pdf_index])):
            # add the `n` page to the `new_pdf_index` file
            new_pdf_files[new_pdf_index].pages.append(page)
            print(f"[*] Assigning Page {n} to the file {unit[new_pdf_index]}")
        else:
            # make a unique filename based on original file name plus the index
            name, ext = os.path.splitext(read_file)
            names = name.split("pdf")
            names.pop()
            names.append("pdf/split")
            name = "/".join(names)
            output_filename = f"{name}/{unit[new_pdf_index]}.pdf"

            # save the PDF file
            new_pdf_files[new_pdf_index].save(output_filename)
            print(f"[+] File: {output_filename} saved.")
            # go to the next file
            new_pdf_index += 1
            # add the `n` page to the `new_pdf_index` file
            new_pdf_files[new_pdf_index].pages.append(page)
            print(f"[*] Assigning Page {n} to the file {unit[new_pdf_index]}")

    # save the last PDF file
    name, ext = os.path.splitext(read_file)
    names = name.split("pdf")
    names.pop()
    names.append("pdf/split")
    name = "/".join(names)
    output_filename = f"{name}/{unit[new_pdf_index]}.pdf"
    new_pdf_files[new_pdf_index].save(output_filename)
    print(f"[+] File: {output_filename} saved.")