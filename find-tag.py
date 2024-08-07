import os
from numpy import size
from odf import text, teletype
from odf.opendocument import load

resultTable = []

pathToProcess = input("Indiquer le répertoire des documents = ")
if pathToProcess == "":
    pathToProcess = "."

#define header names
col_names = ["Document", "Champs de fusion"]

# search and replace a text in a odt document
files_src = [name for name in os.listdir(pathToProcess) if name.endswith(".odt")]

print("DOCUMENT\tCHAMPS DE FUSION")

for file_src in files_src:
    
    # load input document
    document = load(pathToProcess+'/'+file_src)

    # search in paragraphs
    res = ""
    paragraphs = document.getElementsByType(text.P)
    for paragraph in paragraphs:        
        txt = teletype.extractText(paragraph)
        a = txt.split("{")
        if len(a) > 1:
            for x in a:
                b = x.split("}")
                if len(b) > 1:
                    if res == "":
                        res = b[0]
                    else:
                        res = res + ", " + b[0]
                    
    print(file_src+"\t"+res)

input("---PAUSE---")