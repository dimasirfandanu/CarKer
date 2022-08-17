#!/usr/bin/python

import tempfile
from zipfile import ZipFile
import shutil
import os
import convertapi
import envs
from yaspin import yaspin
import platform
import docx2pdf
from time import sleep

# Init
with yaspin(text="Loading... "):
    sleep(3)

# Asking variables
source = input("Sumber Lowongan?: ") 

company = input("Nama Perusahaan?: ")

position = input("Posisi?: ")

name = input("Nama Pelamar?: ")

# Defining consts
rootDIR = os.path.dirname(os.path.abspath(__file__))

berkasUSER = "{}/{}".format(rootDIR, name)
if os.path.exists(berkasUSER) == False:
    os.mkdir(berkasUSER)

def taskCOMPLETE():
    print("✅ Task completed")

# Make working directory
workDIR = tempfile.mkdtemp()

# Preparing PDF
with ZipFile("{}/template.docx".format(rootDIR), "r") as workFILES:
    workFILES.extractall(workDIR)

docxXML = "{}/word/document.xml".format(workDIR)
with open(docxXML, "r") as docxXMLread:
    docxXMLedit = docxXMLread.read()
docxXMLedit = docxXMLedit.replace("Source", source)
docxXMLedit = docxXMLedit.replace("Company", company)
docxXMLedit = docxXMLedit.replace("Position", position)
with open(docxXML, "w") as docxXMLwrite:
    docxXMLwrite.write(docxXMLedit)

docxOUT = "{}/cv.docx".format(workDIR)
shutil.make_archive(docxOUT, "zip", workDIR)
os.rename("{}.zip".format(docxOUT), "{}/cv.docx".format(workDIR)) 

pdfOUT = "{}/cv.pdf".format(workDIR)
if platform.system() == "Linux":
    with yaspin(text="Using libreoffice to create PDF..."):
        os.system("soffice --convert-to pdf {} --outdir {} &> /dev/null".format(docxOUT, workDIR))
    taskCOMPLETE()
else:
    with yaspin(text="Using convertapi to create PDF..."):
        convertapi.api_secret = envs.convertapisecret
        convertapi.convert('pdf', {'File': docxOUT}, from_format = 'docx').save_files(workDIR)
    taskCOMPLETE()

shutil.copy2(pdfOUT, "{}/CV-{}-{}-{}.pdf".format(berkasUSER, name, company, position))
# TODO: Better cross-platform function
# if platform.system() == "Windows":
#     Halo(text="Using Microsoft Office to create PDF...", spinner="dots").start() # TODO: Migrate to yaspin
#     docx2pdf.convert(docxOUT, pdfOUT)

# TODO: Sending email
