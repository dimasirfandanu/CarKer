#!/usr/bin/python

import tempfile
from zipfile import ZipFile
import shutil
import os
import convertapi
# import platform
from halo import Halo
# import docx2pdf
import envs

# Defining root project directory
rootDIR = os.path.dirname(os.path.abspath(__file__))

# Make working directory
workDIR = tempfile.mkdtemp()

# Asking variables
source = input("Sumber Lowongan?: ") 
company = input("Nama Perusahaan?: ")
position = input("Posisi?: ")

# Preparing PDF
with ZipFile("{}/template.docx".format(rootDIR), "r") as workFILES:
    workFILES.extractall(workDIR)

docXML = "{}/word/document.xml".format(workDIR)
with open(docXML, "r") as docXMLread:
    docXMLedit = docXMLread.read()
docXMLedit = docXMLedit.replace("Source", source)
docXMLedit = docXMLedit.replace("Company", company)
docXMLedit = docXMLedit.replace("Position", position)
with open(docXML, "w") as docXMLwrite:
    docXMLwrite.write(docXMLedit)

docxOUT = "{}/cv.docx".format(workDIR)
shutil.make_archive(docxOUT, "zip", workDIR)
os.rename("{}.zip".format(docxOUT), "{}/cv.docx".format(workDIR)) 

workPDF = "{}/cv.pdf".format(workDIR)
# if platform.system() == "Linux":
   #  Halo(text="Using LibreOffice to create PDF...", spinner="dots").start()
   #  os.system("soffice --convert-to pdf {} --outdir {} &> /dev/null".format(docxOUT, workDIR))
# if platform.system() == "Windows":
   #  Halo(text="Using Microsoft Office to create PDF...", spinner="dots").start()
   #  docx2pdf.convert(docxOUT, workPDF)
Halo(text="Using convertapi to convert PDF...", spinner="dots").start()
convertapi.api_secret = envs.convertapisecret
convertapi.convert('pdf', {'File': docxOUT}, from_format = 'docx').save_files(workDIR)
shutil.copy2(workPDF, "{}/berkas/CV-Oddy-{}-{}.pdf".format(rootDIR, company, position))

# TODO: Sending email
