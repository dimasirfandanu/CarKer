#!/usr/bin/python

import tempfile
from zipfile import ZipFile
import shutil
import os
import re
import sys
import subprocess

# Defining root project directory
rootDIR = os.path.dirname(os.path.abspath(__file__))

# Make working directory
workDIR = tempfile.mkdtemp()

# Asking variables
source = input("Sumber Lowongan?:") 
company = input("Nama Perusahaan?:")
position = input("Posisi?:")

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
os.rename(docxOUT + ".zip", "{}/cv.docx".format(workDIR)) 

# TODO: search for better function cross-platfrom
def convert_to(folder, source, timeout=None):
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', folder, source]
    process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    filename = re.search('-> (.*?) using filter', process.stdout.decode())
    return filename.group(1)
def libreoffice_exec():
    # TODO: Provide support for more platforms
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'
convert_to(workDIR, docxOUT)

workPDF = "{}/cv.pdf".format(workDIR)
shutil.copy2(workPDF, "{}/berkas/CV-Oddy-{}-{}.pdf".format(rootDIR, company, position))

# TODO: send email
