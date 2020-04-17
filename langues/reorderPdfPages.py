# http://www.blog.pythonlibrary.org/2012/07/11/pypdf2-the-new-fork-of-pypdf/

#from openpyxl import *
#from openpyxl.utils import *
#from unidecode import unidecode
#import copy
#import win32com.client

import os
import sys
import os.path
import glob
from PyPDF2 import PdfFileReader, PdfFileWriter

decap = lambda s: s[:1].lower() + s[1:] if s else ''

cwd1=r'/Users/nicolas/Desktop/0) Langues/multilingue/3) all grand'
l1=(glob.glob(cwd1+r'/*.pdf'))
cwd2=r'/Users/nicolas/Desktop/0) Langues/multilingue/3) all grand/sub2'
l2=(glob.glob(cwd2+r'/*.pdf'))
cwd3=r'/Users/nicolas/Desktop/0) Langues/multilingue/3) all grand/sub3'

blankPagePath = cwd3 + r"/blankPage.pdf"
phrasesPath = cwd1 + "/iwAlph_to_te.pdf"
# suffix = "_motsSep"
motsPath = cwd2 + "/en_to_fi_motsSep.pdf"
outputPath = cwd3 + "/iwAlph_to_te.pdf"

pdfBlank = PdfFileReader(open(blankPagePath, 'rb'))#file( blankPagePath, "rb"))
pdfPhrases = PdfFileReader(open( phrasesPath, "rb"))
pdfMots = PdfFileReader(open(motsPath, "rb"))

l=[0,4,1,5,2,6,3,7]

output = PdfFileWriter()
output.addPage(pdfBlank.getPage(0))
for i in range(pdfMots.getNumPages()):
    output.addPage(pdfMots.getPage(i))
for i in range(pdfPhrases.getNumPages()-4):
    output.addPage(pdfPhrases.getPage(l[i]+4))
    
outputStream = open(outputPath , "wb")
output.write(outputStream)
outputStream.close()

def fileName(fileName):
    x=len(fileName)
    while fileName[x-1]!='/':
        x-=1
    return fileName[x:]

def addSuffix(mot, suffix):
    return mot[:(-4)] + suffix + ".pdf"
#addSuffix("coucou.pdf","_2")

ll=(glob.glob(cwd1+r'/*.pdf'))
blankPagePath = cwd3 + r"/blankPage.pdf"
l=[0,4,1,5,2,6,3,7]
for i in range(len(ll)):
    phrasesPath = ll[i] # cwd1 + "/iwAlph_to_te.pdf"
    nomFile = fileName(phrasesPath)
    motsPath = cwd2 + "/" + addSuffix(nomFile,"_motsSep")
    outputPath = cwd3 + "/" + nomFile
    
    pdfBlank = PdfFileReader(open(blankPagePath, 'rb'))#file( blankPagePath, "rb"))
    pdfPhrases = PdfFileReader(open( phrasesPath, "rb"))
    pdfMots = PdfFileReader(open(motsPath, "rb"))

    output = PdfFileWriter()
    output.addPage(pdfBlank.getPage(0))
    for i in range(pdfMots.getNumPages()):
        output.addPage(pdfMots.getPage(i))
    for i in range(pdfPhrases.getNumPages()-4):
        output.addPage(pdfPhrases.getPage(l[i]+4))
        
    outputStream = open(outputPath , "wb")
    output.write(outputStream)
    outputStream.close()

    print(nomFile +'\t: done')
    sys.stdout.flush()



####################

from PyPDF2 import PdfFileReader, PdfFileWriter
 
output = PdfFileWriter()
pdfOne = PdfFileReader(file( "some\path\to\a\PDf", "rb"))
pdfTwo = PdfFileReader(file("some\other\path\to\a\PDf", "rb"))
 
output.addPage(pdfOne.getPage(0))
output.addPage(pdfTwo.getPage(0))
 
outputStream = file(r"output.pdf", "wb")
output.write(outputStream)
outputStream.close()


###########


import PyPDF2
 
path = open('path/to/hello.pdf', 'rb')
path2 = open('path/to/another.pdf', 'rb')
 
merger = PyPDF2.PdfFileMerger()
 
merger.merge(position=0, fileobj=path2)
merger.merge(position=2, fileobj=path)
merger.write(open("test_out.pdf", 'wb'))


#############


import PyPDF2
 
path = open('path/to/hello.pdf', 'rb')
path2 = open('path/to/another.pdf', 'rb')
 
merger = PyPDF2.PdfFileMerger()
 
merger.append(fileobj=path2)
merger.append(fileobj=path)
merger.write(open("test_out2.pdf", 'wb'))

###############


from PyPDF2 import PdfFileReader 
p = r'C:\Users\mdriscoll\Documents\reportlab-userguide.pdf' 
pdf = PdfFileReader(open(p, 'rb')) 
pdf.documentInfo
pdf.getNumPages()
info = pdf.getDocumentInfo()
info.author
info.creator
info.producer
info.title
 


######################
from pyPdf import PdfFileWriter, PdfFileReader
output_pdf = PdfFileWriter()

with open(r'input.pdf', 'rb') as readfile:
    input_pdf = PdfFileReader(readfile)
    total_pages = input_pdf.getNumPages()
    for page in xrange(total_pages - 1, -1, -1):
        output_pdf.addPage(input_pdf.getPage(page))
    with open(r'output.pdf', "wb") as writefile:
        output_pdf.write(writefile)




