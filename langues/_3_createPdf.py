
import win32com.client
import os
from openpyxl import *
from openpyxl.utils import *

cwdRoot=r'C:\Users\Nicolas\Desktop\interets\0) Langues\multilingue'
cwdCode=cwdRoot+r'\0.1) Python'
os.chdir(cwdCode)

import _2_bilingue_mots
_2_bilingue_mots.fileNames
import _2_bilingue_phrases

cwdMots=cwdRoot+r'\1.0.1) Mots, excels'
cwdPhrases=cwdRoot+r'\2.0.1) Phrases, excels'
cwdPdfMotsFromEn=cwdRoot+r'\1.2.2) mots, new format, from en'
cwdPdfPhrasesFromEn=cwdRoot+r'\2.2.2) phrases, new format, from en'


def createPdf(nameXlsx, listeNameTabs, namePdf):
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False #True
    wb = o.Workbooks.Open(nameXlsx)
    
    l=[]
    count = wb.Sheets.Count
    for i in range(1,count+1):
        ws = wb.Sheets(i)
        if (ws.Name in listeNameTabs):
            l.append(i)

    wb.WorkSheets(l).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, namePdf)
    
    o.Quit()
    del o

def remetVisible():
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = True
    o.Quit()
    del o

def createPdf_concret_mots(langueFrom, langueTo):
    nameXlsx_1=_2_bilingue_mots.fileNames[langueTo]
    nameXlsx_1=os.path.splitext(nameXlsx_1)[0]
    #print( filename.rsplit( ".", 1 )[ 0 ] )
    nameXlsx=cwdMots+'\\'+nameXlsx_1+'.xlsx'
    namePdf=cwdPdfMotsFromEn+'\\'+nameXlsx_1+'_from_'+langueFrom+'.pdf'

    listeNameTabs=['from_'+langueFrom+'_1','from_'+langueFrom+'_2']
    
    os.chdir(cwdMots)
    _2_bilingue_mots.createDocBilingue(langueFrom, langueTo)
    createPdf(nameXlsx, listeNameTabs, namePdf)
    
    
def createPdf_concret_phrases(langueFrom, langueTo):
    nameXlsx_1_ph=_2_bilingue_phrases.fileNames[langueTo]
    nameXlsx_1_ph=os.path.splitext(nameXlsx_1_ph)[0]
    #print( filename.rsplit( ".", 1 )[ 0 ] )
    nameXlsx_ph=cwdPhrases+'\\'+nameXlsx_1_ph+'.xlsx'
    namePdf_ph=cwdPdfPhrasesFromEn+'\\'+nameXlsx_1_ph+'_from_'+langueFrom+'.pdf'

    listeNameTabs=['from_'+langueFrom+'_1','from_'+langueFrom+'_2']

    os.chdir(cwdPhrases)
    _2_bilingue_phrases.createDocBilingue(langueFrom, langueTo)
    createPdf(nameXlsx_ph, listeNameTabs, namePdf_ph)

def createPdf_concret(langueFrom, langueTo):
    createPdf_concret_mots(langueFrom, langueTo)
    createPdf_concret_phrases(langueFrom, langueTo)


###############################"
#'ru','sw','ja','zh','de',
for s in ['es','it','pt',\
'tr','el','vi','th','ko','ku','de','yi','ro','fi',\
'cs','tl','nl','mg']:
    createPdf_concret('en',s)

createPdf_concret_mots('en','pal')

remetVisible()



###############
# tests

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False #True
wb_path =cwd+r'\0) en.xlsx' # r'c:\user\desktop\sample.xls'
wb = o.Workbooks.Open(wb_path)

ws_index_list = [1,2] #Sheets to print, starting index is 1
path_to_pdf =cwd+r'\test.pdf'

wb.WorkSheets(ws_index_list).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)

o.Visible = True
o.Quit()
del o

################



##########################"


import win32com.client
o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
wb_path = r'c:\user\desktop\sample.xls'
wb = o.Workbooks.Open(wb_path)

ws_index_list = [1,4,5] #say you want to print these sheets
path_to_pdf = r'C:\user\desktop\sample.pdf'
print_area = 'A1:G50'



for index in ws_index_list:
    #off-by-one so the user can start numbering the worksheets at 1
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area

wb.WorkSheets(ws_index_list).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)



#################################

import glob, os
os.chdir("/mydir")
for file in glob.glob("*.txt"):
    print(file)
    
#or simply os.listdir:
import os
for file in os.listdir("/mydir"):
    if file.endswith(".txt"):
        print(os.path.join("/mydir", file))

#or if you want to traverse directory, use os.walk:
import os
for root, dirs, files in os.walk("/mydir"):
    for file in files:
        if file.endswith(".txt"):
             print(os.path.join(root, file))




