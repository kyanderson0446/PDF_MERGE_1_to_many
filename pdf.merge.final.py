from glob import glob
import os
import PyPDF2
from PyPDF2 import PdfFileMerger

folder = input(str("Please Enter \"YYYY MM\": "))                                            #{Folder} should already be created. It just needs to be directed to that folder


path = fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\*"
wc_file = fr"P:\PACS\Finance\FP&A\Finance Package\ReportsToAdd\01-22 Workers Comp Scorecard Consolidated.xlsx v2.pdf"
ar_file = fr"P:\PACS\Finance\FP&A\Finance Package\ReportsToAdd\220216 - AR Dashboard_Jan22.pdf"


try:
    os.mkdir(path[:-1] + "\\Updated w WC AR")
except:
    pass

#Iterate through files and assign the original name with the files appended
for filename in glob(path):
    if filename.endswith(".pdf"):
        pdftopFile = open(filename, 'rb')                                                   #pdf hierarchy
        pdfbottomFile = open(wc_file, 'rb')
        pdfmidFile = open(ar_file, 'rb')
        pdftopReader = PyPDF2.PdfFileReader(pdftopFile)
        pdfmidReader = PyPDF2.PdfFileReader(pdfmidFile)
        pdfbottomReader = PyPDF2.PdfFileReader(pdfbottomFile)
        pdfWriter = PyPDF2.PdfFileWriter()

        for pageNum in range(pdftopReader.numPages):
            pageObj = pdftopReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)
        for pageNum in range(pdfmidReader.numPages):
            pageObj = pdfmidReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)
        for pageNum in range(pdfbottomReader.numPages):
            pageObj = pdfbottomReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        # assign to updated folder
        save_location = fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\Updated w WC AR{os.path.basename(filename)}.pdf"

        pdfOutputFile = open(save_location, 'wb')  # save under Updated w WC with same filename
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        pdftopFile.close()
        pdfmidFile.close()
        pdfbottomFile.close()

        print(f"{os.path.basename(filename)} APPEND {os.path.basename(wc_file)} APPEND {os.path.basename(ar_file)}")