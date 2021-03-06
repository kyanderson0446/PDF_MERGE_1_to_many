from glob import glob
import os
import PyPDF2
from PyPDF2 import PdfFileMerger
""""Input to guide you the pre-generated folder thanks to the VBA script"""
folder = input(str("Please Enter \"YYYY MM\": "))

""""Updated the new attachments in the Share Drive. May need to be renamed"""
path = fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\OriginalFinancePackagePDF\*"        # Save pdfs to this folder. Compelted via VBA
wc_file = fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\ReportsToAdd\WC\*.pdf"           # Save WC to this path as pdf
ar_file = fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\ReportsToAdd\AR\*.pdf"           # Save AR to this path as pdf


"""Begin the process of reports to send by appending the attachments"""
try:
    os.mkdir(fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\ReportsToSend")               # Creates ReportsToSend folder in input(folder)
except:
    pass
counter = 0
#Iterate through files and assign the original name with the files appended
for filename in glob(path):
    if filename.endswith(".pdf"):
        pdftopFile = open(filename, 'rb')                                                   # Pdf hierarchy
        pdfmidFile = open(ar_file, 'rb')
        pdfbottomFile = open(wc_file, 'rb')
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
        save_location = fr"P:\PACS\Finance\FP&A\Finance Package\{folder}\ReportsToSend\{os.path.basename(filename)}"

        pdfOutputFile = open(save_location, 'wb')                                           # Save under Updated w WC with same filename
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        pdftopFile.close()
        pdfmidFile.close()
        pdfbottomFile.close()
        counter = counter + 1

        print(f"{os.path.basename(filename)} APPEND {os.path.basename(ar_file)} APPEND {os.path.basename(wc_file)}")
print("Total Files Merged -> ", counter)
