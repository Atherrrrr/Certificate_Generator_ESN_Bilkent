from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import os
import time

pdfmetrics.registerFont(TTFont('Edwardian Script ITC', 'ITCEDSCR.TTF'))

def create_file(studentName):

  packet = io.BytesIO()
  can = canvas.Canvas(packet)
  can.setPageSize((840, 590))
  can.setFont("Edwardian Script ITC", 52)
  can.drawCentredString(410, 235, studentName)
  can.save()

  #move to the beginning of the StringIO buffer
  packet.seek(0)

  # create a new PDF with Reportlab
  new_pdf = PdfReader(packet)
  # read your existing PDF
  existing_pdf = PdfReader(open("certificate template.pdf", "rb"))
  output = PdfWriter()
  # add the "watermark" (which is the new pdf) on the existing page
  page = existing_pdf.pages[0]
  page.merge_page(new_pdf.pages[0])
  output.add_page(page)
  # finally, write "output" to a real file
  outputStream = open(f'Certificates/{studentName}.pdf', "wb")
  output.write(outputStream)
  outputStream.close()

#############
print('---The Program has started generating pdfs---\n\n\n')
df = pd.read_excel("names.xlsx")

path = "Certificates"
isExist = os.path.exists(path)

if (isExist == False):
  os.mkdir(path)

namesList = []
for name in df.values:
  namesList.append(name[0])

for name in namesList:
  create_file(name)

print('---ALL PDFs Generated--')
time.sleep(600)