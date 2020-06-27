#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from datetime import timedelta
from datetime import datetime
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code128


# In[ ]:


pdfmetrics.registerFont(TTFont('Avenir', 'AvenirLTStd-Book.ttf'))
df=pd.read_excel (r'C:\Users\pc\Desktop\programmes pyth\Dataset.xlsx', sheet_name='Sheet1')#importé
code=df['code']#sauvegarde collonne code dans code
first_name=df['first_name']
last_name=df['last_name']
function=df['function']
date_of_birth=df['date_of_birth']
submission_date=df['submission_date']
picture=df['picture']


# In[124]:


for i in range (200):
    X=str(date_of_birth[i])
    Y=datetime.strptime(X,'%d/%m/%Y')
    a=str(Y.day)
    b=str(Y.month)
    c=str(Y.year)
    if len(a)<2:
        a="0"+a
    if len(b)<2:
        b="0"+b
    birth_day=a+'-'+b+'-'+c
    
    x=str(submission_date[i])
    y=datetime.strptime(x,'%d/%m/%Y')
    z=timedelta(days=15)
    T=y+z
    A=str(T.day)
    B=str(T.month)
    C=str(T.year)
    if len(A)<2:
        A="0"+A
    if len(B)<2:
        B="0"+B
    exp_day=A+'-'+B+'-'+C
    
    
    L=[code[i],first_name[i],last_name[i],function[i],birth_day,exp_day,picture[i]]
    text=str(L[0])#code
    text2=L[1]+' '+L[2]#nom prenom
    text3=L[3]#fonction
    text4=L[4]#birth
    text5=L[5]#expire
    image=L[6]#link to image
    packet = io.BytesIO()# create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Avenir',8)
    can.setFillColorRGB(0,0,0)
    can.drawString(154, 75, text)
    can.drawString(157,55,text4)
    can.drawString(157,35,text5)
    can.setFillColorRGB(255,255,255)
    can.drawString(16,50,text2)
    can.setFont('Avenir',6)
    can.drawString(16,40,text3)
    can.drawImage(image,16,61,width=0.87*inch,height=0.87*inch)
    barcode = code128.Code128(text)
    barcode.drawOn(can, 0, 10)
    can.save()#move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)# read your existing PDF
    existing_pdf = PdfFileReader(open(r"C:\Users\pc\Desktop\programmes pyth\A8.pdf", "rb"))
    output = PdfFileWriter()# add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)# finally, write "output" to a real file
    outputStream = open(r"C:\Users\pc\Desktop\programmes pyth\Output_files\card{0}.pdf".format(i+1), "wb")
    output.write(outputStream)
    outputStream.close()
    