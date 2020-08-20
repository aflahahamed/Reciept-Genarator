from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.units import inch
import pandas as pd
#loading data
a=pd.read_excel('invoice.xlsx')
#loading image
im=Image('windlogo.jpg',width=1.5*inch,height=inch,hAlign='LEFT')
#The invoice to be printed coressponding to the invoice number
b=input('Invoice Number :');

#Data processing
A=a[a["Invoice Number"] == b]
d=A.customer[A.customer.index[0]];
e=str(A.Date[A.Date.index[0]])
A=A.drop(['customer', 'Invoice Number','Date',], axis=1)
A=A.append({'Amount':'INR '+str(sum(A.Amount)),"Tax":'INR '+str(sum( A.Tax[0:-1])),"Quantity":"NOS "+str(sum(A.Quantity[:-1]))} , ignore_index=True)
A=A.round(2)
A=A.fillna(' ')
s=Spacer(0, 30)
elements = []
styles = getSampleStyleSheet()

#Creating PDF
doc = SimpleDocTemplate((str(b)) + '_' + str(d) +'.pdf')

#Printing data in the PDF
elements.append(im)
styles['Title'].spaceAfter=1
elements.append((Paragraph('<para align=right> WALL MART INC.<para/>', styles['Title'])))
elements.append((Paragraph('<para align=right>MANGALORE<para/>',styles['Normal'])))
elements.append((Paragraph('<para align=right>INDIA<para/>',styles['Normal'])))
elements.append(s)
elements.append(Paragraph("<u>INVOICE</u>", styles['Title']))
elements.append(s)
elements.append((Paragraph('Invoice Number :'+ b,styles['Normal'])))
elements.append((Paragraph('Customer Name :'+ str(d),styles['Normal'])))
elements.append((Paragraph('Date :'+ str(e),styles['Normal'])))
elements.append(s)
lista = [A.columns[:,].values.astype(str).tolist()] + A.values.tolist()
ts= [('LINEABOVE', (0,0), (-1,0), 2, colors.darkgray),
     ('LINEBELOW', (0,0), (-1,0), 2, colors.darkgray),
     ('LINEABOVE', (0,2), (-1,-1), 0.25, colors.black),
     ('LINEBELOW', (0,-1), (-1,-1), 2, colors.darkgray),
     ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
     ('ALIGN', (0,0), (-1,0), 'CENTER'),
     ('BACKGROUND', (-1,-1 ), (-1, -1), colors.green),
     ('BACKGROUND', (-2,-1 ), (-2, -1), colors.blueviolet),
     ('BACKGROUND', (-4,-1 ), (-4, -1), colors.darkkhaki),
     ('LINEABOVE', (0,-1), (-1,-1), 2, colors.darkgray)]
table = Table(lista, style=ts)
elements.append(table)
elements.append(s)

elements.append((Paragraph("For queries contact +91 1234567892", styles["Normal"])))
elements.append(s)
elements.append((Paragraph("<u>THANK YOU</u>", styles["Title"])))

doc.build(elements)

