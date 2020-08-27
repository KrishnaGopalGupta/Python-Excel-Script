import xlrd
import xlwt
from xlutils.copy import copy
import time
from sys import argv
import datetime as dtime


#Opoen the Excel
print('Script started at : ',dtime.datetime.now())

book = xlrd.open_workbook(argv[1]) # St By EMail StBYEMAIL.xlsx


# Sheet name
first_sheet = book.sheet_by_index(0)  # St by email Sheet


# Create the workbook and sheet for updating the file
wb = copy(book);
w_sheet = wb.get_sheet(0)


# index of St BY email
print('St By email index')

for i in range(0,first_sheet.ncols):

    if first_sheet.cell(0,i).value=="Portal Primary Contact: Email Address":
        pindex1=i
    elif first_sheet.cell(0,i).value=="Portal Primary Franchise Name":
        pindex2=i
    elif first_sheet.cell(0,i).value=="Portal Store Name":
        pindex3=i      
    elif first_sheet.cell(0,i).value=="Portal Primary Signing Authority: First Name":
        pindex4=i
    elif first_sheet.cell(0,i).value=="Portal Primary Signing Authority: Last Name":
        pindex5=i
    elif first_sheet.cell(0,i).value=="Data Franchise Name":
        pindex6=i
    elif first_sheet.cell(0,i).value=="Data Primary Contact: Email Address":
        pindex7=i
    elif first_sheet.cell(0,i).value=="Data Primary Signing Authority: First Name":
        pindex8=i
    elif first_sheet.cell(0,i).value=="Data Primary Signing Authority: Last Name":
        pindex9=i      
    elif first_sheet.cell(0,i).value== "St Due in 45 Days":
        pindex10=i
    elif first_sheet.cell(0,i).value== "St Over Due":
        pindex11=i
    elif first_sheet.cell(0,i).value== "St Activation Link  Expiring in 14 Days":
        pindex12=i
    elif first_sheet.cell(0,i).value== "St Expired Activation Link":
        pindex13=i
    elif first_sheet.cell(0,i).value== "St Expired in Next 45 days":
        pindex14=i
    elif first_sheet.cell(0,i).value== "St Submitted":
        pindex15=i
    elif first_sheet.cell(0,i).value== "St Assigned":
        pindex16=i
    elif first_sheet.cell(0,i).value=="Exception table":
        pindex17=i
    elif first_sheet.cell(0,i).value=="Comments":
        pindex18=i
    elif first_sheet.cell(0,i).value=="Remarks":
        pindex19=i


# Data Data Update


# Data Indexing


def Data_Update():
    
    print('Data Idexing and Update')

    mbook = xlrd.open_workbook(argv[2]) # Data
    m_sheet = mbook.sheet_by_index(0) # Data Sheet
    
    for i in range(0,m_sheet.ncols):

        if m_sheet.cell(0,i).value=="LLC+Account number (pdt only)":
            mindex1=i
        elif m_sheet.cell(0,i).value=="Primary Signing Authority: Email":
            mindex2=i
        elif m_sheet.cell(0,i).value=="PC Number (Store Number)":
            mindex3=i      #Store
        elif m_sheet.cell(0,i).value=="Primary Signing Authority: First Name":
            mindex4=i
        elif m_sheet.cell(0,i).value=="Primary Signing Authority: Last Name":
            mindex5=i

    for i in range(1,first_sheet.nrows):
        for j in range(1,m_sheet.nrows):
            if int(first_sheet.cell(i,pindex3).value) == int(m_sheet.cell(j,mindex3).value):  # Stores equal
                w_sheet.write(i,pindex6,m_sheet.cell(j,mindex1).value)
                w_sheet.write(i,pindex7,m_sheet.cell(j,mindex2).value)
                w_sheet.write(i,pindex8,m_sheet.cell(j,mindex4).value)
                w_sheet.write(i,pindex9,m_sheet.cell(j,mindex5).value)
                w_sheet.write(i,pindex17,"N/A")
                break;
      

def Excepion_Update():
    print('Data Excepion Data Idexing and Update')
    ebook = xlrd.open_workbook(argv[3]) # Data Exception
    e_sheet = ebook.sheet_by_index(0) # Data Sheet
    
    for i in range(0,e_sheet.ncols):

        if e_sheet.cell(0,i).value=="LLC+Account number (pdt only)":
            eindex1=i
        elif e_sheet.cell(0,i).value=="Primary Signing Authority: Email":
            eindex2=i
        elif e_sheet.cell(0,i).value=="PC Number (Store Number)":
            eindex3=i      #Store
        elif e_sheet.cell(0,i).value=="Primary Signing Authority: First Name":
            eindex4=i
        elif e_sheet.cell(0,i).value=="Primary Signing Authority: Last Name":
            eindex5=i

    for i in range(1,first_sheet.nrows):
        for j in range(1,e_sheet.nrows):
            if int(first_sheet.cell(i,pindex3).value) == int(e_sheet.cell(j,eindex3).value):  # Stores equal
                w_sheet.write(i,pindex6, e_sheet.cell(j,eindex1).value)
                w_sheet.write(i,pindex7, e_sheet.cell(j,eindex2).value)
                w_sheet.write(i,pindex8, e_sheet.cell(j,eindex4).value)
                w_sheet.write(i,pindex9, e_sheet.cell(j,eindex5).value)
                w_sheet.write(i,pindex17, e_sheet.cell(j,eindex3).value)
                break;


def Aging_data_update():
    print('St Aging Data Idexing and Update')
    abook = xlrd.open_workbook(argv[4]) # St Aging Report
    a_sheet = abook.sheet_by_index(0)   # St Aging Report Sheet

    for i in range(0,a_sheet.ncols):
        if a_sheet.cell(0,i).value=="Primary Signing Authority: Email":
            aindex1=i
        elif a_sheet.cell(0,i).value=="Due (in Days)":
            aindex2=i
        elif a_sheet.cell(0,i).value=="Overdue(Days)":
            aindex3=i
        elif a_sheet.cell(0,i).value=="St Expiry Date":
            aindex4=i
        elif a_sheet.cell(0,i).value=="Activation Due in (Days)":
            aindex5=i
        elif a_sheet.cell(0,i).value=="Activation OverDue(Days)":
            aindex6=i

    for i in range(1,first_sheet.nrows):
        
        for j in range(1,a_sheet.nrows):
            if first_sheet.cell(i,pindex1).value.strip() == a_sheet.cell(j,aindex1).value.strip():
                w_sheet.write(i,pindex10, a_sheet.cell(j,aindex2).value)
                w_sheet.write(i,pindex11, a_sheet.cell(j,aindex3).value)
               # w_sheet.write(i,pindex12, a_sheet.cell(j,aindex5).value)
                #w_sheet.write(i,pindex13, a_sheet.cell(j,aindex6).value)
                w_sheet.write(i,pindex14, a_sheet.cell(j,aindex4).value)
                break;
     



Data_Update()
print('Data Update Done')

Excepion_Update()
print('Exception Data Update Done')
            
Aging_data_update()
print('Aging Data Update Done')



wb.save("St Data Fetched Done.xls");

print('Script ended at : ',dtime.datetime.now())

print("Fetched => Data Update, Exceptional Data Update, St Aging Data !!");





