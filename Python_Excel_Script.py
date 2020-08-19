import xlrd
import xlwt
from xlutils.copy import copy
import time

#Opoen the Excel
print('Open the Excel File')


book = xlrd.open_workbook("SAQBYEMAIL.xlsx") # SAQ By EMail
abook = xlrd.open_workbook("Aging.xlsx") # SAQ Aging Report
mbook = xlrd.open_workbook("MDF.xlsx") # MDF
ebook = xlrd.open_workbook("Exception.xlsx") # Exception Data

print('Read the file throgh sheet -- MDF , Exception, SAQ Aging, SAQ By Email')

# Sheet name
first_sheet = book.sheet_by_index(0)  # SAQ by email Sheet
a_sheet = abook.sheet_by_index(0)   # SAQ Aging Report Sheet
m_sheet = mbook.sheet_by_index(0) # MDF Sheet
e_sheet = ebook.sheet_by_index(0) # Exception Sheet

# Create the workbook and sheet for updating the file
wb = copy(book);
w_sheet = wb.get_sheet(0)


# index of SAQ BY email
print('SAQ By email index')

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
    elif first_sheet.cell(0,i).value=="MDF Franchise Name":
        pindex6=i
    elif first_sheet.cell(0,i).value=="MDF Primary Contact: Email Address":
        pindex7=i
    elif first_sheet.cell(0,i).value=="MDF Primary Signing Authority: First Name":
        pindex8=i
    elif first_sheet.cell(0,i).value=="MDF Primary Signing Authority: Last Name":
        pindex9=i      
    elif first_sheet.cell(0,i).value== "SAQ Due in 45 Days":
        pindex10=i
    elif first_sheet.cell(0,i).value== "SAQ Over Due":
        pindex11=i
    elif first_sheet.cell(0,i).value== "SAQ Activation Link  Expiring in 14 Days":
        pindex12=i
    elif first_sheet.cell(0,i).value== "SAQ Expired Activation Link":
        pindex13=i
    elif first_sheet.cell(0,i).value== "SAQ Expired in Next 30 days":
        pindex14=i
    elif first_sheet.cell(0,i).value== "SAQ Submitted":
        pindex15=i
    elif first_sheet.cell(0,i).value== "SAQ Assigned":
        pindex16=i
    elif first_sheet.cell(0,i).value=="Exception table":
        pindex17=i
    elif first_sheet.cell(0,i).value=="Comments":
        pindex18=i
    elif first_sheet.cell(0,i).value=="Remarks":
        pindex19=i


# SAQ Aging report index
print('--SAQ Aging Index--')

for i in range(0,a_sheet.ncols):

    if a_sheet.cell(0,i).value=="SAQ Primary Signing Authority: Email":
        aindex1=i
    elif a_sheet.cell(0,i).value=="Due (in Days)":
        aindex2=i
    elif a_sheet.cell(0,i).value=="Overdue(Days)":
        aindex3=i
    elif a_sheet.cell(0,i).value=="SAQ Expiry Date":
        aindex4=i
    elif a_sheet.cell(0,i).value=="Activation Due in (Days)":
        aindex5=i
    elif a_sheet.cell(0,i).value=="Activation OverDue(Days)":
        aindex6=i


# Variable

s1 = "SAQ Assigned - Yes"
s2 = "SAQ Assigned - No"
sc1 = "Yes"
sc2 = "No"



# #SAQ By email VS SAQ aging

print('------SAQ By email and SAQ aging report---------')

r11 = int(first_sheet.nrows)
r12 = int(a_sheet.nrows)


def SAQaging_update():

    for i in range(1,r11):
        for j in range(1,r12):
            if first_sheet.cell(i,pindex1).value.strip() == a_sheet.cell(j,aindex1).value.strip():
                w_sheet.write(i,pindex10,a_sheet.cell(j,aindex2).value)
                w_sheet.write(i,pindex11,a_sheet.cell(j,aindex3).value)
                w_sheet.write(i,pindex12,a_sheet.cell(j,aindex5).value)
                w_sheet.write(i,pindex13,a_sheet.cell(j,aindex6).value)
                w_sheet.write(i,pindex14,a_sheet.cell(j,aindex4).value)
                break;


# MDF Data Update
print("-----------MDF Data Update---------------")

# MDF Indexing
print('MDF Idexing')

def MDF_Update():
    
    for i in range(0,m_sheet.ncols):

        if m_sheet.cell(0,i).value=="MDF Franchise Name":
            mindex1=i
        elif m_sheet.cell(0,i).value=="MDF Primary Contact: Email Address":
            mindex2=i
        elif m_sheet.cell(0,i).value=="MDF Store Name":
            mindex3=i      #Store
        elif m_sheet.cell(0,i).value=="MDF Primary Signing Authority: First Name":
            mindex4=i
        elif m_sheet.cell(0,i).value=="MDF Primary Signing Authority: Last Name":
            mindex5=i

    for i in range(1,first_sheet.nrows):
        for j in range(1,m_sheet.nrows):
            if int(first_sheet.cell(i,pindex3).value) == int(m_sheet.cell(j,mindex3).value):  # Stores equal
                w_sheet.write(i,pindex6,m_sheet.cell(j,mindex1).value)
                w_sheet.write(i,pindex7,m_sheet.cell(j,mindex2).value)
                w_sheet.write(i,pindex8,m_sheet.cell(j,mindex4).value)
                w_sheet.write(i,pindex9,m_sheet.cell(j,mindex5).value)
                
# Exceptional Data Update

print('---------Exceptional data----------')

print('Exception data indexing')

for i in range(0,e_sheet.ncols):
    if e_sheet.cell(0,i).value=="Exception Franchise Name":
        eindex1=i
    elif e_sheet.cell(0,i).value=="Exception Primary Contact: Email Address":
        eindex2=i
    elif e_sheet.cell(0,i).value=="Exception Store Name":
        eindex3=i      #Store
    elif e_sheet.cell(0,i).value=="Exception Primary Signing Authority: First Name":
        eindex4=i
    elif e_sheet.cell(0,i).value=="Exception Primary Signing Authority: Last Name":
        eindex5=i

def Exceptional_data():
    for i in range(1,first_sheet.nrows):
        for j in range(1,e_sheet.nrows):
            if int(first_sheet.cell(i,pindex3).value) == int(e_sheet.cell(j,eindex3).value):  # Stores equal
                w_sheet.write(i,pindex17,e_sheet.cell(j,eindex3).value)
                w_sheet.write(i,pindex6,e_sheet.cell(j,eindex1).value)
                w_sheet.write(i,pindex7,e_sheet.cell(j,eindex2).value)
                w_sheet.write(i,pindex8,e_sheet.cell(j,eindex4).value)
                w_sheet.write(i,pindex9,e_sheet.cell(j,eindex5).value)
                
                break;

            else:
                w_sheet.write(i,pindex17,"N/A")

# Comments And Remarks Update
print('----Comments and Remarks-----')

remark1="Action Required from Comcast / DBI"
remark2="Action Required from MDR Ops (POD)."
remark3="Confirmation required for launching the SAQ."
remark4="No Action Required"

se1="N/A"

# Use cases

a1  = "User has completed the SAQ but user legal entity name has changed(Use Case 9)"
a2  = "User has completed the SAQ but user data (i.e., Franchise and Email address) has changed(Use Case 10)"
a3  = "SAQ is assigned but user has not completed.(Use Case 17)"
a4  = "User has completed the SAQ.(Use Case 7)"
a5  = "User has not completed the SAQ but user franchise name changed(Use Case 16)"
a6  = "SAQ completed, Exceptional Tab User (Franchise and Email id) is  changed.(Use Case 27)"
a7  = "User has not completed the SAQ but user data (i.e., Franchise and Email address) has changed(Use Case 25)"
a8  = "SAQ completed, Exceptional Tab User email id is  changed.(Use Case 21)"
a10 = "SAQ is assigned but user has completed and email id changed(Use Case 3)."
a11 = "SAQ is assigned but user has not completed and email id changed(Use Case 13)."
a12 = "SAQ not completed, Exceptional Tab User email id is changed.(Use Case 22)."
a13 = "Ready To Launch(Use Case 20)"
a15 = "User has completed the SAQ but email id is changed(Use Case 3)."
a16 = "SAQ completed, Exceptional Tab user legal entity name is changed."
a17 = "SAQ is assigned but user has not completed, Exceptional Tab user legal entity name is changed."
a18  = "User has not completed the SAQ but Exceptional Tab user data (i.e., Franchise and Email address) has changed(Use Case 25)"


#r1 = first_sheet.nrows; # no. of row
#c1 = first_sheet.ncols;  #  no. of column

def Comment_update():
    for i in range(1, r1):
        if first_sheet.cell(i,pindex16).value==s1: # Assigned
            if first_sheet.cell(i,pindex15).value==sc1:   # Completed
                if first_sheet.cell(i,pindex17).value==se1:  # no Exceptional
                    if first_sheet.cell(i,pindex1).value.lower() == first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower() == first_sheet.cell(i,pindex6).value.lower():  # primary email and france
                        w_sheet.write(i,pindex18,a4)
                        w_sheet.write(i,pindex19,remark4)
                    elif first_sheet.cell(i,pindex1).value.lower()!=first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a15)
                        w_sheet.write(i,pindex19,remark2)
                    elif first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()!=first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a1)
                        w_sheet.write(i,pindex19,remark4)
                    else:
                        w_sheet.write(i,pindex18,a2)
                        w_sheet.write(i,pindex19,remark4)
                        
                else:   # Exceptional
                    
                   if first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a4)
                        w_sheet.write(i,pindex19,remark4)
                   elif first_sheet.cell(i,pindex1).value.lower()!=first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a8)
                        w_sheet.write(i,pindex19,remark1) 
                   elif first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()!=first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a16)
                        w_sheet.write(i,pindex19,remark1) 
                   else:
                        w_sheet.write(i,pindex18,a6)
                        w_sheet.write(i,pindex19,remark1) 
                    
                
            else:  # SAQ assigned but not completed
                if first_sheet.cell(i,pindex17).value==se1: # no Exceptional
                    if first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a3)
                        w_sheet.write(i,pindex19,remark4)
                    elif first_sheet.cell(i,pindex1).value.lower()!=first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a11)
                        w_sheet.write(i,pindex19,remark2)
                    elif first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()!=first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a5)
                        w_sheet.write(i,pindex19,remark2)
                    else:
                        w_sheet.write(i,pindex18,a7)
                        w_sheet.write(i,pindex19,remark2)
                else:  # Exceptional
                   if first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a3)
                        w_sheet.write(i,pindex19,remark4)  
                   elif first_sheet.cell(i,pindex1).value.lower()!=first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()==first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a12)
                        w_sheet.write(i,pindex19,remark1)  
                   elif first_sheet.cell(i,pindex1).value.lower()==first_sheet.cell(i,pindex7).value.lower() and first_sheet.cell(i,pindex2).value.lower()!=first_sheet.cell(i,pindex6).value.lower():
                        w_sheet.write(i,pindex18,a17)
                        w_sheet.write(i,pindex19,remark1) 
                   else:
                        w_sheet.write(i,pindex18,a18)
                        w_sheet.write(i,pindex19,remark1) 
                    
      # SAQ Not Assigned
        else:
          
            w_sheet.write(i,pindex18,a13);
            w_sheet.write(i,pindex19,remark3)
           #  if first_sheet.cell(i,19).value.lower()=="N/A":




    '''
    for i in range(1, first_sheet.nrows):
        if (first_sheet.cell(i,pindex16).value.strip()==s1) and (first_sheet.cell(i,pindex15).value.strip() == sc1) and (first_sheet.cell(i,pindex17).value.upper().strip()=="N/A") :
            if (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a4)
                w_sheet.write(i,pindex19,remark4)

            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a15)
                w_sheet.write(i,pindex19,remark2)

            elif (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a1)
                w_sheet.write(i,pindex19,remark4)

            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a2)
                w_sheet.write(i,pindex19,remark4)


        elif (first_sheet.cell(i,pindex16).value.strip() == s1) and (first_sheet.cell(i,pindex15).value.strip() == sc1) and (first_sheet.cell(i,pindex17).value.upper().strip() != "N/A") :
            if (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):   
                w_sheet.write(i,pindex18,a4)
                w_sheet.write(i,pindex19,remark4)  
            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a8)
                w_sheet.write(i,pindex19,remark1)  
            elif (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a16)
                w_sheet.write(i,pindex19,remark1)  
            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                print(first_sheet.cell(i,pindex17).value)
                w_sheet.write(i,pindex18,a6)
                w_sheet.write(i,pindex19,remark1)

        elif (first_sheet.cell(i,pindex16).value.strip() == s1) and (first_sheet.cell(i,pindex15).value.strip() == sc2) and (first_sheet.cell(i,pindex17).value.upper().strip() == "N/A"):
            if (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a3)
                w_sheet.write(i,pindex19,remark4)         
            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a11)
                w_sheet.write(i,pindex19,remark2)           
            elif (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower() != first_sheet.cell(i,pindex6).value.lower()):
                w_sheet.write(i,pindex18,a5)
                w_sheet.write(i,pindex19,remark2)  
            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a7)
                w_sheet.write(i,pindex19,remark2)
            

        elif (first_sheet.cell(i,pindex16).value.strip() == s1) and (first_sheet.cell(i,pindex15).value.strip() == sc2) and (first_sheet.cell(i,pindex17).value.upper().strip() != "N/A"):
            if (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower()):
                w_sheet.write(i,pindex18,a3)
                w_sheet.write(i,pindex19,remark4)        
            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() == first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a12)
                w_sheet.write(i,pindex19,remark1)        
            elif (first_sheet.cell(i,pindex1).value.lower().strip() == first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a17)
                w_sheet.write(i,pindex19,remark1)
            elif (first_sheet.cell(i,pindex1).value.lower().strip() != first_sheet.cell(i,pindex7).value.lower().strip()) and (first_sheet.cell(i,pindex2).value.lower().strip() != first_sheet.cell(i,pindex6).value.lower().strip()):
                w_sheet.write(i,pindex18,a18)
                w_sheet.write(i,pindex19,remark1)

                
      # SAQ Not Assigned
        else:  
            w_sheet.write(i,pindex18,a13);
            w_sheet.write(i,pindex19,remark3)
    '''

print('----------------------------------------')

MDF_Update()
print('MDF Update Done')

time.sleep(10)

Exceptional_data()
print('Exceptional Data Update Done')

SAQaging_update()
print('SAQ Aging Data Update Done')

Comment_update()
print('Comments and Remarks Update Done')

wb.save("SAQ_REC_Updated_1-11.xls");
print("SAQ Recon data is prepared and file is created!!");





