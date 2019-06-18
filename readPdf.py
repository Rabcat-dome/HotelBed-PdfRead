import os
from datetime import datetime
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import BytesIO as StringIO
from cStringIO import StringIO as StringIO2
import xlwt 
from xlwt import Workbook 
import PyPDF2 

path = os.path.join(os.path.dirname(os.path.realpath(__file__)),'pdfFile')

folders = []

# r=root, d=directories, f = files
for r, d, f in os.walk(path):
    for folder in f:
        if folder.endswith('.pdf') or folder.endswith('.PDF'):
            folders.append(os.path.join(r, folder))

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    try:
        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
            interpreter.process_page(page)
        text = retstr.getvalue()
        
    except:
        print("Mal PDF FORM") 
    
    fp.close()
    device.close()
    retstr.close()

    startT = 0
    endT = 0
    while(text.find("HOTELBEDS (THAILAND) LIMITED", endT)!=-1):
        startT = text.find("HOTELBEDS (THAILAND) LIMITED", startT)
        endT = text.find("Please indicate our reference number on each of your invoices.\n--------------------------------------------------------------------------------", startT)
        if endT == -1:
            endT = text.find("Please provide the following services:\n \n--------------------------------------------------------------------------------", startT)
        text = text[0:startT].strip()+"\n"+text[endT+len("Please provide the following services:\n \n--------------------------------------------------------------------------------"):len(text)].strip()
    text = text.strip()

    while True:
        if text.find("Page:",0)!=-1:
            text = text[0:text.find("Page:",0)-1].strip() +"\n"+ text[text.find("Page:",0)+11:].strip()
        else:
            break

    #print(text)
    return text
        
# Workbook is created 
wb = Workbook() 
row = 1
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet1') 
sheet1.col(0).width = 15*256
sheet1.col(1).width = 30*256
sheet1.col(2).width = 20*256
sheet1.col(3).width = 10*256
sheet1.col(4).width = 20*256
sheet1.col(5).width = 15*256
sheet1.col(6).width = 15*256
sheet1.col(7).width = 15*256
sheet1.col(8).width = 15*256
sheet1.col(9).width = 15*256
sheet1.col(10).width = 27*256
sheet1.col(11).width = 20*256
sheet1.col(12).width = 15*256
sheet1.col(13).width = 40*256
sheet1.col(14).width = 38*256
sheet1.col(15).width = 70*256
sheet1.col(16).width = 25*256
sheet1.col(17).width = 20*256
sheet1.col(18).width = 20*256
sheet1.col(19).width = 25*256
sheet1.col(21).width = 30*256
sheet1.col(22).width = 100*256
sheet1.col(24).width = 25*256
sheet1.col(25).width = 40*256
sheet1.col(26).width = 25*256
sheet1.col(28).width = 50*256
sheet1.col(29).width = 40*256
sheet1.col(30).width = 60*256
sheet1.write(0, 0, "refNo")
sheet1.write(0, 1, "CustomerName")
sheet1.write(0, 2, "Agent refNo")
sheet1.write(0,3,"Type")
sheet1.write(0,4,"Booking Date")
sheet1.write(0,5,"Service Id")
sheet1.write(0,6,"Contract")
sheet1.write(0,7,"Service Date")
sheet1.write(0,8,"Service")
sheet1.write(0,9,"Modality")
sheet1.write(0,10,"Service Description")
sheet1.write(0,11,"Modality Description")
sheet1.write(0,12,"Rate")
sheet1.write(0,13,"PAX")
sheet1.write(0,14,"Customer Detail")
sheet1.write(0,15,"REMARK")
sheet1.write(0,16,"Hotel")
sheet1.write(0,17,"Cancel BookingDate")
sheet1.write(0,18,"Modification BookingDate")
sheet1.write(0,19,"Client Mobile")
sheet1.write(0,20,"Arrival")
sheet1.write(0,21,"Departure")
sheet1.write(0,22,"Commercial description")
sheet1.write(0,23,"Serv.Type")
sheet1.write(0,24,"Vehicle")
sheet1.write(0,25,"From")
sheet1.write(0,26,"PickUp Time")
sheet1.write(0,27,"PickUp point")
sheet1.write(0,28,"Transport")
sheet1.write(0,29,"To")
sheet1.write(0,30,"OldBooking (Modification)")

for page in folders:
    text2 = convert_pdf_to_txt(page)
    allText = len(text2)
    startText = 0
    indexStart = 0
    indexEndLine = 0
    reference = []
    referenceNum = ""
    referenceName = ""
    typeDetail=""
    mode=""

    while startText <= allText:
        bookDate=""
        serviceId=""
        contract=""
        serviceDate=""
        service=""
        modality=""
        serviceDes=""
        modalDes=""
        rate=""
        pax=""
        cust=""
        remark=""
        hotel=""
        ModiDate=""
        CancelDate=""
        arrival=""
        depeart=""
        ClientMobile=""
        comDes=""
        serviceType=""
        venhicle=""
        from1=""
        pickuptime=""
        pickuppoint=""
        transport=""
        to1=""
        oldBook = ""
        indexReference = text2.find("REFERENCE", startText,startText+300)

        if indexReference!=-1: #Page2
            indexReferenceNum = text2.find("\n", indexReference+9)
            indexEndLine = text2.find("\n", indexReferenceNum+1)
            reference = text2[indexReferenceNum+1:indexEndLine].strip().split()
            referenceNum = reference[0]

        if(reference[1]!= "Name:"):
            mode = "noname"
            referenceName = reference[1]
            for x in range(len(reference)-2):
                if reference[2+x] in ("TRAVEL","MONTROSE","FIVEFLY","CARASOULS","WEBBEDS","SIMPLE","FLIGHT","NOE","MTRAVEL-Main","HOORAY","WAU!"):
                    break
                referenceName += " "+reference[2+x]

            if indexReference!=-1: #Page2
                indexAgencyRef = text2.find("Agency Reference:", indexEndLine)
                indexEndLine = text2.find("\n", indexAgencyRef+1)
                AgencyReference = text2[indexAgencyRef+17:indexEndLine].strip()

            indexType = text2.find("---", indexEndLine+1,indexEndLine+100)

            if indexType != -1:
                indexEndLine = text2.find("\n", indexType+1)
            typeDetail=""

            if(indexType == -1 or "NEW" in text2[indexType+1:indexType+100]):
                typeDetail="NEW"

                indexBook = text2.find("Creation booking date:", indexEndLine+1)
                if indexBook ==-1:
                    break;
                indexEndLine = text2.find("\n", indexBook+1)
                bookDate = text2[indexBook+22:indexEndLine].strip()

                indexService = text2.find("SERVICE ID", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                serviceId = text2[indexService+len("SERVICE ID"):indexEndLine].strip()

                indexStart = text2.find("Contract:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                contract = text2[indexStart+len("Contract:"):indexEndLine].strip()

                indexStart = text2.find("Service Date", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                serviceDate = text2[indexStart+len("Service Date"):indexEndLine].strip()

                indexStart = text2.find("Service:", indexEndLine+1)
                indexEndLine = text2.find("Modality:", indexStart+1)
                service = text2[indexStart+len("Service:"):indexEndLine-1].strip()

                indexStart = text2.find("Modality:", indexEndLine-1)
                indexEndLine = text2.find("\n", indexStart+1)
                modality = text2[indexStart+len("Modality:"):indexEndLine].strip()

                indexStart = text2.find("Service Description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                serviceDes = text2[indexStart+len("Service Description:"):indexEndLine].strip()

                indexStart = text2.find("Modality Description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                modalDes = text2[indexStart+len("Modality Description:"):indexEndLine].strip()

                indexStart = text2.find("Rate:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                rate = text2[indexStart+len("Rate:"):indexEndLine].strip()

                indexStart = text2.find("PASSENGERS:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                pax = text2[indexStart+len("PASSENGERS:"):indexEndLine].strip()

                indexStart = indexEndLine
                indexEndLine = text2.find("REMARKS:", indexStart-1)
                cust = text2[indexStart:indexEndLine].strip()
                cust1 = cust.splitlines()
                cust=""
                for custs in cust1:
                    cust+=custs.strip() + "\n"
                cust = cust.strip()

                indexStart = text2.find("REMARKS:", indexEndLine)
                indexEndLine = text2.find("Confirmation Number", indexStart-1)
                remark = text2[indexStart+len("REMARKS:"):indexEndLine].strip()

                indexStart = text2.find("of your hotel -", indexStart+1)
                indexEndLine2 = text2.find("\n", indexStart+1)
                hotel = text2[indexStart+len("of your hotel -"):indexEndLine2].strip()

            elif ("CANCELLATION" in text2[indexType+1:indexType+100]):
                typeDetail="CANCEL"

                indexBook = text2.find("Creation booking date:", indexEndLine+1)
                if indexBook == -1:
                    break;
                indexEndLine = text2.find("\n", indexBook+1)
                bookDate = text2[indexBook+len("Creation booking date:"):indexEndLine].strip()

                indexCancel = text2.find("Cancellation booking date:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexCancel+1)
                CancelDate = text2[indexCancel+len("Cancellation booking date:"):indexEndLine].strip()

                indexService = text2.find("SERVICE ID", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                serviceId = text2[indexService+len("SERVICE ID"):indexEndLine].strip()

                indexStart = text2.find("Contract:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                contract = text2[indexStart+len("Contract:"):indexEndLine].strip()

                indexStart = text2.find("Service Date", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                serviceDate = text2[indexStart+len("Service Date"):indexEndLine].strip()

                indexStart = text2.find("Service:", indexEndLine+1)
                indexEndLine = text2.find("Modality:", indexStart+1)
                service = text2[indexStart+len("Service:"):indexEndLine-1].strip()

                indexStart = text2.find("Modality:", indexEndLine-1)
                indexEndLine = text2.find("\n", indexStart+1)
                modality = text2[indexStart+len("Modality:"):indexEndLine].strip()

                indexStart = text2.find("Service Description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                serviceDes = text2[indexStart+len("Service Description:"):indexEndLine].strip()

                indexStart = text2.find("Modality Description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                modalDes = text2[indexStart+len("Modality Description:"):indexEndLine].strip()

                indexStart = text2.find("Rate:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                rate = text2[indexStart+len("Rate:"):indexEndLine].strip()

                indexStart = text2.find("PASSENGERS:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                pax = text2[indexStart+len("PASSENGERS:"):indexEndLine].strip()

                indexStart = indexEndLine
                indexEndLine = text2.find("REMARKS:", indexStart-1)
                cust = text2[indexStart:indexEndLine].strip()
                cust1 = cust.splitlines()
                cust=""
                for custs in cust1:
                    cust+=custs.strip() + "\n"
                cust = cust.strip()

                indexStart = text2.find("REMARKS:", indexEndLine)
                indexEndLine = text2.find("Confirmation Number", indexStart-1)
                remark = text2[indexStart+len("REMARKS:"):indexEndLine].strip()

                indexStart = text2.find("of your hotel -", indexStart+1)
                indexEndLine2 = text2.find("\n", indexStart+1)
                hotel = text2[indexStart+len("of your hotel -"):indexEndLine2].strip()

            elif ("MODIFICATION" in text2[indexType+1:indexType+100]):
                typeDetail="MODIFY"

                indexBook = text2.find("Creation booking date:", indexEndLine+1)
                if indexBook == -1:
                    break;
                indexEndLine = text2.find("\n", indexBook+1)
                bookDate = text2[indexBook+len("Creation booking date:"):indexEndLine].strip()

                indexCancel = text2.find("Modification booking date:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexCancel+1)
                ModiDate = text2[indexCancel+len("Modification booking date:"):indexEndLine].strip()

                indexService = text2.find("SERVICE ID", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                serviceId = text2[indexService+len("SERVICE ID"):indexEndLine].strip()

                indexStart = text2.find("Contract:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                contract = text2[indexStart+len("Contract:"):indexEndLine].strip()

                indexStart = text2.find("Service Date", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                serviceDate = text2[indexStart+len("Service Date"):indexEndLine].strip()

                indexStart = text2.find("Service:", indexEndLine+1)
                indexEndLine = text2.find("Modality:", indexStart+1)
                service = text2[indexStart+len("Service:"):indexEndLine-1].strip()

                indexStart = text2.find("Modality:", indexEndLine-1)
                indexEndLine = text2.find("\n", indexStart+1)
                modality = text2[indexStart+len("Modality:"):indexEndLine].strip()

                indexStart = text2.find("Service Description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                serviceDes = text2[indexStart+len("Service Description:"):indexEndLine].strip()

                indexStart = text2.find("Modality Description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                modalDes = text2[indexStart+len("Modality Description:"):indexEndLine].strip()

                indexStart = text2.find("Rate:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                rate = text2[indexStart+len("Rate:"):indexEndLine].strip()

                indexStart = text2.find("PASSENGERS:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                pax = text2[indexStart+len("PASSENGERS:"):indexEndLine].strip()

                indexStart = indexEndLine
                indexEndLine = text2.find("REMARKS:", indexStart-1)
                cust = text2[indexStart:indexEndLine].strip()
                cust1 = cust.splitlines()
                cust=""
                for custs in cust1:
                    cust+=custs.strip() + "\n"
                cust = cust.strip()

                indexStart = text2.find("REMARKS:", indexEndLine)
                indexEndLine = text2.find("Confirmation Number", indexStart-1)
                remark = text2[indexStart+len("REMARKS:"):indexEndLine].strip()

                indexStart = text2.find("of your hotel -", indexStart+1)
                indexEndLine2 = text2.find("\n", indexStart+1)
                hotel = text2[indexStart+len("of your hotel -"):indexEndLine2].strip()
            else:
                break;
        else:
            mode = "withname"
            referenceName = reference[2]
            for x in range(len(reference)-3):
                referenceName += " "+reference[3+x]

            indexAgencyRef = text2.find("TO.Ref.", indexEndLine,indexEndLine+100)
            if indexAgencyRef != -1:
                indexEndLine = text2.find("\n", indexAgencyRef+1)
                AgencyReference = text2[indexAgencyRef+7:indexEndLine].strip()

            indexMobile = text2.find("mobile contact:", indexEndLine ,indexEndLine+100 )
            if indexMobile != -1:
                indexEndLine = text2.find("\n", indexMobile+1)
                ClientMobile = text2[indexMobile+len("mobile contact:"):indexEndLine].strip()
            else:
                ClientMobile=""

            indexType = text2.find("---", indexEndLine+1,indexEndLine+100)
            if indexType != -1:
                indexEndLine = text2.find("\n", indexType+1)
                typeDetail=""

            if(( indexType == -1 and typeDetail=="NEW") or "NEW" in text2[indexType+1:indexType+100]):
                typeDetail="NEW"

                indexArr= text2.find("Arrival:", indexEndLine+1,indexEndLine+200)
                indexEndLineArr = text2.find("\n", indexArr+1)
                if indexArr != -1:
                    arrival = text2[indexArr+len("Arrival:"):indexEndLineArr].strip()
                else:
                    arrival = ""

                indexDe= text2.find("Departure:", indexEndLine+1,indexEndLine+200)
                indexEndLineDe = text2.find("\n", indexDe+1)
                if indexDe != -1:
                    depeart = text2[indexDe+len("Departure:"):indexEndLineDe].strip()
                else:
                    depeart = ""

                indexBook = text2.find("Creation booking date:", indexEndLine+1)
                if indexBook ==-1:
                    break;
                indexEndLine = text2.find("\n", indexBook+1)
                bookDate = text2[indexBook+len("Creation booking date:"):indexEndLine].strip()

                indexService = text2.find("SERVICE ID:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                serviceId = text2[indexService+len("SERVICE ID:"):indexEndLine].strip()

                indexStart = text2.find("Commercial description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                comDes = text2[indexStart+len("Commercial description:"):indexEndLine].strip()
                
                indexStart = text2.find("Serv.Type:", indexEndLine+1)
                indexEndLine = text2.find("Vehicle:", indexStart+1)
                serviceType = text2[indexStart+len("Serv.Type:"):indexEndLine-1].strip()

                indexStart = text2.find("Vehicle:", indexEndLine-1)
                indexEndLine = text2.find("Paxes:", indexStart+1)
                venhicle = text2[indexStart+len("Vehicle:"):indexEndLine-1].strip()

                indexStart = text2.find("Paxes:", indexEndLine-1)
                indexEndLine = text2.find("\n", indexStart+1)
                pax = text2[indexStart+len("Paxes:"):indexEndLine].strip()

                indexService = text2.find("From:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                if indexDe != -1:
                    indexEndLineTest = text2.find("\n", indexEndLine+1)
                    indexEndLine = text2.find("\n", indexEndLineTest+1)
                    if text2[indexEndLineTest:indexEndLine].isspace():
                        indexEndLine = indexEndLineTest
                from1 = text2[indexService+len("From:"):indexEndLine].strip()

                indexService = text2.find("Pickup Time:", indexEndLine+1,indexEndLine+100)
                if indexService != -1:
                    indexEndLine = text2.find("hrs", indexService+1)
                    pickuptime = text2[indexService+len("Pickup Time:"):indexEndLine-1].strip()
                else:
                    pickuptime = ""

                indexService = text2.find("Pickup Point:", indexEndLine+1,indexEndLine+100)
                if indexService != -1:
                    indexEndLineP = text2.find("\n", indexService+1)
                    pickuppoint = text2[indexService+len("Pickup Point:"):indexEndLineP].strip()
                else:
                    pickuppoint = ""

                indexService = text2.find("Transport:", indexEndLineP+1)
                indexEndLine1 = text2.find("\n", indexService+1)
                transport = text2[indexService+len("Transport:"):indexEndLine1].strip()

                indexService = text2.find("To: ", indexEndLineP+1)
                indexEndLine = text2.find("\n", indexService+1)
                if indexArr != -1:
                    indexEndLineTest = text2.find("\n", indexEndLine+1)
                    indexEndLine = text2.find("\n", indexEndLineTest+1)
                    if text2[indexEndLineTest:indexEndLine].isspace():
                        indexEndLine = indexEndLineTest
                to1 = text2[indexService+len("To: "):indexEndLine].strip()

                if indexEndLine1 > indexEndLine:
                    indexEndLine = indexEndLine1

            elif (( indexType == -1 and typeDetail=="CANCEL") or"CANCELLATION" in text2[indexType+1:indexType+100]):
                typeDetail="CANCEL"

                indexArr= text2.find("Arrival:", indexEndLine+1,indexEndLine+200)
                indexEndLineArr = text2.find("\n", indexArr+1)
                if indexArr != -1:
                    arrival = text2[indexArr+len("Arrival:"):indexEndLineArr].strip()
                else:
                    arrival=""

                indexDe= text2.find("Departure:", indexEndLine+1,indexEndLine+200)
                indexEndLineDe = text2.find("\n", indexDe+1)
                if indexDe != -1:
                    depeart = text2[indexDe+len("Departure:"):indexEndLineDe].strip()
                else:
                    depeart=""

                indexBook = text2.find("Creation booking date:", indexEndLine+1)
                if indexBook ==-1:
                    break;
                indexEndLine = text2.find("\n", indexBook+1)
                bookDate = text2[indexBook+len("Creation booking date:"):indexEndLine].strip()

                indexCancel = text2.find("Cancellation booking date:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexCancel+1)
                CancelDate = text2[indexCancel+len("Cancellation booking date:"):indexEndLine].strip()

                indexService = text2.find("SERVICE ID:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                serviceId = text2[indexService+len("SERVICE ID:"):indexEndLine].strip()

                indexStart = text2.find("Commercial description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                comDes = text2[indexStart+len("Commercial description:"):indexEndLine].strip()
                
                indexStart = text2.find("Serv.Type:", indexEndLine+1)
                indexEndLine = text2.find("Vehicle:", indexStart+1)
                serviceType = text2[indexStart+len("Serv.Type:"):indexEndLine-1].strip()

                indexStart = text2.find("Vehicle:", indexEndLine-1)
                indexEndLine = text2.find("Paxes:", indexStart+1)
                venhicle = text2[indexStart+len("Vehicle:"):indexEndLine-1].strip()

                indexStart = text2.find("Paxes:", indexEndLine-1)
                indexEndLine = text2.find("\n", indexStart+1)
                pax = text2[indexStart+len("Paxes:"):indexEndLine].strip()

                indexService = text2.find("From:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                if indexDe != -1:
                    indexEndLineTest = text2.find("\n", indexEndLine+1)
                    indexEndLine = text2.find("\n", indexEndLineTest+1)
                    if text2[indexEndLineTest:indexEndLine].isspace():
                        indexEndLine = indexEndLineTest
                from1 = text2[indexService+len("From:"):indexEndLine].strip()

                indexService = text2.find("Pickup Time:", indexEndLine+1,indexEndLine+100)
                if indexService != -1:
                    indexEndLine = text2.find("hrs", indexService+1)
                    pickuptime = text2[indexService+len("Pickup Time:"):indexEndLine-1].strip()
                else:
                    pickuptime=""

                indexService = text2.find("Pickup Point:", indexEndLine+1,indexEndLine+100)
                if indexService != -1:
                    indexEndLineP = text2.find("\n", indexService+1)
                    pickuppoint = text2[indexService+len("Pickup Point:"):indexEndLineP].strip()
                else:
                    pickuppoint =""

                indexService = text2.find("To: ", indexEndLineP+1)
                indexEndLine = text2.find("\n", indexService+1)
                if indexArr != -1:
                    indexEndLineTest = text2.find("\n", indexEndLine+1)
                    indexEndLine = text2.find("\n", indexEndLineTest+1)
                    if text2[indexEndLineTest:indexEndLine].isspace():
                        indexEndLine = indexEndLineTest
                to1 = text2[indexService+len("To: "):indexEndLine].strip()

                indexService = text2.find("Transport:", indexEndLineP+1)
                indexEndLine1 = text2.find("\n", indexService+1)
                transport = text2[indexService+len("Transport:"):indexEndLine1].strip()

                if indexEndLine1 > indexEndLine:
                    indexEndLine = indexEndLine1

            elif (( indexType == -1 and typeDetail=="MODIFY") or "MODIFICATION" in text2[indexType+1:indexType+100]):
                typeDetail="MODIFY"

                indexArr= text2.find("Arrival:", indexEndLine+1,indexEndLine+200)
                indexEndLineArr = text2.find("\n", indexArr+1)
                if indexArr != -1:
                    arrival = text2[indexArr+len("Arrival:"):indexEndLineArr].strip()
                else:
                    arrival=""

                indexDe= text2.find("Departure:", indexEndLine+1,indexEndLine+200)
                indexEndLineDe = text2.find("\n", indexDe+1)
                if indexDe != -1:
                    depeart = text2[indexDe+len("Departure:"):indexEndLineDe].strip()
                else:
                    depeart=""

                indexBook = text2.find("Creation booking date:", indexEndLine+1)
                if indexBook ==-1:
                    break;
                indexEndLine = text2.find("\n", indexBook+1)
                bookDate = text2[indexBook+len("Creation booking date:"):indexEndLine].strip()

                indexCancel = text2.find("Modification booking date:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexCancel+1)
                ModiDate = text2[indexCancel+len("Modification booking date:"):indexEndLine].strip()

                indexService = text2.find("SERVICE ID:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                serviceId = text2[indexService+len("SERVICE ID:"):indexEndLine].strip()

                indexStart = text2.find("Commercial description:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexStart+1)
                comDes = text2[indexStart+len("Commercial description:"):indexEndLine].strip()
                
                indexStart = text2.find("Serv.Type:", indexEndLine+1)
                indexEndLine = text2.find("Vehicle:", indexStart+1)
                serviceType = text2[indexStart+len("Serv.Type:"):indexEndLine-1].strip()

                indexStart = text2.find("Vehicle:", indexEndLine-1)
                indexEndLine = text2.find("Paxes:", indexStart+1)
                venhicle = text2[indexStart+len("Vehicle:"):indexEndLine-1].strip()

                indexStart = text2.find("Paxes:", indexEndLine-1)
                indexEndLine = text2.find("\n", indexStart+1)
                pax = text2[indexStart+len("Paxes:"):indexEndLine].strip()

                indexService = text2.find("From:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexService+1)
                if indexDe != -1:
                    indexEndLineTest = text2.find("\n", indexEndLine+1)
                    indexEndLine = text2.find("\n", indexEndLineTest+1)
                    if text2[indexEndLineTest:indexEndLine].isspace():
                        indexEndLine = indexEndLineTest
                from1 = text2[indexService+len("From:"):indexEndLine].strip()

                indexService = text2.find("Pickup Time:", indexEndLine+1,indexEndLine+100)
                if indexService != -1:
                    indexEndLine = text2.find("hrs", indexService+1)
                    pickuptime = text2[indexService+len("Pickup Time:"):indexEndLine-1].strip()
                else:
                    pickuptime=""

                indexService = text2.find("Pickup Point:", indexEndLine+1,indexEndLine+100)
                if indexService != -1:
                    indexEndLineP = text2.find("\n", indexService+1)
                    pickuppoint = text2[indexService+len("Pickup Point:"):indexEndLineP].strip()
                else:
                    pickuppoint = ""

                indexService = text2.find("To: ", indexEndLineP+1)
                indexEndLine = text2.find("\n", indexService+1)
                if indexArr != -1:
                    indexEndLineTest = text2.find("\n", indexEndLine+1)
                    indexEndLine = text2.find("\n", indexEndLineTest+1)
                    if text2[indexEndLineTest:indexEndLine].isspace():
                        indexEndLine = indexEndLineTest
                to1 = text2[indexService+len("To: "):indexEndLine].strip()

                indexService = text2.find("Transport:", indexEndLineP+1)
                indexEndLine1 = text2.find("\n", indexService+1)
                transport = text2[indexService+len("Transport:"):indexEndLine1].strip()

                if indexEndLine1 > indexEndLine:
                    indexEndLine = indexEndLine1

                indexService = text2.find("Old booking:", indexEndLine+1)
                indexEndLine = text2.find("Transport:", indexService+1)
                indexEndLine = text2.find("To:", indexEndLine+1)
                indexEndLine = text2.find("\n", indexEndLine+1)
                oldBook = text2[indexService+len("Old booking:"):indexEndLine].strip()

            else:
                break

        sheet1.write(row, 0, referenceNum)
        sheet1.write(row, 1, referenceName) 
        sheet1.write(row,2,AgencyReference)
        sheet1.write(row,3,typeDetail)
        sheet1.write(row,4,bookDate)
        sheet1.write(row,5,serviceId)
        sheet1.write(row,6,contract)
        sheet1.write(row,7,serviceDate)
        sheet1.write(row,8,service)
        sheet1.write(row,9,modality)
        sheet1.write(row,10,serviceDes)
        sheet1.write(row,11,modalDes)
        sheet1.write(row,12,rate)
        sheet1.write(row,13,pax)
        sheet1.write(row,14,cust)
        sheet1.write(row,15,remark)
        sheet1.write(row,16,hotel)
        sheet1.write(row,17,CancelDate)
        sheet1.write(row,18,ModiDate)
        sheet1.write(row,19,ClientMobile)
        sheet1.write(row,20,arrival)
        sheet1.write(row,21,depeart)
        sheet1.write(row,22,comDes)
        sheet1.write(row,23,serviceType)
        sheet1.write(row,24,venhicle)
        sheet1.write(row,25,from1)
        sheet1.write(row,26,pickuptime)
        sheet1.write(row,27,pickuppoint)
        sheet1.write(row,28,transport)
        sheet1.write(row,29,to1)
        sheet1.write(row,30,oldBook)

        startText = indexEndLine+1
        
        #print("********************")
        #print(text2[startText:startText+400])
        #print(row)
        row +=1

wb.save(datetime.now().strftime('%Y%m%d_%H%M%S')+'_HT.xls') 