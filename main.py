import jinja2
import pyqrcode
# from pyqrcode import QRCode
import png
import os
import pdfkit
from xhtml2pdf import pisa
from PyPDF2 import PdfFileMerger
import pandas as pd
import traceback

options = {'page-size': 'Letter'}
reportTimeCity = {'Bhopal': 'Morning 8:00 AM'}
allowanceCity = {'Bhopal': '150', 'Indore': '100', 'Jaipur': '100', 'Jodhpur': '100'}


def createQRPng(qrName):
    # s = "this is test QR"
    qrLocation = "/Applications/XAMPP/xamppfiles/htdocs/qr/"
    # Generate QR code
    url = pyqrcode.create(qrName)
    # Create and save the png file naming "myqr.png"
    url.png(f"{qrLocation}{qrName}.png", scale=6)
    print(f"QR generated with value {qrName}")
    return str(qrName)


def render_html(groupID, seid, city, studentName, grade, fathersName, examCenter, examDate, shift, reportingTime,
                schoolName, schoolAddress, volunteerName):
    """
    Render html page using jinja
    """
    hope = [5, 6, 7]
    growth = [8, 9]
    if grade in hope:
        template_loader = jinja2.FileSystemLoader(searchpath="./")
        template_env = jinja2.Environment(loader=template_loader)
        template_file = "Templates/Sitare Admit Card-Hindi-Hope-V2.3T/SitareAdmitCardHindiHopeV2.3T.html"
        template = template_env.get_template(template_file)
        output_text = template.render(
            seid=seid,
            allowance=allowanceCity[city],
            studentName=studentName,
            currentSchool=schoolName,
            schoolAddress=schoolAddress,
            fathersName=fathersName,
            examCenter=examCenter,
            examDate=examDate,
            grade=grade,
            shift=shift,
            reportingTime=reportingTime,
            qrCode=createQRPng(seid),  # createQRPng(idNumber)
            volunteerId=volunteerName
        )

        html_path = f'JaipurNewHtml/{volunteerName}-{seid}.html'
        html_file = open(html_path, 'w')
        html_file.write(output_text)
        html_file.close()
        print(f"Hope html with seid {seid} saved successfully")
        pdfkit.from_file(html_path, f'JaipurNew/{volunteerName}-{groupID}-{seid}.pdf', options=options)
        print("PDF Saved Successful")


    elif grade in growth:
        template_loader = jinja2.FileSystemLoader(searchpath="./")
        template_env = jinja2.Environment(loader=template_loader)
        template_file = "Templates/Sitare Admit -GrowthV2.3 T/SitareAdmitGrowthV2.3T.html"
        template = template_env.get_template(template_file)
        output_text = template.render(
            seid=str(seid),
            allowance=allowanceCity[city],
            studentName=str(studentName),
            currentSchool=str(schoolName),
            schoolAddress=str(schoolAddress),
            fathersName=str(fathersName),
            examCenter=str(examCenter),
            examDate=str(examDate),
            grade=str(int(grade)),
            shift=str(shift),
            reportingTime=str(reportingTime),
            qrCode=createQRPng(seid),  # createQRPng(idNumber)
            volunteerId=volunteerName,
        )
        html_path = f'JaipurNewHtml/{volunteerName}-{seid}.html'
        html_file = open(html_path, 'w')
        html_file.write(output_text)
        html_file.close()
        print(f"Growth html with seid {seid} saved successfully")
        pdfkit.from_file(html_path, f'JaipurNew/{volunteerName}-{groupID}-{seid}.pdf', options=options)
        print("PDF Saved Successful")


# render_html("123456881", "JAI-A123456", "Harshit", "Papa", "Mansarovar", "2 Feb 2022", "12:30PM")

def renderBlankAdmitCard(seid):
    """
    Render html page using jinja
    """
    template_loader = jinja2.FileSystemLoader(searchpath="./")
    template_env = jinja2.Environment(loader=template_loader)
    template_file = "Templates/Sitare Blank -AdmitCard.3 T/SitareBlankAdmitCardV2.3T.html"
    template = template_env.get_template(template_file)
    output_text = template.render(
        seid=str(seid),
        qrCode=createQRPng(seid),  # createQRPng(idNumber)
    )
    html_path = f'BlankAdmitCardHtml/{seid}.html'
    html_file = open(html_path, 'w')
    html_file.write(output_text)
    html_file.close()
    print(f"Blank html with seid {seid} saved successfully")
    pdfkit.from_file(html_path, f'JaipurBlank/{seid}.pdf', options=options)
    print("PDF Saved Successful")


def allowance(city):
    if city.casefold() == "Bhopal".casefold():
        return 150
    else:
        return 100


def iterateCsv():
    csvLoc = str(input("Enter CSV Location"))
    df = pd.read_csv(csvLoc)
    # seid, city,studentName, grade, fathersName, examCenter, examDate, shift, reportingTime,schoolName,volunteerName
    for groupID, seid, city, studentName, grade, fathersName, examCenter, examDate, shift, reportingTime, schoolName, schoolAddress, volunteerName in zip(
            df['Group Id'], df['SEID'], df['city'], df['Name'], df['Grade'], df['Father Name'], df['Center'],
            df['Exam Date'], df['Shift'], df['ExamTime'], df['School Name'], df['School Address'],
            df['volunteer Email']):
        try:
            render_html(groupID, seid, city, studentName, grade, fathersName, examCenter,
                        examDate, shift, reportingTime,
                        schoolName, schoolAddress, volunteerName)
        except Exception as e:
            print(f"Something went wrong with {seid} with {traceback.print_exc()}")


def iterateBlankCsv():
    csvLoc = str(input("Enter CSV Location"))
    df = pd.read_csv(csvLoc)
    # seid, city,studentName, grade, fathersName, examCenter, examDate, shift, reportingTime,schoolName,volunteerName
    for seid in df['SEID']:
        try:
            renderBlankAdmitCard(seid)
        except Exception as e:
            print(f"Something went wrong with {seid} with {traceback.print_exc()}")

def iterateBlankPreFilledCsv():
    csvLoc = str(input("Enter CSV Location"))
    df = pd.read_csv(csvLoc)
    # seid, city,studentName, grade, fathersName, examCenter, examDate, shift, reportingTime,schoolName,volunteerName
    for cardType,seid, city, studentName, grade, fathersName, examCenter, examDate, shift, reportingTime, schoolName, schoolAddress, volunteerName in zip(
            df['CardType'], df['SEID'], df['City'], df['Name'], df['Grade'], df['Father Name'], df['Center'],
            df['Exam Date'], df['Shift'], df['ExamTime'], df['School Name'], df['School Address'],
            df['Volunteer Email']):
        try:
            blankPrefilled(cardType,seid,city,studentName,schoolName,schoolAddress,fathersName,examCenter,examDate,grade,shift,reportingTime,volunteerName)
        except Exception as e:
            print(f"Something went wrong with {seid} with {traceback.print_exc()}")

def saveHTML2Pdf(htmlLoc):
    pdfkit.from_file(htmlLoc, 'admitCardTest.pdf')


def mergeInstruction(loc1, loc2, roll):
    merger = PdfFileMerger()
    merger.append(loc1)
    merger.append(loc2)
    merger.write(f"{roll}.pdf")

def blankPrefilled(cardType,seid,city,studentName,schoolName,schoolAddress,fathersName,examCenter,examDate,grade,shift,reportingTime,volunteerName):
    if cardType == "Hope":
        template_loader = jinja2.FileSystemLoader(searchpath="./")
        template_env = jinja2.Environment(loader=template_loader)
        template_file = "Templates/Sitare Admit Card-Hindi-Hope-V2.3T/SitareAdmitCardHindiHopeV2.3T.html"
        template = template_env.get_template(template_file)
        output_text = template.render(
            seid=seid,
            allowance=allowanceCity[city],
            studentName=studentName,
            currentSchool=schoolName,
            schoolAddress=schoolAddress,
            fathersName=fathersName,
            examCenter=examCenter,
            examDate=examDate,
            grade=grade,
            shift=shift,
            reportingTime=reportingTime,
            qrCode=createQRPng(seid),  # createQRPng(idNumber)
            volunteerId=volunteerName
        )

        html_path = f'blankPrefilledHtml/{cardType}-{volunteerName}-{seid}.html'
        html_file = open(html_path, 'w')
        html_file.write(output_text)
        html_file.close()
        print(f"Hope html with seid {seid} saved successfully")
        pdfkit.from_file(html_path, f'blankPrefilled/{cardType}-{volunteerName}-{seid}.pdf', options=options)
        print("PDF Saved Successful")

    elif cardType == "Growth":
        template_loader = jinja2.FileSystemLoader(searchpath="./")
        template_env = jinja2.Environment(loader=template_loader)
        template_file = "Templates/Sitare Admit -GrowthV2.3 T/SitareAdmitGrowthV2.3T.html"
        template = template_env.get_template(template_file)
        output_text = template.render(
            seid=str(seid),
            allowance=allowanceCity[city],
            studentName=str(studentName),
            currentSchool=str(schoolName),
            schoolAddress=str(schoolAddress),
            fathersName=str(fathersName),
            examCenter=str(examCenter),
            examDate=str(examDate),
            grade=str(grade),
            shift=str(shift),
            reportingTime=str(reportingTime),
            qrCode=createQRPng(seid),  # createQRPng(idNumber)
            volunteerId=volunteerName,
        )
        html_path = f'blankPrefilledHtml/{cardType}-{volunteerName}-{seid}.html'
        html_file = open(html_path, 'w')
        html_file.write(output_text)
        html_file.close()
        print(f"Growth html with seid {seid} saved successfully")
        pdfkit.from_file(html_path, f'blankPrefilled/{cardType}-{volunteerName}-{seid}.pdf', options=options)
        print("PDF Saved Successful")


# saveHtmlToPdf()
# convert_html_to_pdf(source_html,output_filename)
# mergeInstruction("Instructions/Sitare Admit Card -G9-Instruction.pdf","Instructions/Sitare Admit Card-G6-Instruction.pdf","mergerdoc")
if __name__ == '__main__':
    print("1. For Growth Admit Card\n2. For Blank Admit Card\n3. Blank Prefilled Card")
    n = int(input("Enter your choice"))
    if n == 1:
        iterateCsv()
    elif n == 2:
        iterateBlankCsv()
    elif n ==3:
        iterateBlankPreFilledCsv()
    else:
        print("Wrong choice")