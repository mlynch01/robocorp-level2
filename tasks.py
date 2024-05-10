from robocorp import browser
from robocorp.tasks import task
from robocorp import log
from RPA.Excel.Files import Files as Excel
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from time import sleep
from pathlib import Path
import aspose.pdf as ap
from zipfile import ZipFile
import os
import csv

RSB_URL = "https://robotsparebinindustries.com/#/robot-order"
CSV_URL = "https://robotsparebinindustries.com/orders.csv"
CSV_FILENAME = "Orders.csv"

@task
def levelTwoTraining():
    """
    Level 2 of the Robocorp Training program. This process accesses the rsb website and orders robots.

    Input: csv file with all orders information
    Output: ZIP file with all order receipts in PDF format
    """
    browser.configure(
        browser_engine = "chrome",
        headless = False,
    )
    browser.goto(RSB_URL)
    downloadCSVFile(CSV_URL, CSV_FILENAME)

    with open(CSV_FILENAME, newline='') as csvfile:
        orders = csv.reader(csvfile, delimiter = ",")
        next(orders, None)
        for order in orders:
            createOrder(order)
     
    with ZipFile("Output/Receipts.zip", "w") as zip:
        for receipt in os.listdir("Receipts"):
            zip.write(receipt)

def downloadCSVFile(URL, FileName):
    """
    Get CSV file from source
    """
    http = HTTP()
    http.download(URL, FileName)

def createOrder(orderInformation):
    """
    Creates the robot orders based on the CSV line
    """
    page = browser.page()
    page.click("//*[@class='btn btn-dark']")
    page.click("//button[contains(text(),'Show model')]")
    partName = page.text_content(f"//tr[td[.='{orderInformation[1]}']]/td[1]")
    page.select_option("#head", partName + " head")
    page.click(f"//label[contains(text(),'{partName}')]/input[@type='radio']")
    page.fill("//input[@placeholder='Enter the part number for the legs']", orderInformation[3])
    page.fill("#address", orderInformation[4])   
    while True:
        try:
            page.click("#order")
            saveReceipt(orderInformation[0])
        except:
            continue
        break
    


def saveReceipt(receiptID):
    """
    Gets receipt HTML and creates pdf file with recepit and robot image
    """
    pdfFileName = f"Receipts/Receipt-{receiptID}.pdf"
    page = browser.page()
    receiptHTML = page.locator("#receipt").inner_html()
    page.locator("#robot-preview-image").screenshot(path="temp.png")

    pdf = PDF()
    pdf.html_to_pdf(receiptHTML, "temp.pdf")
    mender = ap.facades.PdfFileMend()
    mender.bind_pdf("temp.pdf")
    mender.add_image("temp.png", 1, 100.0, 600.0, 200.0, 700.0)
    mender.save(pdfFileName)
    mender.close()
    page.click("#order-another")

