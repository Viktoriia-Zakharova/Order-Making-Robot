from robocorp.tasks import task
from robocorp import browser 

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import shutil
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    clean_up()
    browser.configure(
        slowmo=50,
    )
    open_robot_order_website()
    orders = get_orders()
    fill_the_form(orders)
    archive()

def open_robot_order_website():
    """Opens the order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    orders = Tables().read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])
    return orders

def fill_the_form(orders):
    page = browser.page()
    n = Tables().get_table_dimensions(orders)[0]
    for i in range(n):
        order = Tables().get_table_row(orders,i)
        head = order["Head"]
        body = order["Body"]
        legs = order["Legs"]
        address = order["Address"]
        order_number = order["Order number"]

        page = browser.page()
        close_annoying_modal()
        page.select_option("#head", str(head))
        body_string = "id-body-"+str(body)
        page.click("#"+body_string)
        page.fill("input[placeholder='Enter the part number for the legs']", str(legs))
        page.fill("#address", address)
        
        screenshot_path = screenshot_preview(order_number)

        while True:
            page = browser.page()
            page.click("#order")

            order_another = page.query_selector("#order-another")
            if order_another:
                page = browser.page()
                pdf_path = get_pdf_receipt(order_number)
                embed_screenshot_to_pdf(screenshot_path, pdf_path)
                page.click("#order-another")
                break
        
    
def screenshot_preview(order_number):
    page = browser.page()
    page.click("#preview")
    screenshot_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def get_pdf_receipt(order_number):
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = "output/receipts/{0}.pdf".format(order_number)
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def embed_screenshot_to_pdf(screenshot_path, pdf_path):
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, 
                                   source_path=pdf_path, 
                                   output_path=pdf_path)
    
def archive():
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def clean_up():
    if os.path.isdir("./output/receipts"):
        shutil.rmtree("./output/receipts")
    if os.path.isdir("./output/screenshots"):
        shutil.rmtree("./output/screenshots")