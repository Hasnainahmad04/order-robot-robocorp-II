import shutil
import os


from robocorp.tasks import task
from robocorp import browser

from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.Archive import Archive
from RPA.HTTP import HTTP
from RPA.PDF import PDF

order_page_url = "https://robotsparebinindustries.com/#/robot-order"
order_file_download_url = "https://robotsparebinindustries.com/orders.csv"
pdf = PDF()
page = browser.page()

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    browser.configure(slowmo=200)
    clean_output_directories()
    open_robot_order_website()
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
    archive_receipts()
    clean_up()

def open_robot_order_website():
    """Navigates to the orders website"""
    browser.goto(order_page_url)

def close_annoying_modal():
    """Closes the annoying modal if present"""
    try:
        page.click('text=OK')
    except Exception as e:
        print("Modal already closed")

def download_orders_file():
    """Downloads the orders CSV file"""
    http = HTTP()
    http.download(url=order_file_download_url, target_file='orders.csv', overwrite=True)

def get_orders():
    """Reads the orders from the downloaded CSV file"""
    download_orders_file()
    return Tables().read_table_from_csv(path='orders.csv', header=True)


def fill_the_form(order):
    """Fills the order form"""
    close_annoying_modal()
    page.select_option('#head', str(order['Head']))
    page.set_checked(f"#id-body-{order['Body']}", True)
    page.fill("input[placeholder='Enter the part number for the legs']", str(order['Legs']))
    page.fill('#address', order['Address'])
    page.click('#order')
    while not page.locator("#receipt").is_visible():
        page.click('#order')
    else:
        receipt_path = store_receipt_as_pdf(order['Order number'])
        screenshot_path = take_robot_screenshot(order['Order number'])
        embed_screenshot_to_order_receipt(receipt=receipt_path, screenshot=screenshot_path)
        page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Stores the receipt as a PDF file"""
    path = 'output/receipts/order_{order_number}.pdf'.format(order_number=order_number)
    receipt_html  = page.locator('#receipt').inner_html()
    pdf.html_to_pdf(receipt_html, path)
    return path

def take_robot_screenshot(order_number):
    """Takes a screenshot of the robot and saves it as a file"""
    path = 'output/robot-preview-images/{order_number}.jpeg'.format(order_number=order_number)
    page.locator('#robot-preview-image').screenshot(path=path)
    return path

def embed_screenshot_to_order_receipt(receipt, screenshot):
    """Embeds the screenshot into the PDF receipt"""
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=receipt,
        append=True
    )

def archive_receipts():
    """Creates a ZIP archive of the receipts"""
    archive = Archive()
    archive.archive_folder_with_zip(folder='output/receipts', archive_name='output/receipts.zip', recursive=True)

def clean_output_directories():
    """Cleans up the output directories before processing"""
    shutil.rmtree('output/receipts', ignore_errors=True)
    shutil.rmtree('output/robot-preview-images', ignore_errors=True)
    os.makedirs('output/receipts')
    os.makedirs('output/robot-preview-images')

def clean_up():
    """Cleans up the output directories after processing"""
    shutil.rmtree(path='output/receipts')
    shutil.rmtree(path='output/robot-preview-images')


