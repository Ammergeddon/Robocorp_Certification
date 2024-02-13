import os

from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

out_dir = os.path.join( os.getcwd(), "outdir" )
screenshot_dir = os.path.join( out_dir, "screenshots" )
receipts_dir = os.path.join( out_dir, "receipts" )

@task
def order_robots_from_RobotSpareBin():
    """
        Orders robots from RobotSpareBin Industries Inc.
        Saves the order HTML receipt as a PDF file.
        Saves the screenshot of the ordered robot.
        Embeds the screenshot of the robot to the PDF receipt.
        Creates ZIP archive of the receipts and the images.
    """

    browser.configure( headless=True )
    browser.goto( "https://robotsparebinindustries.com/#/robot-order" )

    for order in download_orders():
        close_modal()
        fill_form_and_send( order )
        store_receipt_as_pdf( order["Order number"] )
        return_to_order_page()

    archive_receipts()
    
def close_modal():
    """
        Deals with the initial modal of the website.
        Grants all rights if it is visible, does nothing if it isn't
    """
    page = browser.page()
    if page.is_visible( "button:text('OK')" ): 
        page.click( "button:text('OK')" )

def download_orders():
    http = HTTP()
    http.download( url="https://robotsparebinindustries.com/orders.csv", overwrite=True )

    library = Tables()
    return library.read_table_from_csv( "orders.csv" )

def fill_form_and_send(order: dict):
    """
        Fills the order form and sends the order. Slight error handling, in case the order isn't send the first time

        order: A complete row from the orders.csv
    """
    page = browser.page()
    page.select_option( "#head", order["Head"] )
    page.click( f"#id-body-{order['Body']}" )
    page.fill( "//input[@placeholder='Enter the part number for the legs']", order["Legs"] )
    page.fill( "#address", order["Address"] )

    while (True):
        page.click( "#order" )
        if not page.is_visible( "//*[@class='alert alert-danger']" ): break

def store_receipt_as_pdf(order_number: str):
    """
        Stores the receipt as pdf. In the process, it takes a screenshot of the robot and
        attaches it to the end of the PDF.

        order_number: The order number displayed in the specific row of the orders.csv file
    """
    pdf = PDF()

    def take_screenshot(order_number):
        screenshot = os.path.join( screenshot_dir, f"robot-{order_number}.png" )
        page.wait_for_load_state( state="domcontentloaded" )
        page.screenshot( path=screenshot, type="png" )
        return screenshot

    def embed_screenshot_to_receipt(screenshot: str, file: str):
        target_document=file
        files=[screenshot]
        pdf.add_files_to_pdf( files=files, target_document=target_document, append=True )

    page = browser.page()
    receipt_html = page.inner_html( "#receipt" )
    receipt_file = os.path.join( receipts_dir, f"robot-{order_number}.pdf" )
    pdf.html_to_pdf( receipt_html, receipt_file )

    embed_screenshot_to_receipt( take_screenshot(order_number), receipt_file )
    
def return_to_order_page():
    page = browser.page()
    page.click( "#order-another" )

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip( folder=receipts_dir, archive_name=f"{out_dir}/archive.zip" )





    