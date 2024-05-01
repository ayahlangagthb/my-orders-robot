from robocorp.tasks import task
from robocorp import browser
import requests
from RPA.Tables import Tables
import os
import zipfile

from robot.api import logger
from RPA.Browser.Selenium import Selenium
from robocorp.log import exception
from RPA.PDF import PDF
import time
from RPA.Browser.Selenium import Selenium

@task
def order_robots_from_RobotSpareBin():
    """Orders robots from RobotSpareBin Industries Inc.

    Ready to build the robot? Know the rules
    My dear, without rules there's only chaos. - Star Wars: The Clone Wars: Senate Murders (2010)

    The robot should use the orders file (.csv ) and complete all the orders in the file.

    Only the robot is allowed to get the orders file. You may not save the file manually on your computer.
    The robot should save each order HTML receipt as a PDF file.
    The robot should save a screenshot of each of the ordered robots.
    The robot should embed the screenshot of the robot to the PDF receipt.
    """

    browser.configure(
        slowmo=700,
    )

    open_robot_order_website("https://robotsparebinindustries.com/#/robot-order")
    csv_file_path = download_csv_file("https://robotsparebinindustries.com/orders.csv")
    orders = get_orders(csv_file_path)
    for order in orders:
        # Log each order row
        logger.info(f"Order: {order}")
        order_number = order["Order number"]
        fill_the_form(order)
        try:
            pdf_file = store_receipt_as_pdf(order_number)
            screenshot_path = screenshot_robot(order_number)
            embed_screenshot_to_receipt(screenshot_path, pdf_file)
        except Exception as e:
            logger.info(f"Error processing order {order_number}: {e}")
        finally:
            return_to_order_form()
            close_annoying_modal()
            time.sleep(2)  # Wait for page to load before processing next order
    #Zip all generated order receipt files to an Archive        
    archive_receipts()
    # Log each order row
    logger.info(f"Done: processing Orders :) ")

def open_robot_order_website(orders_url):
    """Open order form (url)"""
    browser.goto(orders_url)
    close_annoying_modal()

def close_annoying_modal():
    """Wait for Accept Constitutional rights pop up and click ok"""
    page = browser.page()
    try:
        page.wait_for_selector("text='OK'", timeout=6000)
        page.click("text='OK'")
    except Exception as e:
        logger.info(f"Error closing modal: {e}")

def download_csv_file(csv_url):
    """Downloads CSV file from the given URL to the robot's workspace."""
    try:
        response = requests.get(csv_url, timeout=6)
        response.raise_for_status()  # Ensure download succeeded

        workspace_path = os.getenv("ROBOT_WORKSPACE", ".")
        destination_path = os.path.join(workspace_path, "orders.csv")

        with open(destination_path, 'wb') as file:
            file.write(response.content)

        return destination_path
    except Exception as e:
        logger.info(f"Error downloading CSV file: {e}")
        return None

def get_orders(csv_file_path):
    """Reads orders from a CSV file into a table."""
    try:
        tables = Tables()
        orders = tables.read_table_from_csv(
            path=csv_file_path,
            header=True,  # Assumes the first row is the header
            dialect="excel",  # Adjust as necessary
            encoding="utf-8"  # Adjust as necessary
        )
        return orders
    except Exception as e:
        logger.info(f"Error reading orders from CSV: {e}")
        return []

def fill_the_form(order):
    """
    Fills a single form with the information of a single row in the Excel
    """
    page = browser.page()
    try:
        # Select the "Head" part from the dropdown
        page.select_option("select[name='head']", str(order["Head"]))
        body_selector = f"input[name='body'][value='{order['Body']}']"
        page.wait_for_selector(body_selector, timeout=6000)
        page.click(body_selector)
        # Fill in the "Legs" part number using class name and placeholder
        legs_input_selector = "input.form-control[placeholder='Enter the part number for the legs']"
        page.fill(legs_input_selector, str(order["Legs"]))
        # Enter shipping address
        page.fill("input[id='address']", order["Address"])

        # Click the preview button using its 'id'
        page.click("button[id='preview']")

        # Wait for the order button to be available
        page.wait_for_selector("button[id='order']", timeout=6000)
        submit_order()
    except Exception as e:
        logger.info(f"Error filling the form: {e}")

def submit_order():   
    page = browser.page()   
    try:
        # Submit the form by Clicking the order button using its 'id'
        page.click("button[id='order']")
        raise ValueError("Error Submitting Order")
    except Exception as e:
        # Click the link using its class and href attributes
        page.click("a.nav-link.active[href='#/robot-order']")
        exception("An error occurred during execution:", e)

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    try:
        screenshot = f"output/screenshots/order_{order_number}_summary.png"
        page.screenshot(path=screenshot)
        return screenshot
    except Exception as e:
        logger.info(f"Error taking screenshot: {e}")
        return None

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    try:
        # Wait for the specific element to be present on the page
        page.wait_for_selector("#receipt", timeout=10000)
        # Take the embedded data to a pdf
        order_results_html = page.locator("#receipt").inner_html()

        pdf = PDF()
        pdf_file = f"output/receipts/order_{order_number}_results.pdf"
        pdf.html_to_pdf(order_results_html, pdf_file)
        return pdf_file
    except Exception as e:
        logger.info(f"Error storing receipt as PDF: {e}")
        return None

from RPA.PDF import PDF

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embed the robot screenshot to the receipt PDF file"""
    pdf = PDF()

    try:
        # Specify the screenshot with desired properties, e.g., alignment
        files_to_add = [f"{screenshot}:align=center"]
        
        # Append the screenshot to the existing PDF file
        pdf.add_files_to_pdf(files=files_to_add, target_document=pdf_file, append=True)
    except Exception as e:
        logger.info(f"Error occurred while embedding screenshot to receipt PDF: {e}")
    finally:
        # It's good practice to close the PDF to ensure all changes are saved properly
        pdf.close_pdf()

def return_to_order_form():
    """Click the 'Order another robot' button using its 'id'. Reload main page if button not found."""
    page = browser.page()
    try:
        # Attempt to click the button by its ID
        page.click("button[id='order-another']", timeout=60000)
    except Exception as e:
        # Log the exception
        logger.info(f"Error occurred while returning to order form: {e}")
        # Call archive_receipts function in the catch block
        archive_receipts()

@task
def archive_receipts():
    """Create a ZIP file of receipt PDF files."""
    receipts_directory = "output/receipts"
    zip_file_name = "receipts.zip"

    try:
        create_zip(receipts_directory, zip_file_name)
        logger.info(f"Receipts archived successfully as {zip_file_name}")
    except Exception as e:
        logger.info(f"Error archiving receipts: {e}")

def create_zip(directory_to_zip, zip_file_name):
    """Create a ZIP file from a directory."""
    with zipfile.ZipFile(zip_file_name, 'w') as zipf:
        for root, _, files in os.walk(directory_to_zip):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=os.path.relpath(os.path.join(root, file), directory_to_zip))

