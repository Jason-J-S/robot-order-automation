
"""
Robot Order Automation
======================

Orders robots from RobotSpareBin Industries Inc.
Saves the order HTML receipt as a PDF file.
Saves the screenshot of the ordered robot.
Embeds the screenshot of the robot to the PDF receipt.
Creates ZIP archive of the receipts and the images.

Author: Robocorp Tutorial
Website: https://robotsparebinindustries.com/#/robot-order
CSV Orders: https://robotsparebinindustries.com/orders.csv
"""

import os
import zipfile
from pathlib import Path
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,  # Slow down actions for better reliability
        headless=False  # Set to True for production
    )

    # Create output directories
    create_output_directories()

    # Open the robot order website
    open_robot_order_website()

    # Download and get the orders
    orders = get_orders()

    # Fill the form for each order
    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
        preview_the_robot()
        submit_the_order()

        # Store order data
        order_number = row["Order number"]
        pdf_file = store_receipt_as_pdf(order_number)
        screenshot = screenshot_robot(order_number)
        embed_screenshot_to_receipt(screenshot, pdf_file)

        # Go to order another robot
        go_to_order_another_robot()

    # Create ZIP archive with all PDFs
    archive_receipts()

    # Close browser
    browser.context().close()


def create_output_directories():
    """Creates the necessary output directories"""
    Path("output/receipts").mkdir(parents=True, exist_ok=True)
    Path("output/screenshots").mkdir(parents=True, exist_ok=True)


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def get_orders():
    """Downloads the orders file, reads it as a table, and returns the result"""
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv", 
        target_file="orders.csv",
        overwrite=True
    )

    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv", header=True)
    return orders


def close_annoying_modal():
    """Closes the annoying modal that appears when opening the website"""
    page = browser.page()
    try:
        # Wait for modal to appear and click OK
        page.wait_for_selector("text=OK", timeout=5000)
        page.click("text=OK")
    except Exception:
        # Modal might not appear, continue
        pass


def fill_the_form(order):
    """Fills the order form with data from the given order"""
    page = browser.page()

    # Select head
    page.select_option("#head", str(order["Head"]))

    # Select body using radio button
    body_selector = f"#id-body-{order['Body']}"
    page.click(body_selector)

    # Enter legs - find the input field for legs
    legs_input = page.locator("input[placeholder='Enter the part number for the legs']")
    legs_input.clear()
    legs_input.fill(str(order["Legs"]))

    # Enter address
    page.fill("#address", str(order["Address"]))


def preview_the_robot():
    """Clicks the preview button"""
    page = browser.page()
    page.click("#preview")

    # Wait for preview to load
    page.wait_for_selector("#robot-preview-image", timeout=10000)


def submit_the_order():
    """Submits the order. Retries until success due to website flakiness"""
    page = browser.page()

    # Keep trying to submit until success
    while True:
        # Click the order button
        page.click("#order")

        # Check if order was successful by looking for receipt
        if page.locator("#receipt").is_visible():
            print("Order submitted successfully!")
            break

        # Check if there's an error and retry
        if page.locator(".alert-danger").is_visible():
            print("Error submitting order, retrying...")
            continue

        # Add a small delay before retrying
        page.wait_for_timeout(1000)


def store_receipt_as_pdf(order_number):
    """Takes a screenshot of the receipt and stores it as a PDF file"""
    page = browser.page()

    # Wait for receipt to be visible
    page.wait_for_selector("#receipt", timeout=10000)

    # Get the receipt HTML
    receipt_html = page.locator("#receipt").inner_html()

    # Create PDF from HTML
    pdf = PDF()
    pdf_path = f"output/receipts/receipt_{order_number}.pdf"

    # Convert HTML to PDF
    pdf.html_to_pdf(receipt_html, pdf_path)

    return pdf_path


def screenshot_robot(order_number):
    """Takes a screenshot of the robot preview image"""
    page = browser.page()

    # Wait for robot preview to be visible
    page.wait_for_selector("#robot-preview-image", timeout=10000)

    # Take screenshot of the robot image
    screenshot_path = f"output/screenshots/robot_{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)

    return screenshot_path


def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """Embeds the screenshot into the PDF receipt"""
    pdf = PDF()

    # Open the PDF and add the screenshot to it
    files_to_add = [
        pdf_path + ":1",  # First page of the PDF
        screenshot_path   # Screenshot to append
    ]

    pdf.add_files_to_pdf(
        files=files_to_add,
        target_document=pdf_path
    )


def go_to_order_another_robot():
    """Clicks the 'Order another robot' button"""
    page = browser.page()
    page.click("#order-another")


def archive_receipts():
    """Creates a ZIP archive of all PDF receipts in the output directory"""
    archive = Archive()
    archive.archive_folder_with_zip(
        folder="output/receipts", 
        archive_name="output/receipts.zip"
    )
    print("Created ZIP archive: output/receipts.zip")


# Additional helper function for debugging
def log_page_content():
    """Helper function to log current page content for debugging"""
    page = browser.page()
    print("Current page title:", page.title())
    print("Current URL:", page.url())


if __name__ == "__main__":
    order_robots_from_RobotSpareBin()
