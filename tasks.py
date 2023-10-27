import time
from RPA.HTTP import HTTP
from robocorp.tasks import task
from robocorp import browser
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
import pandas as pd
import os
import zipfile
import shutil


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
        screenshot="only-on-failure",
        headless=False,
        slowmo=100,
    )

    # Open the page and Click the button with the class name "btn btn-dark"
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click(".btn.btn-dark")

    # download_excel_file():
    """Downloads CSV file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    # fill_form_with_CSV_data():
    """Read data from csv and fill in the orders form"""
    tables = Tables()
    table = tables.read_table_from_csv("orders.csv")
    for row in table:
        fill_and_submit_sales_form(row)
    zip_and_move_files("output", "Orders")


def fill_and_submit_sales_form(row):
    """Fills in the sales form with data from each row and submits it"""
    page = browser.page()

    order_number = row["Order number"]
    head_option = row["Head"]
    body_radio_value = str(row["Body"])  # Convert to string for comparison
    legs = row["Legs"]
    address = row["Address"]

    # Select the Head option
    page.select_option("#head", head_option)

    # Click the Body radio button based on the "value" attribute
    page.click(
        f"//input[@type='radio' and @name='body' and @value='{body_radio_value}']"
    )

    # Fill in the Legs field
    page.fill('input[type="number"]', legs)

    # Fill in the Shipping Address field
    page.fill('input[type="text"]', address)

    # Preview the form
    page.click("text=Preview")
    max_attempts = 5  # Set the maximum number of attempts
    success = False

    for attempt in range(max_attempts):
        try:
            page.click("#order")
            page.wait_for_selector(
                "xpath=//div[contains(@class, 'alert') and contains(@class, 'alert-success')]",
                timeout=10000,
            )
            print("Submit operation succeeded!")
            success = True

            # Save the receipt as PDF in folder
            page.screenshot(path=f"output/Order Number {order_number}.png")
            order_number_html = page.locator("#receipt").inner_html()
            pdf = PDF()
            pdf.html_to_pdf(
                order_number_html, f"output/Order Number {order_number}.pdf"
            )
            page.click("#order-another")
            page.click(".btn.btn-dark")
            break
        except:
            page.wait_for_selector(
                "xpath=//div[contains(@class, 'alert') and contains(@class, 'alert-danger')]",
                timeout=5000,
            )  # Replace with the actual error message selector
            print(f"Submit failed on attempt {attempt + 1}")
            if attempt < max_attempts - 1:  # If this wasn't the last attempt
                time.sleep(4)
                page.click("#order")
                try:
                    page.wait_for_selector(
                        "xpath=//div[contains(@class, 'alert') and contains(@class, 'alert-success')]",
                        timeout=10000,
                    )
                    print("Submit operation succeeded!")
                    success = True
                    # Take screenshot and save as pdf
                    page.screenshot(path=f"output/Order Number {order_number}.png")
                    order_number_html = page.locator("#receipt").inner_html()
                    pdf = PDF()
                    pdf.html_to_pdf(
                        order_number_html, f"output/Order Number {order_number}.pdf"
                    )
                    page.click("#order-another")
                    page.click(".btn.btn-dark")
                    break
                except:
                    print("Submit failed, trying again.")
            else:
                print("Max attempts reached, moving on to next order.")
                break


def zip_and_move_files(folder_path, zip_file_name):
    """Zip all .pdf and .png files in a folder and move the resulting zip file to the same location."""
    zip_file_name = zip_file_name + ".zip"  # Add the .zip extension to the file name
    zip_file_path = os.path.join(folder_path, zip_file_name)
    with zipfile.ZipFile(zip_file_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                _, ext = os.path.splitext(file)
                if ext.lower() in [".pdf", ".png"]:
                    file_path = os.path.join(root, file)
                    zipf.write(
                        file_path, arcname=os.path.relpath(file_path, folder_path)
                    )


# Specify the folder you want to zip files from
folder_path = "output"
time.sleep(10)
