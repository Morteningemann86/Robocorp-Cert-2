from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
import os
from pathlib import Path
import zipfile

from resources.variables import *
from resources.selectors import *

pdf = PDF()


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    init()
    open_robot_order_website()
    orders = get_orders()

    for order in orders:
        try:
            close_annoying_modal_popup()
            fill_form(order)
            click_preview_button()

            for retry in range(MAX_RETRIES):
                try:
                    click_order_button()
                    check_server_error_message()
                    break
                except Exception as e:
                    print(
                        f"Error ordering spareparts, retrying {retry+1}/{MAX_RETRIES}: {e}"
                    )

            path_pdf = store_receipt_as_pdf(order["Order number"])
            path_png = store_receipt_as_png(order["Order number"])
            append_png_to_pdf(path_pdf, path_png)

            click_order_another_robot_button()

        except Exception as e:
            print(f'Error in processing order {order["Order number"]}: {e}')

    archive_receipts(DIR_RECEIPT_FOLDER, PATH_RECEIPT_ZIP)
    print("stop")


def init():
    """Load configuration and initialize settings"""
    browser.configure(headless=False, slowmo=200)
    global page
    page = browser.page()


def open_robot_order_website():
    """Open robot order website"""
    browser.goto(URL_ROBOT_ORDER_WEBSITE)


def get_orders():
    """Get orders from CSV file and return as a table"""
    http = HTTP()
    tables = Tables()
    http.download(
        url=URL_CSV_FILE_ORDERS, target_file=PATH_DOWNLOAD_FILE_ORDERS, overwrite=True
    )
    table_orders = tables.read_table_from_csv(
        path=PATH_DOWNLOAD_FILE_ORDERS, header=True
    )
    return table_orders


def close_annoying_modal_popup():
    """Close annoying modal popup"""
    page.click('button:text("OK")')


def fill_form(row):
    """Fill the form with the order details and submit"""
    # TODO: add the order details to the form
    order_num = row["Order number"]
    print(f"Order number: {order_num}")

    select_head(row["Head"])
    select_body(row["Body"])
    select_legs(row["Legs"])
    select_address(row["Address"])


def select_head(input):
    """Select head from dropdown"""
    print(f"head input: {input}")
    page.select_option(DROPDOWN_HEAD, input)


def select_body(input):
    """Select body from radio buttons"""
    print(f"body input: {input}")
    radio_body = f'//input[@type="radio" and @value="{input}"]'
    page.click(radio_body)


def select_legs(input):
    """Select legs from input field"""
    print(f"legs input: {input}")
    page.fill(INPUT_LEGS, input)


def select_address(input):
    """Select address from input field"""
    print(f"Address input {input}")
    page.fill(INPUT_ADDRESS, input)


def click_preview_button():
    """Click preview button after filling the form"""
    page.click(BUTTON_PREVIEW)


def click_order_button():
    """Click order button after filling the form"""
    page.click(BUTTON_ORDER)


def click_order_another_robot_button():
    """Click order another robot button after ordering a robot"""
    page.click(BUTTON_ORDER_ANOTHER_ROBOT)


def check_server_error_message():
    """Check if there is a server error message and raise exception if there is"""
    if page.is_visible(ALERT_MESSAGE):
        element = page.locator(ALERT_MESSAGE)
        error_message = element.inner_text()
        print(f"Error message: {error_message}")
        raise Exception(f"Server error message: {error_message}")


def store_receipt_as_pdf(order_num):
    """Store receipt as PDF file"""
    path_pdf = f"{DIR_RECEIPT_FOLDER}/receipt-{order_num}.pdf"
    receipt_html = page.locator(ROBOT_RECEIPT).inner_html()
    pdf.html_to_pdf(receipt_html, path_pdf)
    return path_pdf


def store_receipt_as_png(order_num):
    """Store receipt screenshot as PNG file"""
    path_png = f"{DIR_RECEIPT_FOLDER}/receipt-{order_num}.png"
    element = page.locator(PREVIEW_IMAGE)
    element.screenshot(path=path_png)
    return path_png


def append_png_to_pdf(pdf_path, png_path):
    """
    Appends a PNG image to an existing PDF file and saves the result with a new filename.

    :param pdf_path: The file path of the existing PDF to which the PNG will be appended.
    :param png_path: The file path of the PNG image to append to the PDF.
    """
    # Create a new filename for the output PDF with the appended PNG
    output_filename = f"{os.path.splitext(pdf_path)[0]}_with_image.pdf"

    # List of files to append: the specified PDF and PNG
    list_of_files = [
        pdf_path,  # The PDF file to which you want to append the PNG
        f"{png_path}:align=center",  # The PNG image you want to append, with alignment
    ]

    # Call the method to append the PNG to the PDF
    pdf.add_files_to_pdf(
        files=list_of_files, target_document=output_filename, append=True
    )

    # Delete the original PDF file
    os.remove(pdf_path)

    # Rename the new PDF file to the original PDF filename
    Path(output_filename).rename(pdf_path)


def archive_receipts(directory_path, output_zip_path):
    def zipdir(path, ziph):
        for root, dirs, files in os.walk(path):
            for file in files:
                ziph.write(
                    os.path.join(root, file),
                    os.path.relpath(os.path.join(root, file), os.path.join(path, "..")),
                )

    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipdir(directory_path, zipf)
