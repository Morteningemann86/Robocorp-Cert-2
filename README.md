# RobotSpareBin Industries Inc. Order Automation Script

This Python script automates the process of ordering robots from RobotSpareBin Industries Inc. It utilizes various RPA (Robotic Process Automation) libraries to interact with web pages, handle files, and process data. The script performs the following tasks:

- Opens the RobotSpareBin order website.
- Retrieves order details from a CSV file.
- For each order, fills out the form on the website, submits it, and handles possible retries for errors.
- Saves a PDF receipt and a PNG screenshot of each order.
- Embeds the screenshot into the corresponding PDF receipt.
- Archives all receipts and screenshots into a ZIP file.

## Libraries and External Files

- `robocorp.browser`, `RPA.HTTP`, `RPA.Excel.Files`, `RPA.Tables`, `RPA.PDF`: Libraries from RoboCorp RPA Framework for browser automation, HTTP requests, Excel file handling, table manipulation, and PDF processing.
- `os`, `pathlib`, `zipfile`: Standard Python libraries for operating system interaction, path manipulation, and handling ZIP files.
- `resources.variables`, `resources.selectors`: External Python modules containing constants used in the script (e.g., URLs, file paths, and CSS selectors).

## Functions

### `order_robots_from_RobotSpareBin()`

Main function that orchestrates the robot ordering process. It initializes the browser, navigates to the order website, and processes each order found in the CSV file. After processing orders, it archives the receipts and screenshots.

### `init()`

Initializes browser settings and global variables.

### `open_robot_order_website()`

Navigates to the RobotSpareBin order website.

### `get_orders()`

Retrieves order details from a CSV file and returns them as a list of dictionaries.

### `close_annoying_modal_popup()`

Closes any modal popups that might obstruct the ordering process.

### `fill_form(row)`

Fills the order form with details from a single order row.

### `select_head(input)`, `select_body(input)`, `select_legs(input)`, `select_address(input)`

These functions fill in the respective sections of the order form based on the order details.

### `click_preview_button()`

Clicks the preview button after the order form is filled.

### `click_order_button()`

Submits the order form by clicking the order button.

### `click_order_another_robot_button()`

Clicks the button to initiate another order after the current one is completed.

### `check_server_error_message()`

Checks for server error messages and raises an exception if any are found.

### `store_receipt_as_pdf(order_num)`, `store_receipt_as_png(order_num)`

Saves the order receipt as a PDF and a PNG screenshot, respectively.

### `append_png_to_pdf(pdf_path, png_path)`

Embeds the PNG screenshot into the PDF receipt.

### `archive_receipts(directory_path, output_zip_path)`

Archives all receipts and screenshots into a ZIP file.

## Usage

To run the script, execute it from a command line or an IDE that supports Python execution. Ensure that all the necessary libraries are installed and that the `resources` directory contains the correct `variables.py` and `selectors.py` files with up-to-date values.
