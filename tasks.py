from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem
from PIL import Image
import pandas as pd

PREVIEW_DIR = "output/preview/"
FINAL_DIR = "output/receipts/"
ZIP_FILE = "output/receipts.zip"

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    tearup()
    process_csv_records()
    cleanup()

def tearup():
    fs = FileSystem()
    fs.remove_file(ZIP_FILE)
    recreate_directory(PREVIEW_DIR)
    recreate_directory(FINAL_DIR)

def cleanup():
    fs = FileSystem()
    if fs.does_directory_exist(PREVIEW_DIR):
        fs.remove_directory(PREVIEW_DIR, recursive=True)

    if fs.does_directory_exist(FINAL_DIR):
        fs.remove_directory(FINAL_DIR, recursive=True)

def recreate_directory(directory_name):
    fs = FileSystem()
    if fs.does_directory_exist(directory_name):
        fs.remove_directory(directory_name, recursive=True)
    fs.create_directory(directory_name)

def process_csv_records():    
    orders = get_orders()

    if not(orders.empty) :
        print("Opening browser")
        open_robot_order_website()

    # Loop through each row in the DataFrame
    for index, row in orders.iterrows():
        # Access row elements by column name
        order_number = row['Order number']
        print(f"Processing record {index}: {order_number}")
        close_annoying_modal()
        fill_the_form(row)
        receipt_pdf = store_receipt_as_pdf(order_number)
        screenshot_png = screenshot_robot(order_number)
        #resize_robot_picture(screenshot_png)
        embed_screenshot_to_receipt(screenshot_png, receipt_pdf)
        start_new_order()
    
    print("Archive receipts")
    archive_receipts()

def get_orders():
    """Downloads orders file from the given URL"""    
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    # Load the CSV file into a DataFrame
    orders = pd.read_csv('orders.csv')
    print("Read csv file")
    return orders

def open_robot_order_website():
    """Navigates to the given URL"""    
    browser.configure(
        slowmo=100,
    )
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    """Clicks the 'Ok' button on the init popup window"""
    page = browser.page()
    page.click("button:text('OK')")
    #button = page.locator('xpath=//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
    #root > div > div.modal > div > div > div > div > div > button.btn.btn-dark
    #button.wait_for()
    
    #button.click()
    # 
    
def fill_the_form(row):
    """Enter order"""
    page = browser.page()
    page.select_option("#head", str(row['Head']))
    page.check(f'input[type="radio"][name="body"][value="{str(row["Body"])}"]')
    page.fill(f'input[placeholder="Enter the part number for the legs"]', str(row['Legs']))
    page.fill("#address", row['Address'])
    page.click("#preview")
    submit_order()

def submit_order():
    """Submit order"""
    page = browser.page()    
    page.click("#order")

    error_message_locator = page.locator(".alert.alert-danger[role='alert']")
    
    if error_message_locator.is_visible():
        submit_order()

def store_receipt_as_pdf(order_number):
    """Store receipe to a pdf file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    
    pdf = PDF()
    output_file = f'receipt_{order_number}.pdf'
    pdf.html_to_pdf(receipt_html, f'{PREVIEW_DIR}{output_file}')
    return output_file

def screenshot_robot(order_number):
    """Create screenshot of a robot and save it as a PDF file"""    
    output_file=f'{PREVIEW_DIR}robot_preview_{order_number}.png'
    browser.page().screenshot(path=output_file, clip=browser.page().query_selector("#robot-preview-image").bounding_box())
    return output_file

def resize_robot_picture(imagefilename):
    img = Image.open(imagefilename)
    base_width= 300
    wpercent = (base_width / float(img.size[0]))
    hsize = int((float(img.size[1]) * float(wpercent)))
    
    img.resize((base_width, hsize), Image.Resampling.LANCZOS).save(imagefilename)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Append screenshot to the receipe"""
    pdf = PDF()
    """
    list_of_files = [f'{PREVIEW_DIR}{pdf_file}', f'{screenshot}:align=center']        
    pdf.add_files_to_pdf(files=list_of_files, target_document=f'{FINAL_DIR}{pdf_file}')
    """
    pdf.add_watermark_image_to_pdf(image_path=screenshot, 
                                   source_path=f'{PREVIEW_DIR}{pdf_file}', 
                                   output_path=f'{FINAL_DIR}{pdf_file}')

def start_new_order():
    """Finish the process and restart order entry"""
    page = browser.page()    
    page.click("#order-another")

def archive_receipts():
    """Place all receipts into a ZIP archive"""
    zip = Archive()
    zip.archive_folder_with_zip(FINAL_DIR, ZIP_FILE, include='receipt_*.pdf')
