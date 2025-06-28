from robocorp import browser
from robocorp.tasks import task

from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.Tables import Tables, Table
from RPA.Archive import Archive

@task
def order_robots_from_robot_spare_bin() -> None:
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    orders = get_orders_from_csv_file()
    if orders is None or orders.size < 1:
        print("Failed to get the list of orders")

        return

    browser.configure(slowmo = 100)
    open_robot_order_website()

    for order in orders:
        try:
            order_id = order['Order number']
            close_annoying_modal()

            fill_the_form(order)
            receipt_path = store_receipt_as_pdf(order_id)
            if not receipt_path:
                print("There was an error on the website while filling order:", order_id, ', will retry...')
                continue

            screenshot_path = screenshot_robot(order_id)
            embed_screenshot_to_receipt(screenshot_path, receipt_path)
            open_add_another_order_page()
        except Exception as error:
            print(f"An unknown error occurred while processing order {order['Order number']}: {error}")

    archive_receipts()


def get_orders_from_csv_file() -> Table | None:
    """Get orders from a CSV file"""
    try:
        orders_csv_url = "https://robotsparebinindustries.com/orders.csv"

        http = HTTP()
        http.download(orders_csv_url, overwrite = True)

        table = Tables()
        return table.read_table_from_csv('orders.csv')
    except Exception:
        pass


def open_robot_order_website() -> None:
    """Open the RobotSpareBin website"""
    robot_spare_bin_url = "https://robotsparebinindustries.com/#/robot-order"
    browser.goto(robot_spare_bin_url)


def close_annoying_modal() -> None:
    """Close the modal by clicking the 'OK' button"""
    try:
        order_page = browser.page()
        if order_page.is_visible("button:text('OK')"):
            order_page.click("button:text('OK')")
    except:
        pass


def fill_the_form(order: list[dict[str, str]]) -> None:
    """Fill in the order form and click the 'order' button"""
    try:
        order_page = browser.page()

        order_page.select_option("select#head", str(order['Head']))
        order_page.set_checked(f"input#id-body-{order['Body']}", True)
        order_page.fill("input[placeholder='Enter the part number for the legs']", order['Legs'])
        order_page.fill("input#address", order['Address'])

        order_page.click("button#preview")
        order_page.click("button#order")

        retry_count = 0
        page = browser.page()
        while page.is_visible("div[class='alert alert-danger']") and retry_count < 5:
            retry_count += 1
            error_message = page.locator("div[class='alert alert-danger']").inner_text()
            print(f"A system error occurred on the website filling submitting order {order['Order number']}: {error_message}")

            order_page.click("button#order")
    except:
        pass


def store_receipt_as_pdf(order_id: str) -> str:
    """Get the HTML of the order receipt and bot image and save them to a pdf file"""
    order_complete_page = browser.page()
    if not order_complete_page.is_visible("div#order-completion"):
        return

    pdf = PDF()
    pdf_file_path = f"output/receipts/order-{order_id}-receipt.pdf"
    order_receipt_html = order_complete_page.locator("div#order-completion").inner_html()
    pdf.html_to_pdf(order_receipt_html, pdf_file_path)

    return pdf_file_path


def screenshot_robot(order_id: str) -> str:
    """Take a screenshot of the ordered robot"""
    order_complete_page = browser.page()

    image_file_path = f"output/receipts/order-{order_id}-image.png"
    order_complete_page.locator("div#robot-preview-image").screenshot(path = image_file_path)

    return image_file_path


def embed_screenshot_to_receipt(screenshot_path, pdf_file_path) -> None:
    """Embed the screenshot of the robot to the pdf receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path = screenshot_path,
        source_path = pdf_file_path,
        output_path = pdf_file_path
    )


def open_add_another_order_page() -> None:
    """Click 'ORDER ANOTHER BOT' button to open the bot order page"""
    page = browser.page()
    page.click("button#order-another")


def archive_receipts():
    """Archive The folder with the receipts"""
    archive = Archive()
    archive.archive_folder_with_zip('output/receipts', 'output/receipts.zip', recursive = True)
