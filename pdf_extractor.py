from pypdf import PdfReader
import re
import tabula
import pandas
import pyperclip

import cv2
import numpy as np
from pyautogui import typewrite,click, screenshot as screen, press
from time import sleep




def invoice_text_extractor(filepath):
    reader = PdfReader(filepath)
    invoice_text = ""
    for page in reader.pages:
        invoice_text += page.extract_text() + "\n"

    # Define regex patterns to extract relevant information
    invoice_number_pattern = r"Invoice #\s*:\s*(\d+)"
    tracking_number_pattern = r"Tracking #\s*([A-Za-z0-9]+)"
    serial_number_pattern = r"(?:s/n\s*([A-Z\d]+)|Serial Number:\s*([A-Z\d\-]+)|\b[A-Z\d\-]+\b)"
    so_number_pattern = r"S\.O\. #\s*:\s*([\d\/]+)"
    po_number_pattern = r"P\.?O\.?\s*Number\s*([A-Za-z]{2}\d{5}|\d{5}|196|PO-\d+)"

    # Extract invoice number
    invoice_number = re.search(invoice_number_pattern, invoice_text)
    if invoice_number:
        invoice_number = invoice_number.group(1)

    # Extract tracking number
    tracking_number = re.search(tracking_number_pattern, invoice_text)
    if tracking_number:
        tracking_number = tracking_number.group(1)

    # Extract serial numbers (s/n)
    # serial_numbers_sn = re.findall(r"s/n\s*([A-Z\d]+)", invoice_text)
    #
    # # Extract serial numbers (Serial Number:)
    # serial_numbers_serial = re.findall(r"Serial Number:\s*([A-Z\d]+)", invoice_text)

    # Use re.findall() to find all matches of the serial number pattern in the input string
    matches = re.findall(serial_number_pattern, invoice_text)

    # Initialize a list to store extracted serial numbers
    serial_numbers = []

    # Iterate over each match tuple
    for match in matches:
        # Iterate over each group in the match tuple
        for group in match:
            # Check if the group is not None (i.e., it contains a valid serial number)
            if group:
                serial_numbers.append(group)

    # Extract S.O. (Sales Order) number
    so_number = re.search(so_number_pattern, invoice_text)
    if so_number:
        so_number = so_number.group(1)



    # print("Invoice #: ", invoice_number)
    # print("Tracking #: ", tracking_number)
    # print("Serial Numbers: ", serial_numbers_serial)
    # print("s/n: ", serial_numbers_sn)
    # print("S.O. #: ", so_number)
    # print("P.O. #: ", po_number)
    return invoice_number,tracking_number,serial_numbers,so_number

def get_po(filepath):
    try:
        so_PO = tabula.read_pdf_with_template(input_path=filepath, template_path="setup/po_so.json",
                                              stream=True)
        po_number = str(so_PO[0]['P.O. Number'].iloc[0]).strip()
        ship_via = str(so_PO[0]['Ship Via:'].iloc[0]).strip()
        return po_number,ship_via
    except Exception:
        print(f"couldnt find po in {Exception}")


def find_and_click_grayscale(filename):
    # Capture screen
    sleep(.1)
    screencap = screen()
    screenshot = cv2.cvtColor(np.array(screencap), cv2.COLOR_RGB2BGR)

    # Convert to grayscale
    gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

    # Example: Template matching for finding a grayscale element
    template = cv2.imread(f"{filename}", 0)
    w, h = template.shape[::-1]
    res = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    threshold = 0.9
    loc = np.where(res >= threshold)

    # Click on the center of the found element
    if loc[0].any():
        center_x = int(loc[1][0] + w // 2)
        center_y = int(loc[0][0] + h // 2)
        click(center_x, center_y)
        sleep(.2)











