#import PyMuPDF
import fitz  # PyMuPDF
import time
import argparse
import openpyxl
import pytesseract
from PIL import Image, ImageFilter, ImageOps
import pandas as pd
import re
import cv2
import numpy as np
from io import BytesIO
from PIL import Image
from subprocess import Popen, PIPE
import requests
from datetime import datetime
from google.cloud import vision
import io

# Path to your PDF file
#pdf_path = 'two_images.pdf'
#output_path = 'spreadsheet.xlsx'
#output_jpeg_path1 = 'image1.jpeg'
#output_jpeg_path2 = 'image2.jpeg'

# Function to extract images from PDF
def extract_images_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    images = []
    for i in range(len(doc)):
        for img in doc.get_page_images(i):
            xref = img[0]
            base_image = Image.open(BytesIO(fitz.Pixmap(doc, xref).tobytes("png")))
            images.append(base_image)
    return images

def preprocess_image_for_ocr(image,output_jpeg_path1,output_jpeg_path2):
    # Open the PDF
    ## crop left upper
     # Update this path
    crop_area = (200, 1800, 1170, 3350)  # Example crop area
    cropped_image = image.crop(crop_area)
    angle = 270  # Rotate 90 degrees counterclockwise
    rotated_image = cropped_image.rotate(angle, expand=True)
    rotated_image.save(output_jpeg_path1, "JPEG")
    image_data1 = open(output_jpeg_path1, "rb").read()


    ## crop right
    crop_area = (350, 650, 750, 1150)  # Example crop area
    cropped_image = image.crop(crop_area)
    angle = 270  # Rotate 90 degrees counterclockwise
    rotated_image = cropped_image.rotate(angle, expand=True)
    rotated_image.save(output_jpeg_path2, "JPEG")
    image_data2 = open(output_jpeg_path2, "rb").read()
    return image_data1, image_data2

def ocr_Azure_output(image_data,subscription_key,read_url):
    headers = {'Ocp-Apim-Subscription-Key': subscription_key, 'Content-Type': 'application/octet-stream'}
    # Send the request
    response = requests.post(read_url, headers=headers, data=image_data)
    response.raise_for_status()

    # Retrieve the URL to get the OCR results
    operation_url = response.headers["Operation-Location"]

    # Wait for the analysis to complete
    analysis = {}
    while not "analyzeResult" in analysis:
        response_final = requests.get(response.headers["Operation-Location"], headers=headers)
        analysis = response_final.json()
        time.sleep(1)
    return analysis["analyzeResult"]["readResults"][0]['lines']



def get_whole(image_data1_output):
    res=''
    for i in range(len(image_data1_output)):
        res=res+image_data1_output[i]['text']+' '
    return res

def fraction_to_float(fraction_str):
    """
    Convert a fraction string to a floating-point number. Returns None if conversion is not possible.
    """
    try:
        if '/' in fraction_str:
            numerator, denominator = fraction_str.split('/')
            return float(numerator) / float(denominator)
        else:
            return float(fraction_str)
    except ValueError:
        return None


def correct_ocr_result(ocr_result, possible_values):
    """
    Correct the OCR result by finding the closest match in the possible values.
    """
    ocr_value = fraction_to_float(ocr_result)
    if ocr_value is not None:
        return ocr_result  # Return the original result if it's a valid number

    # Assuming '3/y' should be corrected to '3/4' as it's the most logical correction based on the context
    if ocr_result == "3/y":
        ocr_value = fraction_to_float("3/4")
    else:
        # Implement other specific corrections if necessary
        pass

    # Convert all possible values to floats
    possible_floats = [fraction_to_float(val) for val in possible_values if fraction_to_float(val) is not None]

    # Find the closest value
    closest_value = min(possible_floats, key=lambda x: abs(x - ocr_value))

    # Convert the closest float back to its string representation
    closest_str = possible_values[possible_floats.index(closest_value)]
    return closest_str

def keep_only_numbers(input_string):
    # Replace all non-digit characters with an empty string
    only_numbers = re.sub(r'\D', '', input_string)
    return only_numbers

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--pdf', dest='pdf', type=str)
    parser.add_argument('--jpeo', dest='jpeo', type=str)
    parser.add_argument('--jpet', dest='jpet', type=str)
    parser.add_argument('--out', dest='out', type=str)
    parser.add_argument('--key', dest='key', type=str)
    parser.add_argument('--end', dest='end', type=str)
    args = parser.parse_args()
    # Path to your PDF file
    # pdf_path = 'two_images.pdf'
    # output_path = 'spreadsheet.xlsx'
    # output_jpeg_path1 = 'image1.jpeg'
    # output_jpeg_path2 = 'image2.jpeg'
    pdf = args.pdf
    jpeo = args.jpeo
    jpet = args.jpet
    out = args.out
    endpoint = args.end
    subscription_key = args.key
    read_url = endpoint + 'vision/v3.1/read/analyze'

    images = extract_images_from_pdf(pdf)
    lis=[]
    # Process each image
    for image in images:
        image_data1, image_data2 = preprocess_image_for_ocr(image,jpeo,jpet)
        image_data1_output = ocr_Azure_output(image_data1,subscription_key,read_url)
        image_data1_text=get_whole(image_data1_output)
        pattern = r"NEAREST\.?\s*CROSS STREET\s+(.*?)\s+\.?\s*SERVICE"
        nearest_cross_street = re.search(pattern, image_data1_text, re.IGNORECASE).group(1)
        #image_data2_output = ocr_Azure_output(image_data2, subscription_key, read_url)
        #image_data2_text = get_whole(image_data2_output)
        renew_size = re.search(r".*RENEW(.*?)SIZE", image_data1_text, re.IGNORECASE).group(1).replace(" ", "")
        possible_values = ['3/4', '1/2', '1']
        if renew_size:
            renew_size = correct_ocr_result(renew_size, possible_values)
        #renew_size = image_data1_output[15]['text'].replace(" ", "")
        renew_date = datetime.strptime(re.search(r".*DATE(.*?)FOREMAN", image_data1_text, re.IGNORECASE).group(1).replace(" ", ""), '%m-%d-%y').strftime('%m-%d-%Y')
        image_data2_output = ocr_Azure_output(image_data2,subscription_key,read_url)[1]['text']
        No_house= re.search(r'NO.\s+(.+)', image_data2_output).group(1)
        No_house=keep_only_numbers(No_house)
        extracted_data = [nearest_cross_street,renew_size,renew_date,No_house]
        #df.append(extracted_data, ignore_index=True)
        lis.append(extracted_data)

    df = pd.DataFrame(lis, columns=["Nearest_Cross_Street", "Renew_Size", "Renew_Date", "House_No"])
    # Save the dataframe to an Excel spreadsheet
    df.to_excel(out, index=False)
    print(f"Extraction complete. Data saved to {out}")


