import cv2
import pytesseract
from pytesseract import Output
import pandas as pd
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'

def preprocess_image(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
    return thresh

def extract_table_data(image_path):
    preprocessed_img = preprocess_image(image_path)
    ocr_data = pytesseract.image_to_data(preprocessed_img, output_type=Output.DICT)
    
    rows = []
    current_row = []
    prev_top = None

    for i in range(len(ocr_data['text'])):
        text = ocr_data['text'][i].strip()
        if text:
            # Detect new row by comparing the 'top' value
            top = ocr_data['top'][i]
            if prev_top is not None and abs(top - prev_top) > 10:  # Threshold for row change
                rows.append(current_row)
                current_row = []
            current_row.append(text)
            prev_top = top
    
    if current_row:
        rows.append(current_row)
    
    return rows

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False, header=False)

if __name__ == "__main__":
    image_path = "table_image.jpg"  
    output_file = "output.xlsx"
    

    table_data = extract_table_data(image_path)
    print("Extracted Data:", table_data)
    
    save_to_excel(table_data, output_file)
    print(f"Data saved to {output_file}")
