# -*- coding: utf-8 -*-
"""
Created on Mon May  5 16:43:47 2025

@author: gustavo.andrade
"""

from PIL import Image
import pytesseract

# Optional: Set path to tesseract executable (only if not in PATH)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_image(image_path):
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        return f"Error: {e}"

# Example usage
if __name__ == "__main__":
    image_path = r"C:\Users\gustavo.andrade\Pictures\Amilton.jpg"  # Replace with your image file path
    text = extract_text_from_image(image_path)
    print("Extracted Text:\n", text)
