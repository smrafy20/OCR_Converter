import cv2
import numpy as np
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
from docx import Document
import os
import time

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def process_pdf(pdf_path):
    try:
        # Extract file name without extension
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_word_path = f"{base_name}.docx"
        
        # Create a temporary folder for images
        temp_folder = "temp_images"
        if not os.path.exists(temp_folder):
            os.makedirs(temp_folder)
        
        # Convert PDF to images
        print("Converting PDF to images...")
        images = convert_from_path(pdf_path)
        
        # Create a new Word document
        doc = Document()
        
        # Process each page
        for i, image in enumerate(images):
            image_path = f"{temp_folder}/page_{i+1}.jpg"
            image.save(image_path, "JPEG")
            print(f"Processing page {i+1}...")

            img = cv2.imread(image_path)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            blur = cv2.GaussianBlur(gray, (5,5), 0)
            _, threshold = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_image_path = f"{temp_folder}/processed_{i+1}.jpg"
            cv2.imwrite(processed_image_path, threshold)

            try:
                text = pytesseract.image_to_string(
                    Image.open(processed_image_path), 
                    lang='ben+eng'
                )

                if len(text.strip()) < 10:
                    eng_text = pytesseract.image_to_string(
                        Image.open(processed_image_path), 
                        lang='eng'
                    )
                    ben_text = pytesseract.image_to_string(
                        Image.open(processed_image_path), 
                        lang='ben'
                    )
                    text = eng_text if len(eng_text) > len(ben_text) else ben_text
                    print(f"Used fallback OCR for page {i+1}")
                else:
                    print(f"Used ben+eng OCR for page {i+1}")

                doc.add_paragraph(text)

            except Exception as e:
                print(f"OCR error on page {i+1}: {e}")
                doc.add_paragraph(f"[OCR failed for page {i+1}]")

        try:
            doc.save(output_word_path)
            print(f"Word document saved to {output_word_path}")
        except Exception as save_error:
            print(f"Error saving to default location: {save_error}")
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            alt_output_path = os.path.join(desktop_path, output_word_path)
            try:
                doc.save(alt_output_path)
                print(f"Word document saved to alternative location: {alt_output_path}")
            except Exception as alt_save_error:
                print(f"Error saving to alternative location: {alt_save_error}")

        # Clean up temporary files
        for file in os.listdir(temp_folder):
            os.remove(os.path.join(temp_folder, file))
        os.rmdir(temp_folder)

    except Exception as e:
        print(f"Error: {e}")

# Run the conversion
process_pdf("test.pdf")
