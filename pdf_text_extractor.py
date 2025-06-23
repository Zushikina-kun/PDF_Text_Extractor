import os
import fitz  # PyMuPDF
import easyocr
from PIL import Image
import numpy as np
from io import BytesIO
import gc

# Initialize the OCR reader globally to avoid reloading multiple times
reader = easyocr.Reader(["en"], gpu=False)

def update_progress(current, total):
    """Simple progress update function."""
    print(f"Progress: {current}/{total} pages processed.")

def extract_text_from_pdf(pdf_path, update_progress):
    """Extract text from a PDF file using EasyOCR."""
    text = ""
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            # Render the page to an image
            pix = page.get_pixmap(dpi=150)  # Render page at 150 DPI
            img = Image.open(BytesIO(pix.tobytes(output="png")))  # Convert pixmap to PIL Image

            # Convert PIL Image to NumPy array for easyocr
            img_array = np.array(img)

            # Perform OCR using easyocr
            page_text = reader.readtext(img_array, detail=0, paragraph=True)
            text += f"\n--- Page {i + 1} ---\n" + "\n".join(page_text)

            # Update progress
            update_progress(i + 1, len(doc))

            # Clean up memory
            img.close()
            gc.collect()

    return text

def save_text_to_file(text, output_file):
    """Save extracted text to a file."""
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(text)

def main():
    # Input PDF file
    pdf_path = input("Enter the path to the PDF file: ").strip()
    if not os.path.exists(pdf_path):
        print("Error: The specified file does not exist.")
        return

    # Output text file
    output_file = input("Enter the output text file name: ").strip()
    if not output_file.endswith(".txt"):
        output_file += ".txt"

    print("Starting OCR...")
    try:
        extracted_text = extract_text_from_pdf(pdf_path, update_progress)
        save_text_to_file(extracted_text, output_file)
        print(f"OCR completed. Extracted text saved to: {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
