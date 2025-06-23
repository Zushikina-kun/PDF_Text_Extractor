import os
import threading
import tempfile
import gc
from tkinter import Tk, filedialog, Label, Button, StringVar, OptionMenu, ttk
from PIL import Image
import fitz  # PyMuPDF
import easyocr
from docx import Document

# Initialize EasyOCR reader
reader = easyocr.Reader(['en'], gpu=False)

def extract_text_from_image(image_path):
    """
    Extract text from an image file using EasyOCR.
    """
    return reader.readtext(image_path, detail=0, paragraph=True)

def extract_text_from_pdf(pdf_path, update_progress):
    """
    Extract text from a PDF file using PyMuPDF and EasyOCR.
    """
    text = ""
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=300)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
                tmp_img.write(pix.tobytes())
                tmp_img_path = tmp_img.name

            page_text = reader.readtext(tmp_img_path, detail=0, paragraph=True)
            os.remove(tmp_img_path)
            text += f"\n--- Page {i + 1} ---\n" + "\n".join(page_text)
            update_progress(i + 1, len(doc))
            gc.collect()
    return text

def save_text_to_file(text, output_path, file_type="txt"):
    """
    Save extracted text to a file in TXT or DOCX format.
    """
    if file_type == "txt":
        with open(output_path, "w", encoding="utf-8") as file:
            file.write(text)
    elif file_type == "docx":
        doc = Document()
        for para in text.split('\n'):
            doc.add_paragraph(para)
        doc.save(output_path)
    else:
        raise ValueError("Unsupported file type. Use 'txt' or 'docx'.")

def browse_file():
    file_path.set(filedialog.askopenfilename(filetypes=[("PDF and Image Files", "*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff")]))

def save_file():
    output_path.set(filedialog.asksaveasfilename(defaultextension=f".{output_format.get()}",
                                                 filetypes=[("Text Files", "*.txt"), ("Word Documents", "*.docx")]))

def update_progress(current, total):
    progress_var.set(int((current / total) * 100))
    progress_bar.update()

def process_file():
    """
    Process the selected file for text extraction and save the output.
    """
    if not file_path.get():
        status_label.config(text="Please select a file.")
        return
    if not output_path.get():
        status_label.config(text="Please specify an output file.")
        return

    def task():
        try:
            if file_path.get().lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                text = "\n".join(extract_text_from_image(file_path.get()))
                update_progress(1, 1)
            elif file_path.get().lower().endswith('.pdf'):
                text = extract_text_from_pdf(file_path.get(), update_progress)
            else:
                status_label.config(text="Unsupported file format.")
                return

            save_text_to_file(text, output_path.get(), file_type=output_format.get())
            status_label.config(text=f"Text extracted and saved to {output_path.get()}")
        except Exception as e:
            status_label.config(text=f"Error: {str(e)}")

    threading.Thread(target=task).start()

def main():
    global file_path, output_path, output_format, status_label, progress_var, progress_bar

    root = Tk()
    root.title("PDF and Image Text Extractor")
    root.geometry("700x350")

    file_path = StringVar()
    output_path = StringVar()
    output_format = StringVar(value="txt")
    progress_var = StringVar(value=0)

    Label(root, text="Select File:").grid(row=0, column=0, padx=10, pady=10)
    Button(root, text="Browse", command=browse_file).grid(row=0, column=1, padx=10, pady=10)
    Label(root, textvariable=file_path, wraplength=400).grid(row=0, column=2, padx=10, pady=10)

    Label(root, text="Output Format:").grid(row=1, column=0, padx=10, pady=10)
    OptionMenu(root, output_format, "txt", "docx").grid(row=1, column=1, padx=10, pady=10)

    Label(root, text="Save As:").grid(row=2, column=0, padx=10, pady=10)
    Button(root, text="Save", command=save_file).grid(row=2, column=1, padx=10, pady=10)
    Label(root, textvariable=output_path, wraplength=400).grid(row=2, column=2, padx=10, pady=10)

    Button(root, text="Process", command=process_file).grid(row=3, column=0, columnspan=3, pady=20)

    progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate", variable=progress_var, maximum=100)
    progress_bar.grid(row=4, column=0, columnspan=3, pady=10)

    status_label = Label(root, text="", fg="blue")
    status_label.grid(row=5, column=0, columnspan=3, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
