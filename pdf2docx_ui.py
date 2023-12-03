import gradio as gr
from pdf2docx import Converter
import docx2txt
from docx import Document, ImagePart
import os


def convert_pdf_to_docx_with_display(pdf_file):
    tmp_dir = "./docx"
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_file = os.path.join(tmp_dir, "output.docx")
    # Convert PDF to DOCX
    cv = Converter(pdf_file)
    cv.convert(tmp_file)
    cv.close()

    # Extract text from DOCX
    docx_text = docx2txt.process(tmp_file)
    document = Document(tmp_file)
    # Extract images from DOCX
    images = []
    image_dir = os.path.join(tmp_dir, "images")
    os.makedirs(image_dir, exist_ok=True)
    for embed, related_part in document.part.related_parts.items():
        if isinstance(related_part, ImagePart):
            image_path = os.path.join(image_dir, f'image_{embed}.png')
            with open(image_path, 'wb') as f:
                f.write(related_part.image.blob)
                images.append(image_path)

    return tmp_file, docx_text, images


def convert_and_display_pdf_to_docx(text_pdf, pdf_file):
    print(text_pdf)
    docx_file, docx_text, images = convert_pdf_to_docx_with_display(pdf_file)
    return docx_file, docx_text, images


iface = gr.Interface(
    fn=convert_and_display_pdf_to_docx,
    inputs=["text","file"],
    outputs=["file", "text", "gallery"],
    title="[努力的小雨] PDF to DOCX Converter",
    description="上传pdf文件，并将其转化为docx文件且在界面单独显示文件的文字和图像~",
)

iface.launch()
