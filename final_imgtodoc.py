import glob
import os
from pytesseract import image_to_string
import docx

# extracts text from image
def img_txt(f):
    text = image_to_string(f)
    return text

def text_to_docx(text, docx_file):
    doc = docx.Document()

    # Add the text to the document
    doc.add_paragraph(text)

    # Save the document
    doc.save(docx_file)


path = 'D:\\MSIS\\RA\\jpg files\\'          # path of folder containing images
out = 'D:\\MSIS\\RA\\jpg files\\docx\\'     # path of folder to store final docx files


try:
    for i, doc in enumerate(glob.iglob(path + "*.jpg")):
        filename = doc.split('\\')[-1][:-5]
        in_file = os.path.abspath(doc)
        print(in_file)
        outfile = out+filename+'.docx'
        txt = img_txt(in_file)
        text_to_docx(txt, outfile)
except:
    print("Error -"+filename)