import glob
import os
from pdf2image import convert_from_path
from pytesseract import image_to_string
from PIL import Image

# converts contents in pdf to image
def pdf_to_img(p):
    images = convert_from_path(p)
    return images

# extracts text from images
def img_txt(f):
    text = image_to_string(f)
    return text

# writes text to .txt file
def pdf_to_txt(file):
    images = pdf_to_img(file)
    ft = ""
    for pg,img in enumerate(images):
        f.write(img_txt(img))


path = 'D:\\MSIS\\RA\\'                           # folder where the .pdf files are stored
outpath = "D:\\MSIS\\RA\\"                        # path of output folder
try:
    for i, doc in enumerate(glob.iglob(path + "*.pdf")):
        try:
            filename = doc.split('\\')[-1]
            in_file = os.path.abspath(doc)
            f = open(outpath+ filename[0:-4] + ".txt",'w')
            pdf_to_txt(in_file)
            f.close()
        except:
            print("error "+ filename)

finally:
    pass



