import docx
import glob
import os


path = ""                # provide the path to the folder containing text documnets
outpath = ""             # path of output folder

def text_to_docx(text_file, docx_file):
    # Open the text file
    with open(text_file, "r") as f:
        text = f.read()
        # print(text)

    # Create a new docx file
    doc = docx.Document()

    # Add the text to the document
    doc.add_paragraph(text)

    # Save the document
    doc.save(docx_file)


for i, doc in enumerate(glob.iglob(path + "*.txt")):
    filename = doc.split('\\')[-1][:-4]
    in_file = os.path.abspath(doc)
    out_file = outpath+filename+".docx"    
    print(out_file)
    text_to_docx(in_file,out_file)
