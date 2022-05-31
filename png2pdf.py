import glob
from fpdf import FPDF
import os

image_dir = r'C:\Users\1907043\Downloads\result\list'
file_list = os.listdir(image_dir)

pdf = FPDF()
for image in file_list:
    pdf.add_page()
    pdf.image(os.path.join(image_dir, image), 0, 0, 210, 297)
pdf.output(os.path.join(image_dir, '국정과제_관련부서 발췌본.pdf'), "F")





