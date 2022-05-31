from pdf2image import convert_from_path
from pdf2image import pdfinfo_from_path
import glob
import os

def convert_pdf(pdf_dir, filename):
    location = os.path.join(pdf_dir, filename)
    pages = convert_from_path(location, 600, poppler_path = r'C:\Program Files (x86)\poppler-0.68.0\bin')
    name, ext = os.path.splitext(filename)
    for i, page in enumerate(pages):
        if len(pages) == 1:
            page.save(r'C:/Users/1907043/Downloads/result/png/' + name + '.png', 'PNG')
            page.save(r'C:/Users/1907043/Downloads/result/jpg/' + name + '.jpg', 'JPEG')
        else:
            page.save(r'C:/Users/1907043/Downloads/result/png/' + name + '(' + f"{i+1:03}" + ')' + '.png', 'PNG')
            page.save(r'C:/Users/1907043/Downloads/result/jpg/' + name + '(' + f"{i+1:03}" + ')' + '.jpg', 'JPEG')
        # page.save(f'{img_dir+pdf_[len(path):-4]}_page{i+1:0>2d}.jpg', 'JPEG')
        # print(f'{pdf_[len(path):-4]}_page{i+1:0>2d}.jpg saved...')
    print(filename, ': ', i+1, "/", len(pages))

pdf_dir = r"C:\Users\1907043\Downloads\pdf"
file_list = os.listdir(pdf_dir)


location = os.path.join(pdf_dir, file_list[0])
info = pdfinfo_from_path(location, poppler_path = r'C:\Program Files (x86)\poppler-0.68.0\bin')

name, ext = os.path.splitext(file_list[0])

maxPages = info['Pages']
for i in range(1, maxPages+1, 10):
    res = convert_from_path(location, dpi=600, first_page=i, last_page=min(i+10-1, maxPages), poppler_path = r'C:\Program Files (x86)\poppler-0.68.0\bin')
    for idx, r in enumerate(res):
        r.save(r'C:/Users/1907043/Downloads/result/png/' + name + '(' + f"{i+idx:03}" + ')' + '.png', 'PNG')

for i in file_list:
    convert_pdf(pdf_dir, i)
    print("Done!")




