from PyPDF2 import PdfFileWriter, PdfFileReader

pdf_dir = r"C:\Users\1907043\Downloads\pdf"
file_list = os.listdir(pdf_dir)
location = os.path.join(pdf_dir, file_list[0])
new_location = os.path.join(pdf_dir, 'new.pdf')

pages_to_keep = [0,6,7,8,9,10,14,15,16,17,18,19,20,26,28,40,49,50,51,52,53,54,58,59,61,62,68,69,70,135,136,181]
infile = PdfFileReader(location, 'rb')
output = PdfFileWriter()

for i in pages_to_keep:
    p = infile.getPage(i)
    output.addPage(p)
    
with open(new_location, 'wb') as f:
    output.write(f)
