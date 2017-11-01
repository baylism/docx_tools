from docx_class_4 import Document
import os
import csv

files = []

for filename in os.listdir('files'):
    if filename.endswith('.docx') and not filename.startswith('~'):
        files.append('files/' + filename)

outfile = open('counts.csv', 'w', newline='')
outfile_writer = csv.writer(outfile)
outfile_writer.writerow(['File', 'Characters inc spaces and notes', 'Characters inc spaces'])

for item in files:
    doc = Document(item)
    print([item, doc.count_total_chars(), doc.count_chars()])
    outfile_writer.writerow([item, doc.count_total_chars(), doc.count_chars()])

outfile.close()

print(files)



#t = Document('files/cgessay.docx')



#print(t.get_text())
#print(t.count_chars())
#print(t.get_paras())
#print(t.get_text())
