import os
import csv
from docx import Document


header = ['OPS','Date Effective','Currency','Version','Classification']
doclist = []

# path to the folder
path = (r'C:\Users\dan\Desktop\docs')
file_list = []
# appends files in the docs folder to file_list and insert the path + \ in it
for doc_num in os.listdir(path):
    file_list.append(path +'\\' + doc_num)
# uses the Document thing to specifiy its docx i think

def read(file_list=file_list):
    for doc_path in file_list:
        document = Document(doc_path)
        # append each line of text in document to the a new value in doclist
        #doclist = []
        for p in document.paragraphs:
            doclist.append(p.text)
    return doclist

read()


with open('test.csv','w') as csvfile:
    writer = csv.writer(csvfile)

    writer.writerow(header)

    for i in range(2,len(doclist),5):
        new_row = doclist[i:i+5]
        writer.writerows([new_row])

