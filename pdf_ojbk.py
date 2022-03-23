import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2 import PdfFileMerger

def getFileName(filedir):

    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []

def handle(pdf_filename):
    merger = PdfFileMerger()
    signature = 33
    for f in pdf_filename:
        if 'out' in f:
            output = f
        if '初始' in f:
            input1 = open(f, "rb")
            #print(input1)
        if '附件一' in f:
            input2 = open(f, "rb")
        if '附件二' in f:
            input3 = open(f, "rb")
        if '附件三' in f:
            input4 = open(f, "rb")
        if '附件四' in f:
            input5 = open(f, "rb")
        if '附件五' in f:
            input6 = open(f, "rb")

    #merger.append(fileobj = input1)
    merger.merge(position=signature, fileobj=input2)
    #merger.merge(position=signature+1, fileobj=input3)
    #merger.merge(position=signature+2, fileobj=input4)
    merger.merge(position=signature+3, fileobj=input5)
    merger.append(input6)
    output = open(output, "wb")
    merger.write(output)

if __name__ == "__main__":
    pdf_filename =  getFileName('K:\mashuaifei\A2019044-T002-01 SD大鼠灌胃给予TL139单次给药毒性试验总结报告\新建文件夹')
    print(pdf_filename)
    handle(pdf_filename)