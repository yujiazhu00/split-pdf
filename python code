import openpyxl
from PyPDF2 import PdfFileReader, PdfFileWriter

def read_excel(filename,ini_row,ini_col):
    wb = openpyxl.load_workbook(filename)
    sheet = wb['Sheet']
    max_row = sheet.max_row
    output_list = []
    for i in range(0,max_row-1):
        output_list = output_list + [sheet.cell(row=ini_row+i, column=ini_col).value]
    return output_list


def page_number_start(input_list):
    new_list = [0]
    length = len(input_list)
    for i in range(0,length):
        value = new_list[i]+input_list[i]
        new_list = new_list +[value]
    return new_list


def extract_pages(master_file,name_list,page_list):
    length = len(name_list)
    for i in range(0,length):
        master_file_path = master_file
        pdf_file_path = name_list[i]
        file_base_name = pdf_file_path.replace('.pdf', '')
        pdf = PdfFileReader(master_file_path)
        pages = list(range(page_list[i], page_list[i+1]))
        pdfWriter = PdfFileWriter()
        for page_num in pages:
            pdfWriter.addPage(pdf.getPage(page_num))
        with open('{0}_trans.pdf'.format(file_base_name), 'wb') as f:
            pdfWriter.write(f)
            f.close()

def split_pdf(pdf_file,excel_file,name_row,name_col,page_row,page_col):
    namelist = read_excel(excel_file,name_row,name_col)
    pagelist = read_excel(excel_file,page_row,page_col)
    pagelist_final = page_number_start(pagelist)
    extract_pages(pdf_file,namelist,pagelist_final)

