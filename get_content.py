from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches
import datetime
from docxtpl import DocxTemplate,RichText

def generate_doc_filename():
    doc_filename = str(datetime.datetime.now()) + '.docx'
    return doc_filename

wb = load_workbook('VPNTC.xlsx')
document = Document()
first_line = 5

# const
nm = wb.sheetnames[2]
print(nm)

def write_doc_table(document, contents):
    table = document.add_table(rows=8, cols=2)
    headcol_cells = table.columns[0].cells
    table.columns[0].width = Inches(0.2)
    headcol_cells[0].text = '模块'
    headcol_cells[1].text = '功能'
    headcol_cells[2].text = '测试点'
    headcol_cells[3].text = '测试拓扑'
    headcol_cells[4].text = '测试步骤'
    headcol_cells[5].text = '预期结果'
    headcol_cells[6].text = '测试结果'
    headcol_cells[7].text = '备注'
    print('**************************')
    sndcol_cells = table.columns[1].cells
    sndcol_cells[0].text = str(contents[1])
    print(sndcol_cells[0].text)
    sndcol_cells[1].text = str(contents[2])
    sndcol_cells[2].text = str(contents[3])
    sndcol_cells[3].text = ''
    sndcol_cells[4].text = str(contents[5])
    sndcol_cells[5].text = str(contents[6])
    sndcol_cells[6].text = str(contents[7])
    sndcol_cells[7].text = str(contents[8])

ws = wb[nm]
current_module = None
current_function = None
records = []

for i in range (first_line, ws.max_row):
    print("#####################################")
    cid = ws[i][0].value
    module = ws[i][1].value
    function = ws[i][2].value
    case = ws[i][3].value
    level = ws[i][4].value
    steps = ws[i][5].value
    presult = ws[i][6].value
    result = ws[i][7].value
    beizhu = ws[i][8].value

    if module == None:
        module = current_module
    else:
        current_module = module
    
    if function == None:
        function = current_function
    else:
        current_function = function
    
    records.append({'module':str(module), 'function':str(function), 'case':str(case), 'top':'',
    'steps':RichText(str(steps)), 'presult':RichText(str(presult)), 'result':str(result), 'beizhu':str(beizhu)})
    print(records)

    #write_doc_table(document, record)

#document.save('test1.docx')

doc = DocxTemplate("templates/template2.docx")
context = {
    'contents':records,
}
doc.render(context)
doc.save("output/generated_doc5.docx")