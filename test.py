import datetime

filename = str(datetime.datetime.now()) + '.docx'
print(filename)
print(type(filename))



a = None
print(type(a))
b = str(a)
print(type(b))
print(b)

from docx import Document
document = Document()

table = document.add_table(rows=8, cols=2)
headcol_cells = table.columns[0]
headcol_cells.width = 10
sndcol_cells = table.columns[1]
sndcol_cells.width = 20
document.save('test2.docx')