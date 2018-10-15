from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx import *

document = Document()
styles = document.styles
for s in styles:
    if s.type == WD_STYLE_TYPE.TABLE:
            print(s.name)