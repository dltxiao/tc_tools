from docxtpl import DocxTemplate

doc = DocxTemplate("templates/template2.docx")
list1 = [
    {'module':'1','function':'2','case':'3','top':'4','steps':'5','presult':'6','result':'7','beizhu':'8'},
    {'module':'11','function':'12','case':'13','top':'14','steps':'15','presult':'16','result':'17','beizhu':'18'},
    ]
context = {
    'contents':list1,
}
doc.render(context)
doc.save("output/generated_doc4.docx")