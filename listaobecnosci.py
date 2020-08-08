from docxtpl import DocxTemplate
import jinja2
import pandas as pd


doc = DocxTemplate('lista_obecnosci.docx')
note_data = pd.read_excel('lo.xlsx')
context = dict(zip(note_data['var'], note_data['nazwisko']))
doc.render(context)
doc.save('listaobecnosci1.docx')


