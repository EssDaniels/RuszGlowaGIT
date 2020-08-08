from docxtpl import DocxTemplate
import jinja2
import pandas as pd
from openpyxl import load_workbook
import os
from subprocess import Popen
LIBRE_OFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"


def listaobecnosci():
        workbook = load_workbook(filename="notes_data.xlsx")
        sheet = workbook.active
        i = 0
        a = 4
        b = a + 1
        c = b + 1
        d = str(sheet["H2"].value)
        f = int(d) / 3
        t = 1
        zmiana = 'IV'

        for i in range(0, int(f) + 1):
            doc = DocxTemplate('note6-template.docx')

            context = {
                'gr': zmiana,
                'nr': t,
                'miesiac': str(sheet["J2"].value),
                'pracownik1': str(sheet["C"+str(a)].value),
                'pracownik2': str(sheet["C"+str(b)].value),
                'pracownik3': str(sheet["C"+str(c)].value)
            }
            s = str(t)
            nazwapliku = "lista_obecnosci" + zmiana + s + ".docx"
            doc.render(context)

            doc.save(nazwapliku)

            def convert_to_pdf(input_docx, out_folder):
                p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
                           out_folder, input_docx])
                print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
                p.communicate()

            sample_doc = nazwapliku
            out_folder = 'Zmiana ' + zmiana
            convert_to_pdf(sample_doc, out_folder)
            os.remove(nazwapliku)

            a = c + 1
            b = a + 1
            c = b + 1
            t = t + 1


listaobecnosci()

