from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd

doc = DocxTemplate("notificacion.docx")

#carga de datos dinamicos, desde el excel

df = pd.read_excel("calculocnr.xlsx", sheet_name="informacion")

for index, fila in df.iterrows():
    context = {'NOMBRE_USUARIO': fila['NOMBRE'],
               'DIRECCION': fila['DIRECCION'],
               'IMPORTE': fila['IMPORT'],
               'CNR': fila['CNR'],
               'NUMERACION': fila['NUMERACION'],
               'NIS': fila['CODSUM']
    }

    context.update(my_context)

    doc.render(context)
    doc.save(f"doc_generado_{index}.docx")