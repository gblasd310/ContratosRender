from msilib import Directory
from docxtpl import DocxTemplate
import pandas as pd
import os

directoryOutputs = 'C:/Users/Gustavo Blas/OneDrive - Financera Sustentable de México SA de CV SFP/REFERENCIAS_PV_OCTUBRE'
directoryCSV = 'C:/Files_Manager_Finsus/src/referencias_PV_octubre_individual.csv'

data = pd.read_csv(directoryCSV, encoding='utf-8')


for dt in data.index:
    sheet_reference = DocxTemplate('C:/Files_Manager_Finsus/layouts/Hoja de Referencia.docx')

    context = {
        'referencia_bancaria'   :   str(data['referencia'][dt]).zfill(10),
        'vin'                   :   str(data['vin'][dt]),
        'credito'               :   str(data['credito'][dt])
    }

    try:
        os.stat(directoryOutputs)
    except:
        os.mkdir(directoryOutputs)

    sheet_reference.render(context)
    sheet_reference.save(directoryOutputs + '/' + str(data['credito'][dt]) + '_' + str(data['vin'][dt]) + '_' + str(data['referencia'][dt]).zfill(10) + ".docx")

print('se han generado todas las referencias correctamente...')

