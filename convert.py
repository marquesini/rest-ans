import os
import win32com.client as win32
import sys
import zipfile
import xml
from xml.etree import ElementTree
from openpyxl import load_workbook
import pandas as pd

from urllib import request

from flask import Flask, send_file

app = Flask(__name__)

REMOTE_URL = 'https://github.com/raizen-analytics/data-engineering-test/raw/master/assets/vendas-combustiveis-m3.xls'

def downloadFile():
    try:
        request.urlretrieve(REMOTE_URL, 'files/in/dados.xls')
    except Exception as error:
        print('Ocorreu um erro ao fazer o download do arquivo.')
        print(error)

def xls2Xlsx():
    try:
        file_name = os.getcwd() + '/files/in/dados.xls'

        if os.path.exists(file_name):
            os.remove(file_name)

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file_name)     
        wb.SaveAs(file_name + 'x', FileFormat = 51)
        wb.Close()                               
        excel.Application.Quit()
        print('Conversão realizada com sucesso.')
    except Exception as error:
        print('Ocorreu um erro ao converter o arquivo.')
        print(error)

def unzipXlsx():
    with zipfile.ZipFile(os.getcwd() + '/files/in/dados.xlsx', 'r') as zip_ref:
        zip_ref.extractall('files/out/xlsx-unziped')

def getPivotCache(source):

    definition = 'pivotCacheDefinition1' if source == 'oil' else 'pivotCacheDefinition2'

    definitions = f'files/out/xlsx-unziped/xl/pivotCache/{definition}.xml'

    defdict = {}
    columnas = []
    e = xml.etree.ElementTree.parse(definitions).getroot()
    for fields in e.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cacheFields'):
        for cidx, field in enumerate(fields.getchildren()):
            columna = field.attrib.get('name')
            defdict[cidx] = []
            columnas.append(columna)
            for value in field.getchildren()[0].getchildren():
                tagname = value.tag
                defdict[cidx].append(value.attrib.get('v', 0))

    dfdata = []

    records = 'pivotCacheRecords1' if source == 'oil' else 'pivotCacheRecords2'

    bdata = f'files/out/xlsx-unziped/xl/pivotCache/{records}.xml'

    for event, elem in xml.etree.ElementTree.iterparse(bdata, events=('start', 'end')):
        if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r' and event == 'start':
            tmpdata = []
            for cidx, valueobj in enumerate(elem.getchildren()):
                tagname = valueobj.tag
                vattrib = valueobj.attrib.get('v')
                rdata = vattrib
                if tagname == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}x':
                    try:
                        rdata = defdict[cidx][int(vattrib)]
                    except:
                        logging.error(
                            'this it not should happen index cidx = {} vattrib = {} defaultidcts = {} tmpdata for the time = {} xml raw {}'
                            .format(cidx, vattrib, defdict, tmpdata, xml.etree.ElementTree.tostring(
                                elem, encoding='utf8', method='xml')
                            )
                        )
                tmpdata.append(rdata)
                
            if tmpdata:
                dfdata.append(tmpdata)

            elem.clear()
    
    df = pd.DataFrame(dfdata)
    df.to_csv(f'files/out/{source}.csv', index=False)

def prepareDownload():
    downloadFile()
    xls2Xlsx()
    unzipXlsx()

@app.route('/oil')
def get_oil():
    prepareDownload()
    getPivotCache('oil')
    return send_file('files/out/oil.csv', as_attachment=True)

@app.route('/diesel')
def get_diesel():
    prepareDownload()
    getPivotCache('diesel')
    return send_file('files/out/diesel.csv', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')