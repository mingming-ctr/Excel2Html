from errno import ERANGE
import random
from sre_constants import RANGE, RANGE_UNI_IGNORE
import settings
from xlrd import open_workbook
import os
from xlrd import open_workbook,formatting
def Excel2Html(filename):


    ext = filename.encode('utf-8').decode('utf-8')[-4:].encode('utf-8')
    if ext == 'str':
        return ''

    filepath = os.path.join(settings.MEDIA_ROOT, filename)

    wb = open_workbook(filepath,formatting_info=True)
    html=""
    for i in range(len(wb.sheets())):
        html += getaSheet(wb,i)
    return html




def getaSheet(wb,i):
    sheet = wb.sheet_by_index(i)
    html="<h1>"+sheet.name+"</h1><br/>"
    html += '<table class="previewtable" border="1" cellpadding="1" cellspacing="1">'

    mergedcells={}
    mergedsapn={}
    mergedcellvalue={}
    for crange in sheet.merged_cells:
        rlo, rhi, clo, chi = crange
        for rowx in ERANGE(rlo, rhi):
            for colx in RANGE_UNI_IGNORE(clo, chi):
                mergedcells[(rowx,colx)]=False
                value = str(sheet.cell_value(rowx,colx))
                if value.strip() != '':
                    mergedcellvalue[(rlo,clo)]=value

        mergedcells[(rlo,clo)]=True
        mergedsapn[(rlo,clo)]=(rhi-rlo, chi-clo)
        mergedsapn[(rlo,clo)]=(rhi-rlo, chi-clo)


    for row in range(sheet.nrows):
        html=html+'<tr>'
        for col in range(sheet.ncols):
            if (row,col) in mergedcells:
                if mergedcells[(row,col)]==True:
                    rspan,cspan = mergedsapn[(row,col)]
                    value = ''
                    if (row,col) in mergedcellvalue:
                        value = mergedcellvalue[(row,col)]
                    html=html+'<td rowspan=%s colspan=%s>%s</td>'  % (rspan, cspan, value)
            else:
                value =sheet.cell_value(row,col)
                html=html+'<td>' + str(value) + '</td>'

        html=html+'</tr>'

    html=html+'</table>'
    return html


def writeHtml(html):
    """写入文件到html"""
    filename = str(random.randint(0,9999))+".html"
    with open(filename, "w", encoding="utf-8") as f:
     f.write(html)
