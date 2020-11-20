from docx import *
import re
from copy import deepcopy
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT as WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as WD_ALIGN_PARAGRAPH
from XMLReader import XMLReader
from docx.shared import Pt
from docx.enum.style import WD_BUILTIN_STYLE
import sys


def tableNilaiCPL(doc):
    table = doc.tables[4]
    for row in range(len(table.rows)):
        if 'Peta CPL – CP MK' in table.cell(row,0).text:
            for col in range(len(table.columns)):
                str_peta_cpl = table.cell(row,col)
                if str_peta_cpl.tables:
                    return str_peta_cpl.tables[0]


def ambilBobotCPl(tabel):
    cpl_dict = {}
    for column in range(1,len(tabel.columns)):
        for row in range(1,len(tabel.rows)):
            if tabel.cell(row,column).text != '':
                cpl_key = tabel.cell(0,column).text
                cpl_value = tabel.cell(row,column).text
                cpl_dict[re.sub('CPL\s','',cpl_key)] = float(cpl_value)
    return cpl_dict

def CpmkCpl(tabel):
    cpmk_cpl = {}
    for row in range(1,len(tabel.rows)):
        cpl_isi=[]
        for column in range(1,len(tabel.columns)):
            if tabel.cell(row,column).text != '':
                cpmk_cpl_key = tabel.cell(row,0).text
                cpl = tabel.cell(0,column).text
                cpl_isi.append(re.sub('CPL\s','',cpl))
        cpmk_cpl[cpmk_cpl_key]=cpl_isi
    return cpmk_cpl

def ambilMinggu(table):
    minggu = []
    for row in range(1,len(table.rows)-1):
        minggu.append(table.cell(row,0).text)
    return minggu

def ambilCpmk(table):
    cpmk = []
    for row in range(1,len(table.rows)-1):
        cpmk.append(table.cell(row,1).text)
    return cpmk

def ambilBentukPenilaian(table):
    bentuk_penilaian = []
    for row in range(1,len(table.rows)-1):
        bentuk_penilaian.append(table.cell(row,2).text.splitlines()[0])
    return bentuk_penilaian

def ambilBobotCpmk(table):
    bobot = []
    for row in range(1,len(table.rows)-1):
        bobot.append(table.cell(row,3).text)
    return bobot

def ambilFailDesc(table):
    fail_desc = []
    for row in range(1,len(table.rows)-1):
        if len(table.cell(row,2).text.splitlines()) > 1:
            fail_desc.append(table.cell(row,2).text.splitlines()[-1])
        else:
            cpmk_ke = re.findall('CPMK\s\d',table.cell(row,1).text)[0]
            fail_desc.append(f'Tidak Lulus {cpmk_ke}')
    return fail_desc

def writeNormalStyle(cell,fill):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER # pylint: disable=no-member
    cell_bobot = cell.paragraphs[0]
    cell_bobot.alignment = WD_ALIGN_PARAGRAPH.CENTER # pylint: disable=no-member
    cell_bobot.text = fill

def writeBoldStyle(cell,fill):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER # pylint: disable=no-member
    cell_paragraph = cell.paragraphs[0]
    cell_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER # pylint: disable=no-member
    cell_paragraph.add_run(fill).bold = True

def writeCpmk(cell, fill):
    cell.text = fill.split('\n')[0]
    for line in fill.split('\n'):
        if 'CPMK' in line:
            pass
        else:
            cell.add_paragraph(line)

def writeCpl(cell, cpmk_cpl_dict, cpmk):
    cpl_cell = cell.paragraphs[0]
    cpl_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER # pylint: disable=no-member
    for key,value in cpmk_cpl_dict.items():
        if key in cpmk:
            cpl_cell.text = '\n'.join(value)
            break

def writeNiliaXBobot(cell,nilai,bobot):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER # pylint: disable=no-member
    cell_nilai_bobot = cell.paragraphs[0]
    cell_nilai_bobot.alignment = WD_ALIGN_PARAGRAPH.CENTER # pylint: disable=no-member
    nilaixbobot = float(nilai)*(float(bobot)/100)
    cell_nilai_bobot.text = str(round(nilaixbobot,2))

def writeKetercapaianCpl(cell,cpl_cell,cell_nilaixbobot,cpl_nilai_dict):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER # pylint: disable=no-member
    cell_ketercapaiancpl = cell.paragraphs[0]
    cell_ketercapaiancpl.alignment = WD_ALIGN_PARAGRAPH.CENTER # pylint: disable=no-member
    nilaicpl = 0.0
    cpl = cpl_cell.text.split('\n')
    for key,value in cpl_nilai_dict.items():
        if key in cpl:
            nilaicpl += (value*float(cell_nilaixbobot.text))
    cell_ketercapaiancpl.text = str(round(nilaicpl,1))

def writeDesc(cell,nilai,cpmk,fail_desc):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER # pylint: disable=no-member
    cell_desc = cell.paragraphs[0]
    cell_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER # pylint: disable=no-member
    if float(nilai) > 50:
        cell_desc.text = 'Lulus {}'.format(cpmk.split("\n")[0])
    else:
        cell_desc.text = fail_desc

