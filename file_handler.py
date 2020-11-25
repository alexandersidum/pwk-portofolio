from flask import Response
from generate_table import *
import os
import tempfile

class FileHandler():
    

    def is_xml_file_allowed(self, filename):
        return '.' in filename and filename.rsplit('.',1)[1] == 'xml'

    def is_docx_file_allowed(self, filename):
        return '.' in filename and filename.rsplit('.',1)[1] == 'docx'

    def generate_output_table(self, docx_file, xml_file, t_fd=None, t_fname=None ):
        try:
            doc = Document(docx_file)
            xml_data = XMLReader(xml_file).getData()
            table_cpl = tableNilaiCPL(doc)
            bobot_cpl = ambilBobotCPl(table_cpl)
            cpmk_cpl = CpmkCpl(table_cpl)
            rae_table = doc.tables[-2]
            cpmk = ambilCpmk(rae_table)
            mg_ke = ambilMinggu(rae_table)
            bentuk_penilaian = ambilBentukPenilaian(rae_table)
            bobot_cpmk = ambilBobotCpmk(rae_table)
            fail_desc = ambilFailDesc(rae_table)


            for x in xml_data:
                data_nilai = deepcopy(x[1:])
                template_doc = Document('template_tbl.docx')
                template_section = template_doc.sections[0]
                left, right, top, bottom, gutt = template_section.left_margin, template_section.right_margin, template_section.top_margin, template_section.bottom_margin, template_section.gutter
                new_width, new_height = template_section.page_width, template_section.page_height
                section = doc.add_section()
                section.left_margin, section.right_margin, section.top_margin, section.bottom_margin, section.gutter = left, right, top, bottom, gutt
                section.page_width , section.page_height = new_width, new_height
                keterangan = doc.add_paragraph().add_run('Portofolio penilaian & evaluasi proses dan hasil belajar {}'.format(data_nilai[0]))
                font = keterangan.font
                font.name = 'Britannic Bold'
                font.size = Pt(12)
                table = template_doc.tables[-1]
                new_tbl = deepcopy(table._tbl)
                paragraph =  doc.add_paragraph()
                paragraph._p.addnext(new_tbl)

                tbl = doc.tables[-1]
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER # pylint: disable=no-member
                tbl.style = 'Table Grid'

                for i in range(len(mg_ke)):
                    cell = tbl.add_row().cells
                    writeBoldStyle(cell[0],mg_ke[i])
                    writeCpl(cell[1],cpmk_cpl,cpmk[i])
                    writeCpmk(cell[2],cpmk[i])
                    writeNormalStyle(cell[3],bentuk_penilaian[i])
                    writeNormalStyle(cell[4],bobot_cpmk[i])
                    writeNormalStyle(cell[5],data_nilai[i+1])
                    writeNiliaXBobot(cell[6],data_nilai[i+1],bobot_cpmk[i])
                    writeKetercapaianCpl(cell[7],cell[1],cell[6],bobot_cpl)
                    writeDesc(cell[-1],data_nilai[i+1],cpmk[i],fail_desc[i])
            if t_fname is not None and t_fd is not None :
                os.close(t_fd)
                doc.save(t_fname)
                return True
            else :
                return True
        except :
            return False
