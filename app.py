from flask import render_template, Flask, request, flash, url_for, redirect, send_from_directory, after_this_request
from generate_table import *
import os

#Pakai quart?
#Relative path
#Apa bisa dilakukan didalam client
#Supaya nggak bentrok gimana
#Coba deploy ke free hosting
#request file perlu disave?
#Perbagus UI


app = Flask(__name__)
app.secret_key = "super secret key"
app.config['UPLOAD_FOLDER']    =   'output_files'
ALLOWED_EXTENSIONS = {'xml', 'docx'}


def is_file_allowed(filename):
    return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS


@app.route('/', methods = ['GET', 'POST'])
def index():
    if request.method=='POST':
        if 'xml_file' not in request.files or 'docx_file' not in request.files:
            flash('File not found')
            return redirect(request.url)
        else:
            xml = request.files['xml_file']
            docx = request.files['docx_file']
            if(is_file_allowed(xml.filename) and is_file_allowed(docx.filename)):
                flash('File is good')
                generate_output_table(docx, xml)

            else :
                flash('File is not good')
            return redirect(url_for('index', isReady=isOutputExist()))
    else:
        return render_template('home.html', isReady=isOutputExist())

@app.route('/download',methods = ['GET', 'POST'] )
def download():
    #belum tahu gimana cara manage upload download
    uploads = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'])

    # @after_this_request
    # def delete_output():
    #     os.remove(os.path.join(app.root_path, app.config['UPLOAD_FOLDER'], 'output.docx'))
    #     return redirect(url_for('index', isReady=isOutputExist()))

    return send_from_directory(directory=uploads, filename='output.docx')

@app.route('/clear_files')
def clear_files():
    os.remove(os.path.join(app.root_path, app.config['UPLOAD_FOLDER'], 'output.docx'))
    return redirect(url_for('index', isReady=isOutputExist()))


def generate_output_table(doc_input, xml_input):
    try:
        doc = Document(doc_input)
        xml_data = XMLReader(xml_input).getData()


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

        doc.save(os.path.join(app.root_path, app.config['UPLOAD_FOLDER'], 'output.docx'))

    except IndexError:
        print(f'Usage : python [doc file] [xml file] [saved file]')

def isOutputExist():
    return os.path.isfile(os.path.join(app.root_path, app.config['UPLOAD_FOLDER'], 'output.docx'))


if(__name__)=='__main__':
    app.run(debug=True)
