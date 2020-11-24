from flask import render_template, session, Flask, Response, request, flash, url_for, redirect, send_from_directory, after_this_request,send_file
from flask_session import Session
from file_handler import FileHandler
import tempfile
import os


# BUGG
# backup di app5
# Exception handling belum
# Temp file masih belum terdelete jika tidak didownload !!
# menggunakan home test
# Named temp file?

app = Flask(__name__)
app.secret_key = '11d06ffa54be4e60b5f51dd1434296b0'
app.config['UPLOAD_FOLDER'] = 'output_files'
app.config['SESSION_TYPE'] = 'filesystem'
output_dir = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'])
Session(app)
ALLOWED_EXTENSIONS = {'xml', 'docx'}


def is_file_allowed(filename):
    return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS


@app.route('/', methods = ['GET', 'POST'])
def index():
    if request.method=='POST':
        if 'xml_file' not in request.files or 'docx_file' not in request.files:
            # File tidak diinput
            flash('File not found')
            return redirect(request.url)
        else:
            xml = request.files['xml_file']
            docx = request.files['docx_file']
            if(is_file_allowed(xml.filename) and is_file_allowed(docx.filename)):                
                # Kalau file valid?
                session['fd'], session['fname'] = tempfile.mkstemp(suffix='.docx' , dir=output_dir)
                session['is_file_ready'] = FileHandler(docx_file=docx, xml_file=xml).generate_output_table(session.get('fname'))
                session['is_file_error'] = False
            else :
                # Kalau file tidak valid
                session['is_file_error'] = True
                session['is_file_ready'] = False
            return redirect(url_for('index', is_ready=session.get('is_file_ready'), is_file_error=session.get('is_file_error')))
    else:
        return render_template('home.html', is_ready=session.get('is_file_ready'), is_file_error=session.get('is_file_error'))

@app.route('/download')
def download():
    if is_file_ready():
        with open (session.get('fname'), 'rb') as f:
            data = f.readlines()
        os.close(session.get('fd'))
        os.remove(session.get('fname'))
        return Response(data, headers={
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=response_output.docx'
        })
    else :
        session['is_file_ready'] = False
        return redirect(url_for('index', is_ready=session.get('is_file_ready'), is_file_error=session.get('is_file_error')))

def is_file_ready():
    return session.get('fname') is not None and os.path.isfile(session.get('fname'))

# def clean_up():
#     for f in os.scandir(output_dir):
#         os.remove(f.path)


if(__name__)=='__main__':
    app.run()
