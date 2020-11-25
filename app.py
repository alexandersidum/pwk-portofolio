from flask import render_template, session, Flask, Response, request, flash, url_for, redirect, send_from_directory, after_this_request,send_file
from flask_session import Session
from apscheduler.schedulers.background import BackgroundScheduler
from file_handler import FileHandler
import tempfile
import os
import atexit

# Exception handling belum
# Alert untuk error saat convert dan submitan kosong belum

app = Flask(__name__)
app.secret_key = '11d06ffa54be4e60b5f51dd1434296b0'
app.config['UPLOAD_FOLDER'] = 'output_files'
app.config['SESSION_TYPE'] = 'filesystem'
output_dir = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'])
Session(app)

def clean_up():
    for f in os.scandir(output_dir):
        os.unlink(f.path)

scheduler = BackgroundScheduler()
scheduler.add_job(func=clean_up, trigger="interval", minutes=2)
scheduler.start()


@app.route('/', methods = ['GET', 'POST'])
def index():
    fh = FileHandler()
    if request.method=='POST':
        if 'xml_file' not in request.files or 'docx_file' not in request.files:
            # File tidak diinput
            return redirect(request.url)
        else:
            xml = request.files['xml_file']
            docx = request.files['docx_file']
            if(fh.is_xml_file_allowed(xml.filename) and fh.is_docx_file_allowed(docx.filename)):                
                # Kalau file valid?
                session['fd'], session['fname'] = tempfile.mkstemp(suffix='.docx' , dir=output_dir)
                session['is_file_ready'] = fh.generate_output_table(docx_file=docx, xml_file=xml, t_fd=session.get('fd') ,t_fname=session.get('fname'))
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
        try :
            with open (session.get('fname'), 'rb') as f:
                data = f.readlines()
            os.remove(session.get('fname'))
            return Response(data, headers={
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition': 'attachment; filename=response_output.docx'
            })
        except:
            session['is_file_ready'] = False,
            return redirect(url_for('index', is_ready=session.get('is_file_ready'), is_file_error=session.get('is_file_error')))
    else :
        session['is_file_ready'] = False
        return redirect(url_for('index', is_ready=session.get('is_file_ready'), is_file_error=session.get('is_file_error')))

@app.route('/download-cth-docx')
def download_cth_portofolio():
   return send_from_directory(os.path.join(app.root_path, 'static'), filename='portfolio_sip.docx', as_attachment=True)

@app.route('/download-cth-xml')
def download_cth_nilai():
    return send_from_directory(os.path.join(app.root_path, 'static'), filename='DK4304_A_2018_1_36100.xml', as_attachment=True)

def is_file_ready():
    return session.get('fname') is not None and os.path.isfile(session.get('fname'))


atexit.register(lambda : scheduler.shutdown(wait=False))

if(__name__)=='__main__':
    app.run()
