# vim: ai ts=4 sts=4 et sw=4 ft=python
import os
import sys
root = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(root, 'site-packages'))


from bottle import Bottle, static_file, request , response, redirect
from csv2xls import xls
import sae
from StringIO import StringIO

def ensure_encoding(s, encodings=('gbk', 'utf8')):
    for encoding in encodings:
        try:
            return s.decode(encoding)
        except UnicodeDecodeError:
            pass
    return s.decode('ascii', 'ignore')

app = Bottle()

@app.error(404)
@app.route('/')
def index(error = None):
    return static_file('index.html', root = '.');


@app.route('/convert', method = 'post')
def convert():
    csvfile = request.files.csvfile
    if csvfile and csvfile.file and csvfile.filename.endswith('.csv'):

        xlsobj = xls()
        xlsobj.options.infile_names = ['FAKENAME']
        xlsobj.options.outfile_name = 'ANOTHERFAKENAME'
        xlsobj.options.set_default_options()
        xlsobj.options.check_options()
        xlsobj.process_csvs(ensure_encoding(csvfile.file.read().strip()))
        xlsfile = xlsobj.csvs_2_xls()
        data = xlsfile.getvalue()
        xlsfile.close()

        response.add_header('Content-Type', 'application/vnd.ms-excel')
        response.add_header('Content-Disposition', 'attachment; filename="%s"' % (csvfile.filename.rstrip('.csv') + '.xls'))

        return data
        #return xlsfile sae cant send file ...

    return redirect('/')

application = sae.create_wsgi_app(app)
