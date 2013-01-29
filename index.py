from bottle import route, run, template, static_file, request , response
from tempfile import NamedTemporaryFile as tmpfile

from csv2xls import xls
from os import remove


@route('/')
def index():
    return static_file('index.html', root = '.');


@route('/convert', method = 'post')
def convert():
    csvfile = request.files.csvfile
    if csvfile and csvfile.file and csvfile.filename.endswith('.csv'):
    	tempcsv = tmpfile(delete = False)
	tempxls = tmpfile(delete = False)
	tempcsv.write(csvfile.file.read())
	tempcsv.close()
	tempxls.close()

        xlsobj = xls()
        xlsobj.options.infile_names = [tempcsv.name]
        xlsobj.options.outfile_name = tempxls.name
        xlsobj.options.set_default_options()
        xlsobj.options.check_options()
        xlsobj.process_csvs()
        xlsobj.csvs_2_xls()

	try:
	    return static_file(tempxls.name, root = '/', download = csvfile.filename.rstrip('.csv') + '.xls')
	finally:
	    remove(tempcsv.name)
	    remove(tempxls.name)

    return 'missiing'

run(host='localhost', port=9586)
