from tempfile import NamedTemporaryFile as tmpfile
from csv2xls import xls




def decode(s, encodings=('gbk', 'utf8')):
	for encoding in encodings:
		try:
			return s.decode(encoding)
		except UnicodeDecodeError:
			pass
	return s.decode('ascii', 'ignore')

tempcsv = tmpfile(delete = False)
tempxls = tmpfile(delete = False)

tempcsv.write('''
11111,21111
11111,21111
''')
tempcsv.close()
tempxls.close()

print tempcsv.name
this_instance = xls()
f = open('/tmp/fuck.csv')
this_instance.options.infile_names = ['A']
this_instance.options.outfile_name = 'B'
this_instance.options.set_default_options()
this_instance.options.check_options()
this_instance.process_csvs(decode(f.read()))


f = this_instance.csvs_2_xls()

g = open('/tmp/x.xls', 'w')
g.write(f.getvalue())
g.close()

#f = open('/tmp/fuck.csv')

#s = f.read()
#
#def decode(s, encodings=('gbk', 'utf8')):
#	for encoding in encodings:
#		try:
#			return s.decode(encoding)
#		except UnicodeDecodeError:
#			pass
#	return s.decode('ascii', 'ignore')
#
#
#print decode(s)


