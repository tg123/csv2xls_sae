from tempfile import NamedTemporaryFile as tmpfile
from csv2xls import xls




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
this_instance.options.infile_names = ['/tmp/tmp1ht51X']
this_instance.options.outfile_name = tempxls.name
this_instance.options.set_default_options()
this_instance.options.check_options()
this_instance.process_csvs()
this_instance.csvs_2_xls()

