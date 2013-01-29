# vim: ai ts=4 sts=4 et sw=4 ft=python

# Copyright 2008 by Guido van Steen. Consult the LICENSE file of csv2xls for the terms of use. 
# Modified by tgic for Sina App Engine

import sys 
import os
import optparse 
import copy
import warnings 
from StringIO import StringIO

#name_of_pyexcelerator_module_sub_directory = os.sep + "pyexcelerator"
#name_of_pyexcelerator_module_directory = os.path.abspath(__file__)[:os.path.abspath(__file__).rfind(os.sep)] + name_of_pyexcelerator_module_sub_directory 
#sys.path.append(name_of_pyexcelerator_module_directory)

def errmsg(msg):
    pass

try:
    import pyExcelerator as pyexcelerator 
except ImportError:
    errmsg("Python module pyexcelerator cannot be imported")

#sys.path.remove(name_of_pyexcelerator_module_directory)

#name_of_afm_module_sub_directory = os.sep + "afm"
#name_of_afm_module_directory = os.path.abspath(__file__)[:os.path.abspath(__file__).rfind(os.sep)] + name_of_afm_module_sub_directory 
#sys.path.append(name_of_afm_module_directory)

#try:
#    import afm
#    pass
#except ImportError:
#    errmsg("Python module afm cannot be imported")

#sys.path.remove(name_of_afm_module_directory)

warnings.simplefilter('error', DeprecationWarning)

external_non_available_string = "NA" 

format_seperator_string = ":" 
formats_seperator_string = "::"

default_format_string = "default" + format_seperator_string + "general"
default_transpose_formatting = "false" 
default_font_name = "Helvetica" 
#name_of_adobe_fonts_directory = os.sep + "Adobe-Core35_AFMs-314" + os.sep 
#default_font_metrics_file = os.path.abspath(__file__)[:os.path.abspath(__file__).rfind(os.sep)] + name_of_adobe_fonts_directory + default_font_name + ".afm"
default_column_width = "48"
default_assume_rownames = "false"
default_assume_colnames = "true"
#default_convert_to_floats = "default" 
default_convert_to_floats = "never" 
default_separator = "," # Support for different separators to be added later 

def comma2list(string):
    if string == "": 
        return []
    else: 
        return string.split(",")
    
#def string_width(string, metrics_file_name):
#    try: 
#        fh = open(metrics_file_name)
#    except IOError: 
#        errmsg("Font metrics file cannot be opened")
#    afm_object = afm.AFM(fh)
#    fh.close()
#    f_string_width = afm_object.string_width_height(string)[0]
#    return float(str(f_string_width))/100

class csv_options(object): 
    def __init__(self):
        self.infile_name = ""
        self.sheet_name = ""
        self.format = ""
        self.transpose_formatting = ""
        self.font_name = ""
        self.font_metrics_file = ""
        self.column_width = ""
        self.assume_rowname = ""
        self.assume_colname = ""
        self.convert_to_floats = "" 

    def set_infile_name(self, name): 
        self.infile_name = name

    def set_sheet_name(self, name): 
        self.sheet_name = name 
    
    def set_format(self, format): 
        self.format = format

    def set_transpose_formatting(self, bool_string): 
        self.transpose_formatting = bool_string

    def set_font_name(self, name): 
        self.font_name = name

    def set_font_metrics_file(self, file_name): 
        self.font_metrics_file = file_name

    def set_column_width(self, size): 
        self.column_width = size

    def set_assume_rowname(self, bool_string): 
        self.assume_rowname = bool_string 

    def set_assume_colname(self, bool_string): 
        self.assume_colname = bool_string 

    def set_convert_to_float(self, string): 
        self.convert_to_float = string 

class csv(object): 
    def __init__(self): 
        self.contents = []
        self.column_names = []
        self.row_names = []
        self.options = csv_options() 
        self.csvdata = ''

    def make_empty(self): 
        self.__init__()

    def get_column_names_from_infile(self):
        file_name = self.options.infile_name 
        fh = StringIO(self.csvdata)
        next_line = fh.readline()
        self.column_names = next_line.replace("\n","").split(",")
        fh.close() 
        
    def get_row_names_from_infile(self):
        file_name = self.options.infile_name 
        fh = StringIO(self.csvdata)
        next_line = fh.readline()
        while next_line != "": 
            next_list_of_floats = []
            split_next_line = next_line.split(",") 
            self.row_names = self.row_names + [split_next_line[0]] 
            next_line = fh.readline()
        fh.close() 

    def get_data_from_infile(self, assume_rownames_argument, assume_colnames_argument, convert_to_float):
        file_name = self.options.infile_name 
        fh = StringIO(self.csvdata)
        if assume_colnames_argument.lower() in ["true","1"]: 
            next_line = fh.readline() # Skip to next line
        next_line = fh.readline() 
        while next_line != "": 
            next_list_of_elements = []
            split_next_line = next_line.split(",")
            if assume_rownames_argument.lower() in ["false","0"]: # 0-th column is not part of the data 
                split_next_line = split_next_line[0:len(split_next_line)] 
            else: 
                split_next_line = split_next_line[1:len(split_next_line)] 
            for cell in split_next_line: 
                next_element = cell.replace("\n","")
                next_list_of_elements = next_list_of_elements + [next_element]
            self.contents = self.contents + [next_list_of_elements]
            next_line = fh.readline()
        fh.close() 

    def set_data_formats(self):
        # Find the default format
        split_formats = self.options.format.split(formats_seperator_string)
        if "default" in self.options.format:
            for format in split_formats: 
                split_format = format.split(format_seperator_string) 
                if split_format[0] == "default":
                    default_format = split_format[1] 
                    split_formats.remove(format)
        else:             
            default_format = "general"
        # Find the number of columns or rows 
        if self.options.transpose_formatting.lower() in ["false","0"]: 
            number_of_columns_or_rows = len(self.row_names) 
        else:
            number_of_columns_or_rows = len(self.column_names) 
        # Construct a temporary dictionary containing the specified formats 
        temp_formats = {}
        for format in split_formats: 
            try: 
                x = int(format.split(format_seperator_string)[0])
                y = str(format.split(format_seperator_string)[1])
            except ValueError: 
                errmsg("Invalid formats argument")
            except IndexError: 
                errmsg("Invalid formats argument")
            if x < 0: 
                x = number_of_columns_or_rows - x 
            temp_formats[x] = y
        # Fill out self.data_formats 
        self.data_formats = copy.deepcopy(self.contents) # Without "deep copying" self.data_formats would be another name for self.contents 
        try: 
            for col_count in range(len(self.contents[0])): 
                for row_count in range(len(self.contents)):
                    self.data_formats[row_count][col_count] = default_format
                    if self.options.transpose_formatting.lower() in ["false","0"]: 
                        if (temp_formats.has_key(col_count)):
                            self.data_formats[row_count][col_count] = temp_formats[col_count]
                    else: 
                        if (temp_formats.has_key(row_count)):
                            self.data_formats[row_count][col_count] = temp_formats[row_count]
        except IndexError: 
            errmsg("Infile name is not a valid csv file")

class xls_options(object):
    def __init__(self):
        self.infile_names = []
        self.outfile_name = ""
        self.sheet_names = []
        self.formats = []
        self.transpose_formattings = [] 
        self.font_names = []
        self.font_metrics_files = []
        self.column_widths = []
        self.assume_rownames = []
        self.assume_colnames = []
        self.convert_to_floats = []

    def get_options(self, options):
        self.infile_names = comma2list(options.__dict__["infile_names"])
        self.outfile_name = options.__dict__["outfile_name"] # Only one outfile is allowed 
        self.sheet_names = comma2list(options.__dict__["sheet_names"]) 
        self.formats = comma2list(options.__dict__["formats"])
        self.transpose_formattings = comma2list(options.__dict__["transpose_formattings"])
        self.font_names = comma2list(options.__dict__["font_names"])
        self.font_metrics_files = comma2list(options.__dict__["font_metrics_files"])
        self.column_widths = comma2list(options.__dict__["column_widths"])
        self.assume_rownames = comma2list(options.__dict__["assume_rownames"])
        self.assume_colnames = comma2list(options.__dict__["assume_colnames"])
        self.convert_to_floats = comma2list(options.__dict__["convert_to_floats"])

    def set_default_options(self):
        if self.sheet_names == []: 
            for i in range(len(self.infile_names)): 
                #self.sheet_names = self.sheet_names + [os.path.splitext(os.path.split(self.infile_names[i])[1])[0]]
                self.sheet_names = self.sheet_names + ['Sheet' + str(i + 1)]
        if self.formats == []: 
            for i in range(len(self.infile_names)): 
                self.formats = self.formats + [default_format_string]
        if self.transpose_formattings == []: 
            for i in range(len(self.infile_names)): 
                self.transpose_formattings = self.transpose_formattings + [default_transpose_formatting]
        if self.font_names == []: 
            for i in range(len(self.infile_names)): 
                self.font_names = self.font_names + [default_font_name]
        #if self.font_metrics_files == []:
        #    for i in range(len(self.infile_names)): 
        #        self.font_metrics_files = self.font_metrics_files + [default_font_metrics_file]
        if self.column_widths == []: 
            for i in range(len(self.infile_names)): 
                self.column_widths = self.column_widths + [default_column_width]
        if self.assume_rownames == []: 
            for i in range(len(self.infile_names)): 
                self.assume_rownames = self.assume_rownames + [default_assume_rownames]
        if self.assume_colnames == []: 
            for i in range(len(self.infile_names)): 
                self.assume_colnames = self.assume_colnames + [default_assume_colnames]
        if self.convert_to_floats == []: 
            for i in range(len(self.infile_names)): 
                self.convert_to_floats = self.convert_to_floats + [default_convert_to_floats]

        if len(self.formats) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.formats = self.formats + [self.formats[0]]
        if len(self.transpose_formattings) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.transpose_formattings = self.transpose_formattings + [self.transpose_formattings[0]]
        if len(self.font_names) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.font_names = self.font_names + [self.font_names[0]]
        if len(self.font_metrics_files) == 1:
            for i in range(1,len(self.infile_names)): 
                self.font_metrics_files = self.font_metrics_files + [self.font_metrics_files[0]]
        if len(self.column_widths) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.column_widths = self.column_widths + [self.column_widths[0]]
        if len(self.assume_rownames) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.assume_rownames = self.assume_rownames + [self.assume_rownames[0]]
        if len(self.assume_colnames) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.assume_colnames = self.assume_colnames + [self.assume_colnames[0]]
        if len(self.convert_to_floats) == 1: 
            for i in range(1,len(self.infile_names)): 
                self.convert_to_floats = self.convert_to_floats + [self.convert_to_floats[0]]

    def check_options(self):
        if self.infile_names == []:
            errmsg("No infile argument")
        #for f in self.infile_names: 
        #    try: 
        #        fh = open(f,"r")
        #    except IOError: 
        #        errmsg("Infile name(s) do not exist")
        #    else: 
        #        fh.close()
        if self.outfile_name == "":
            errmsg("No outfile argument")
        elif "," in self.outfile_name: 
            errmsg("Bad outfile argument")
        for f in self.infile_names: 
            if os.path.abspath(f) == os.path.abspath(self.outfile_name): 
                errmsg("Infile name(s) coincide with outfile name")
        #try: 
        #    fh = open(self.outfile_name, "w")
        #except IOError: 
        #    errmsg("Outfile name cannot be created")
        #else: 
        #    fh.close()

class xls(object):
    def __init__(self):
        self.options = xls_options()
        self.csv_objects_to_be_processed = []
        self.current_csv_object = csv()
        self.xls_object = pyexcelerator.Workbook()

    def process_csvs(self, csvdata): 
        for i in range(len(self.options.infile_names)):
            self.current_csv_object.options.set_infile_name(self.options.infile_names[i])
            try: 
                self.current_csv_object.options.set_sheet_name(self.options.sheet_names[i])
            except IndexError: 
                errmsg("Invalid sheets argument")
            try: 
                self.current_csv_object.options.set_format(self.options.formats[i])
            except IndexError: 
                errmsg("Invalid formats argument")
            try: 
                self.current_csv_object.options.set_transpose_formatting(self.options.transpose_formattings[i]) 
            except IndexError: 
                errmsg("Invalid transpose_formattings argument")
            try: 
                self.current_csv_object.options.set_font_name(self.options.font_names[i]) 
            except IndexError: 
                errmsg("Invalid font_names argument")
            #try: 
            #    self.current_csv_object.options.set_font_metrics_file(self.options.font_metrics_files[i])
            #except IndexError: 
            #    errmsg("Invalid font_metrics_files argument")
            try: 
                self.current_csv_object.options.set_column_width(self.options.column_widths[i]) 
            except IndexError: 
                errmsg("Invalid column_widths argument")
            try: 
                self.current_csv_object.options.set_assume_rowname(self.options.assume_rownames[i]) 
            except IndexError: 
                errmsg("Invalid assume_rownames argument")
            try: 
                self.current_csv_object.options.set_assume_colname(self.options.assume_colnames[i]) 
            except IndexError: 
                errmsg("Invalid assume_colnames argument")
            try: 
                self.current_csv_object.options.set_convert_to_float(self.options.convert_to_floats[i]) 
            except IndexError: 
                errmsg("Invalid convert_to_floats argument")

            self.current_csv_object.csvdata = csvdata
            self.current_csv_object.get_column_names_from_infile()
            self.current_csv_object.get_row_names_from_infile()
            assume_rowname_argument = self.current_csv_object.options.assume_rowname
            assume_colname_argument = self.current_csv_object.options.assume_colname
            convert_to_float_argument = self.current_csv_object.options.convert_to_float
            self.current_csv_object.get_data_from_infile(assume_rowname_argument, assume_colname_argument, convert_to_float_argument)
            self.current_csv_object.set_data_formats()
            self.csv_objects_to_be_processed = self.csv_objects_to_be_processed + [copy.copy(self.current_csv_object)] # "Ordinary" copy seems sufficient here. 
            self.current_csv_object.make_empty()

    def export_current_csv_to_xls(self):
        current_style = pyexcelerator.XFStyle()
        current_style.font.name = self.current_csv_object.options.font_name
        current_style.alignment = pyexcelerator.Alignment()
        current_style.alignment.horz = pyexcelerator.Alignment.HORZ_RIGHT
        number_of_columns = len(self.current_csv_object.contents[0]) 
        number_of_rows = len(self.current_csv_object.contents) 
        assume_rowname = self.current_csv_object.options.assume_rowname 
        assume_colname = self.current_csv_object.options.assume_colname 
        convert_to_float = self.current_csv_object.options.convert_to_float 
        for row_count in range(number_of_rows): 
            for col_count in range(number_of_columns):
                current_style.num_format_str = self.current_csv_object.data_formats[row_count][col_count]
                cell = self.current_csv_object.contents[row_count][col_count] 
                if assume_colname in ["false","0"]: 
                    r_xls = row_count # Due to NO column names 
                else: 
                    r_xls = row_count + 1 # Due to column names 
                if assume_rowname in ["false", "0"]: 
                    c_xls = col_count # Due to NO row names 
                else: 
                    c_xls = col_count + 1 # Due to row names 
                if convert_to_float in ["default"]: 
                    if str(cell).strip().replace(" ","").replace(".","").replace("-","").replace("+","").replace("E","").replace("e","").isdigit(): 
                        self.xls_object.xls_sheet.write(r_xls,c_xls,float(cell),current_style) 
                    else: # Not a digit 
                        self.xls_object.xls_sheet.write(r_xls,c_xls,cell,current_style) 
                if convert_to_float in ["always"]:
                    if cell.strip().replace(".","").isdigit(): 
                        self.xls_object.xls_sheet.write(r_xls,c_xls,cell,current_style) 
                    else: 
                        self.xls_object.xls_sheet.write(r_xls,c_xls,external_non_available_string,current_style) 
                if convert_to_float in ["never"]: 
                        self.xls_object.xls_sheet.write(r_xls,c_xls,cell,current_style) 
                    
    def export_column_names_to_xls(self):
        current_style = pyexcelerator.XFStyle()
        current_style.font.name = self.current_csv_object.options.font_name
        current_style.alignment = pyexcelerator.Alignment()
        current_style.alignment.horz = pyexcelerator.Alignment.HORZ_RIGHT
        for i in range(len(self.current_csv_object.column_names)): 
            self.xls_object.xls_sheet.write(0, i, self.current_csv_object.column_names[i].replace('"',''),current_style)
    
    def export_row_names_to_xls(self):
        current_style = pyexcelerator.XFStyle()
        current_style.font.name = self.current_csv_object.options.font_name
        current_style.alignment = pyexcelerator.Alignment()
        current_style.alignment.horz = pyexcelerator.Alignment.HORZ_LEFT
        for i in range(len(self.current_csv_object.row_names)): 
            self.xls_object.xls_sheet.write(i, 0, self.current_csv_object.row_names[i].replace('"',''),current_style)

    def save_xls_to_file(self):
        f = StringIO()
        self.xls_object.save(f)
        f.seek(0)
        return f

    def remove_xls(self):
        os.remove(self.options.outfile_name)

    def export_column_widths_to_xls_sheet(self): 
        #if self.current_csv_object.options.assume_rowname in ["false", "0"]: 
            for i in range(0,len(self.current_csv_object.column_names)): 
                self.xls_object.xls_sheet.col(i).width = int((float(self.current_csv_object.options.column_width) - 1) * 48 + 96)
        #else: 
        #    col_width = 0 
        #    for i in self.current_csv_object.row_names: 
        #        try: 
        #            current_col_width = string_width(i,self.current_csv_object.options.font_metrics_file)
        #        except IOError: 
        #            errmsg("Font metrics file(s) cannot be opened")
        #        except RuntimeError:
        #            errmsg("Invalid font metrics file(s)")
        #        except KeyError:
        #            errmsg("Unable to determine width of first column")
        #        col_width = max(current_col_width, col_width)
        #    self.xls_object.xls_sheet.col(0).width = int(col_width * 48 + 96)
        #    for i in range(1,len(self.current_csv_object.column_names)): 
        #        self.xls_object.xls_sheet.col(i).width = int((float(self.current_csv_object.options.column_width) - 1) * 48 + 96)

    def csvs_2_xls(self): 
        for i in range(len(self.options.infile_names)):
            self.current_csv_object = self.csv_objects_to_be_processed[i]
            self.xls_object.xls_sheet = self.xls_object.add_sheet(self.current_csv_object.options.sheet_name)
            self.export_column_widths_to_xls_sheet()
            self.export_current_csv_to_xls()
            if self.current_csv_object.options.assume_colname in ["true","1"]: 
                self.export_column_names_to_xls()
            if self.current_csv_object.options.assume_rowname in ["true","1"]: 
                self.export_row_names_to_xls()
        try: 
            return self.save_xls_to_file()
        except DeprecationWarning: 
            self.remove_xls()
            errmsg("Too large column_width argument")

def main():
    parser = optparse.OptionParser()
    parser.add_option("-i", "--infile_names", dest="infile_names", default="", help="set infilenames")
    parser.add_option("-o", "--outfile_name", dest="outfile_name", default="", help="set outfilename")
    parser.add_option("-s", "--sheet_names", dest="sheet_names", default="", help="set sheetnames")
    parser.add_option("-f", "--formats", dest="formats", default="", help="set colum or row formats")
    parser.add_option("-t", "--transpose_formattings", dest="transpose_formattings", default="", help="set transpose formattings")
    parser.add_option("-n", "--font_names", dest="font_names", default="", help="set font names")
    parser.add_option("-m", "--font_metrics_files", dest="font_metrics_files", default="", help="set font metrics files")
    parser.add_option("-w", "--column_widths", dest="column_widths", default="", help="set colum widths")
    parser.add_option("-r", "--assume_rownames", dest = "assume_rownames", default="", help="set handling of rownames")
    parser.add_option("-x", "--assume_colnames", dest = "assume_colnames", default="", help="set handling of colnames")
    parser.add_option("-c", "--convert_to_floats", dest = "convert_to_floats", default="", help="set conversion mode")

    (options, args) = parser.parse_args()
    if len(args) != 0:
        errmsg("No support for unnamed arguments")
    this_instance = xls()
    this_instance.options.get_options(options)
    this_instance.options.set_default_options()
    this_instance.options.check_options()
    this_instance.process_csvs()
    this_instance.csvs_2_xls()

if __name__ == '__main__':
    main()

# Planned for future versions: 
# * get rid of bugs 
# * add support for csv-like files using separators different from ",", such as tab-delimted files. 
# * add support for xls dates. 
# * add support for the formatting of single cells.

