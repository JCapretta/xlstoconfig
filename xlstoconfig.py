#!/usr/bin/python

"""
Module docstring
"""

import argparse
import sys
import yaml
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from jinja2 import Environment, FileSystemLoader

class Config(object):
    """docstring here"""
    def __init__(self, yml, workbook=None):
        self.workbook = workbook
        if 'workbook' in yml and not yml['workbook'] == 'PARENT_WORKBOOK':
            self.workbook = load_workbook(filename=yml['workbook'], data_only=True)
        if 'children' in yml:
        #create dictionary attributes of this object
            self.make_dict_attributes(yml['children'])
        if 'values' in yml:
        #create scalar attributes of this object
            for key in yml['values']:
                setattr(self, key, yml['values'][key])

    def read_attributes(self, yml, attr):
        """read the yaml to get the pointers to a table within a spreadsheet.
        Each row is a new child object within a dictionary attribute of a parent object"""
        workbook, worksheet, columns = get_locations(yml, self.workbook)
        if worksheet:
            range_string = get_range(yml, workbook)
            rows = worksheet[range_string]
            if ('autocolumn' in yml and yml['autocolumn']) or not 'columns' in yml:
                autocolumns = auto_column(yml, workbook)
                columns.update({x: autocolumns[x] for x in autocolumns if x not in columns})
            for row in rows:
                name = row[columns['name']].value
                if name not in getattr(self, attr):
                    getattr(self, attr).update({name:Config(yml, workbook)})
                make_class_attributes(getattr(self, attr)[name], row, columns)
        else:
            setattr(self, attr, Config(yml, workbook))

    def set_attributes(self, attrs):
        """create scalar attributes of this object"""
        for key in attrs.keys():
            setattr(self, key, attrs[key])

    def append_attributes(self, attrs):
        """create or append to a list attribute of this object"""
        for key in attrs.keys():
            if hasattr(self, key):
                setattr(self, key, getattr(self, key)+[attrs[key]])
            else:
                setattr(self, key, [attrs[key]])

    def make_dict_attributes(self, children):
        """create atrributes of this object that are to be dictionaries"""
        for key in children.keys():
            setattr(self, key+"_dict", {})
            self.read_attributes(children[key], key+"_dict")
            if isinstance(getattr(self, key+"_dict"), dict):
                setattr(self, key, getattr(self, key+"_dict").values())
            else:
                setattr(self, key, getattr(self, key+"_dict"))

def auto_column(yml, workbook):
    """automatically generate the 'column' dictionary"""
    worksheet = workbook[yml['worksheet']]
    range_string = get_range(yml, workbook)
    cols = {}
    first_col = range_string.split(":")[0][0]
    last_col = range_string.split(":")[1][0]
    top_row = int(range_string.split(":")[0][1])
    #if top row of range is already at the top don't decrement. otherwise decrement.
    top_row = 1 if top_row == 1 else top_row-1
    header_range = first_col+str(top_row)+':'+last_col+str(top_row)
    header = worksheet[header_range]
    for i in range(len(header[0])):
        cols[header[0][i].value] = i
    cols['name'] = ord(first_col.upper())-65
    return cols

def get_range(yml, workbook):
    """get the range and return a range string"""
    if ('autorange' in yml and yml['autorange']) or not 'range' in yml:
        worksheet = workbook[yml['worksheet']]
        range_string = worksheet.calculate_dimension()
        min_col, min_row, max_col, max_row = range_boundaries(range_string)
        if max_row > min_row:
            min_row += 1
        range_string = '%s%d:%s%d' % (
            get_column_letter(min_col), min_row,
            get_column_letter(max_col), max_row
            )
    else:
        range_string = yml['range']
    return range_string

def make_class_attributes(obj, row, columns):
    """docstring"""
    rowvals = {}
    rowlistvals = {}
    if 'lists' in columns:
        rowlistvals[list] = {row[int(columns[' '][list])].value for list in columns['lists']}
        obj.append_attributes(rowlistvals)
    for column in columns.keys():
        if not column == 'lists':
            if columns[column] or columns[column] == 0:#test if 0. zero is not null
                rowvals[column] = row[int(columns[column])].value
                obj.set_attributes(rowvals)

def get_locations(locations, workbook=None):
    """load workbook"""
    columns = {}
    if 'workbook' in locations:
        if not locations['workbook'] == 'PARENT_WORKBOOK':
            #data_only=True so only the values are read from the spreadhseet. not the formulas.
            workbook = load_workbook(filename=locations['workbook'], data_only=True)
    worksheet = workbook[locations['worksheet']]
    if 'columns' in locations:
        columns = locations['columns']
    return workbook, worksheet, columns

class MyParser(argparse.ArgumentParser):
    """This class inherits from ArgumentParser and redefines error()"""
    def error(self, message):
        """override error method to print help and error message"""
        sys.stderr.write('error: %s\n' % message)
        self.print_help()
        sys.exit(2)

def get_args():
    """parse arguments from command line"""
    parser = MyParser(description="John's config generator")
    parser.add_argument('file', help='yaml file')
    parser.add_argument('-t', '--template', help='Jinja2 template file')
    parser.add_argument('-j',
                        dest='create_skeleton',
                        help='Generate a skeleton for building a jinja2 template',
                        action='store_true')
    if not sys.argv:
    #print help if no arguments
        parser.print_help()
        sys.exit(1)
    parsed_args = parser.parse_args()
    return parsed_args

class Runner(object):
    """this class could just be put into main()"""
    def __init__(self):
        """init for runner Class"""
        self.p_args = get_args()
        self.load_yaml()
        self.root = 'config'
        if self.p_args.create_skeleton:
            print jinja_skeleton(self.root, self.locations)
        self.config = Config(self.locations)
        if self.p_args.template:
            print self.render_template()

    def load_yaml(self):
        """load yaml file into memory"""
        if self.p_args.file:
            with open(self.p_args.file, 'r') as file_descriptor:
                self.locations = yaml.load(file_descriptor)

    def render_template(self):
        """Render the jinja2 template"""
        env = Environment(loader=FileSystemLoader('./'))
        template = env.get_template(self.p_args.template)
        return template.render(config=self.config)

def jinja_skeleton(parent_name, my_yaml):
    """Create a skeleton for your jinja template.
    Good to use as a starting point if you don't have a jinja template already"""
    if 'children' in my_yaml:
        jinja_text = ''
        for key in my_yaml['children'].keys():
            if 'workbook' in my_yaml['children'][key]:
                jinja_text += '{% for ' + key + ' in ' + parent_name + '.' + key + ' %}'
                jinja_text += '\n' + jinja_skeleton(key, my_yaml['children'][key])
                jinja_text += '\n{% endfor %}\n'
            else:
                jinja_text += '{% set ' + key + ' = ' + parent_name + '.' + key + ' %}'
                jinja_text += '\n' + jinja_skeleton(key, my_yaml['children'][key])
        return jinja_text
    return ' '

def main():
    """Main fnction"""
    Runner()

if __name__ == '__main__':
    main()
