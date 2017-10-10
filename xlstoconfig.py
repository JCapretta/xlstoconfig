#!/usr/bin/python

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_to_tuple, range_boundaries
import argparse
from jinja2 import Environment, FileSystemLoader
import yaml
import sys

class Config:
    def __init__(self, yaml, wb=None):
        if 'workbook' in yaml:
            if not yaml['workbook']=='PARENT_WORKBOOK':
                self.wb = load_workbook(filename = yaml['workbook'], data_only=True)
            else:
                self.wb = wb
        else:
            self.wb = wb
        if 'children' in yaml:
        #create dictionary attributes of this object
            self.make_dict_attributes(yaml['children'])
        if 'values' in yaml:
        #create scalar attributes of this object
            for key in yaml['values']:
                setattr(self,key,yaml['values'][key])
        
    def read_attributes(self,yaml,attr):
    #read the yaml to get the pointers to a table within a spreadsheet. each row is a new child object within a dictionary attribute of a parent object
        wb,ws,columns = self.get_locations(yaml,self.wb)
        
        if ('autocolumn' in yaml and yaml['autocolumn']) or not 'columns' in yaml:
            autocolumns = self.auto_column(yaml,wb)   
            for key in autocolumns.keys():
                if key not in columns.keys():
                    columns[key]=autocolumns[key]
        if ws:
            range_string=self.get_range(yaml,wb)
            rows=ws[range_string]
            if ('autocolumn' in yaml and yaml['autocolumn']) or not 'columns' in yaml:
                autocolumns = self.auto_column(yaml,wb)   
                for key in autocolumns.keys():
                    if key not in columns.keys():
                        columns[key]=autocolumns[key]
            for row in rows:
                name = row[columns['name']].value
                if name not in getattr(self,attr):
                    getattr(self,attr).update({name:Config(yaml,wb)})
                self.make_class_attributes(getattr(self,attr)[name],row,columns)
        else:
            setattr(self,attr,Config(yaml,wb))
        
    def set_attributes(self,attrs):
    #create scalar attributes of this object
        for key in attrs.keys():
            setattr(self,key,attrs[key])
            
    def append_attributes(self,attrs):
    # create or append to a list attribute of this object
        for key in attrs.keys():
            if hasattr(self,key):
                setattr(self,key,getattr(self,key)+[attrs[key]])
            else:
                setattr(self,key,[attrs[key]])

    def make_dict_attributes(self,children):
    #create atrributes of this object that are to be dictionaries
        for key in children.keys():
            setattr(self,key+"_dict",{})
            self.read_attributes(children[key],key+"_dict")
            if isinstance(getattr(self,key+"_dict"),dict):
                setattr(self,key,getattr(self,key+"_dict").values())
            else:
                setattr(self,key,getattr(self,key+"_dict"))
                
    def make_class_attributes(self,obj,row,columns):
        rowvals={}
        rowlistvals={}
        if 'lists' in columns:
            for list in columns['lists']:
                rowlistvals[list]=row[int(columns[' '][list])].value
            obj.append_attributes(rowlistvals)
        for column in columns.keys():
            if not column=='lists':
                if columns[column] or columns[column]==0:#test if 0. zero is not null
                    rowvals[column]=row[int(columns[column])].value        
                    obj.set_attributes(rowvals)
  
    def get_locations(self,locations,wb=None):
        #load workbook
        columns={}
        ws=None
        if 'workbook' in locations:
            if not locations['workbook']=='PARENT_WORKBOOK':
                wb = load_workbook(filename = locations['workbook'], data_only=True) #data_only=True so you the values are read from the spreadhseet. not the formulas.
        ws = wb[locations['worksheet']]
        if 'columns' in locations:
            columns=locations['columns']
        return wb,ws,columns
    
    def auto_column(self,yaml,wb):
        #automatically generate the 'column' dictionary
        ws = wb[yaml['worksheet']]
        range_string = self.get_range(yaml,wb)
        cols={}
        #get header row
        first_col = range_string.split(":")[0][0]
        last_col = range_string.split(":")[1][0]
        top_row = int(range_string.split(":")[0][1])
        top_row = 1 if top_row == 1 else top_row-1 #if top row of range is already at the top don't decrement. otherwise decrement.
        header_range = first_col+str(top_row)+':'+last_col+str(top_row)
        #print header_range
        header = ws[header_range]
        #print "header len "+str(len(header[0]))
        for i in range(len(header[0])):
            cols[header[0][i].value]=i 
        cols['name']=ord(first_col.upper())-65
        return cols
        
    def get_range(self,yaml,wb):
        #get the range and return a range string
        if ('autorange' in yaml and yaml['autorange']) or not 'range' in yaml:          
            ws=wb[yaml['worksheet']]
            range_string = ws.calculate_dimension()
            min_col, min_row, max_col, max_row = range_boundaries(range_string)
            if max_row > min_row:
                min_row+=1
            range_string = '%s%d:%s%d' % (
            get_column_letter(min_col), min_row,
            get_column_letter(max_col), max_row
        )
        else:             
            range_string = yaml['range'] 
        return range_string
                 
def jinjaSkeleton(parent_name,yaml):
    #create a skeleton for your jinja template. good to use as a starting point if you don't have a jinja template already
    if 'children' in yaml:
        jinja_text=''
        for key in yaml['children'].keys():
            if 'workbook' in yaml['children'][key]:
                jinja_text += '{% for ' + key + ' in ' + parent_name + '.' + key + ' %}'+ '\n' + jinjaSkeleton(key,yaml['children'][key]) + '\n{% endfor %}\n'    
            else:
                jinja_text += '{% set ' + key + ' = ' + parent_name + '.' + key + ' %}'+ '\n' + jinjaSkeleton(key,yaml['children'][key])
        return jinja_text
    else:
        return ' '

class MyParser(argparse.ArgumentParser):
    def error(self, message):
    #override error method to print help and error message
        sys.stderr.write('error: %s\n' % message)
        self.print_help()
        sys.exit(2)
		
def get_args(args=sys.argv[1:]):
    parser = MyParser(description="John's config generator")
    parser.add_argument('file',help='yaml file')
    parser.add_argument('-t', '--template', help='Jinja2 template file')
    parser.add_argument('-j', dest='create_skeleton', help='Generate a skeleton for building a jinja2 template', action='store_true')
    if len(args)==0:
    #print help if no arguments
        parser.print_help()
        sys.exit(1)
    parsed_args = parser.parse_args(args)
    return parsed_args 
        
class Runner:
    def __init__(self, args=sys.argv[1:]):
        self.p_args=get_args(args)
        self.load_yaml()
        self.root='config'
        if self.p_args.create_skeleton:
            print self.jinjaSkeleton(self.root,self.locations)
        self.config=Config(self.locations)
        if self.p_args.template:
            print self.render_template()
        
    def load_yaml(self):
        if self.p_args.file:
            try:
                with open(self.p_args.file, 'r') as file_descriptor:
                    self.locations = yaml.load(file_descriptor)
            except Exception as ex:
                template = "An exception of type {0} occurred. Arguments:\n{1!r}"
                message = template.format(type(ex).__name__, ex.args)
                print message
                sys.exit(4) 
                
    def jinjaSkeleton(self,parent_name,yaml):
        #create a skeleton for your jinja template. good to use as a starting point if you don't have a jinja template already
        if 'children' in yaml:
            jinja_text=''
            for key in yaml['children'].keys():
                if 'workbook' in yaml['children'][key]:
                    jinja_text += '{% for ' + key + ' in ' + parent_name + '.' + key + ' %}'+ '\n' + jinjaSkeleton(key,yaml['children'][key]) + '\n{% endfor %}\n'    
                else:
                    jinja_text += '{% set ' + key + ' = ' + parent_name + '.' + key + ' %}'+ '\n' + jinjaSkeleton(key,yaml['children'][key])
            return jinja_text
        else:
            return ' '            

    def render_template(self):
        env = Environment(loader=FileSystemLoader('./'))
        template = env.get_template(self.p_args.template)
        return template.render(config = self.config)
        
	
def main():

    my_runner=Runner()

if __name__ == '__main__':
  main()
