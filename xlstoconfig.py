#!/usr/bin/python

from openpyxl import load_workbook
import argparse
from jinja2 import Environment, FileSystemLoader
import yaml

class PanoramaConfig:
    def __init__(self, yaml, wb=None):
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
        self.wb,ws,columns = self.get_locations(yaml,self.wb)
        if ws:
            for row in ws.iter_rows(min_col=yaml['min_col'],min_row=yaml['min_row']or ws.min_row, max_col=yaml['max_col']or ws.max_column, max_row=yaml['max_row']or ws.max_row):
                name = row[columns['name']].value
                if name not in getattr(self,attr):
                    getattr(self,attr).update({name:PanoramaConfig(yaml,self.wb)})
                self.make_class_attributes(getattr(self,attr)[name],row,columns)
        else:
            setattr(self,attr,PanoramaConfig(yaml,self.wb))
        
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
                rowlistvals[list]=row[int(columns['lists'][list])].value
            obj.append_attributes(rowlistvals)
        for column in columns.keys():
            if not column=='lists':
                if columns[column] or columns[column]==0:#test if 0. zero is not null
                    rowvals[column]=row[int(columns[column])].value        
                    obj.set_attributes(rowvals)

        
    def get_locations(self,locations,wb=None):
        #load workbook
        columns=None
        #wb=None
        ws=None
        if 'workbook' in locations:
            if not locations['workbook']=='PARENT_WORKBOOK':
                wb = load_workbook(filename = locations['workbook'], data_only=True) #data_only=True so you the values are read from the spreadhseet. not the formulas.
            ws = wb[locations['worksheet']]
        if 'columns' in locations:
            columns=locations['columns']
            
        return wb,ws,columns
        

                
                
def jinjaSkeleton(parent_name,yaml):
    #create a skeleton for your jinja template. good to use as a starting point if you don't have a jinja template already
    if 'children' in yaml:
        jinja_text=''
        for key in yaml['children'].keys():
            if 'workbook' in yaml['children'][key]:
                jinja_text += '{% for ' + key + ' in ' + parent_name + '.' + key + '.values() %}'+ '\n' + jinjaSkeleton(key,yaml['children'][key]) + '\n{% endfor %}\n'    
            else:
                jinja_text += '{% set ' + key + ' = ' + parent_name + '.' + key + ' %}'+ '\n' + jinjaSkeleton(key,yaml['children'][key])
        return jinja_text
    else:
        return ' '

                
def main():
    
    #arguments
    parser = argparse.ArgumentParser(description="John's config generator")
    parser.add_argument('-f', '--file',help='config workbook')
    parser.add_argument('-t', '--template')

    args = parser.parse_args()
    
    #yaml
    
    with open(args.file, 'r') as file_descriptor:
        locations = yaml.load(file_descriptor)
    
    # print the jinja skeleton. 
    print jinjaSkeleton('config',locations)

    #initialize switch config
    config = PanoramaConfig(locations)
    
    
    env = Environment(loader=FileSystemLoader('./'))
    template = env.get_template(args.template)
    print template.render(config = config)
    

if __name__ == '__main__':
  main()
