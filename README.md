**xlstoconfig**
-----------

**Overview** 

xlstoconfig takes data from an excel file and presents it for use with a Jinja2 template. 

**Components:**

 - xlstoconfig.py: python script
 - Jinja2 template
 - Excel file: this holds data the you want to use in your Jinja2 template

xlstoconfig reads each sheet within a xls workbook and passes the data to jinja2 as a list of dictionaries. The name of the list is taken from the name of the workseet. Each row from the worksheet is a dictionary within the list. The top row of the worksheet defines the keys of the dictionaries. 

Example use cases:

 - Generating Cisco config
 - Creating a xml file 

**Quick Jinja2 introduction**
-----------------------------

A basic Python script using Jinja2 involves:
 
 - Defining/loading a Jinja2 template
 - Passing data to Jinja2 for use in the template

The data object that you pass to Jinja2 is referenced in your template. The Jinja2 template can reference:

 - Items within a list
 - Keys and values within a dictionary
 - Attributes of objects

**Example:**

    >>> from jinja2 import Template
    >>> template = Template('Hello {{ name }}!')
    >>> template.render(name='John Doe')
    u'Hello John Doe!'

**Example**
-------

**Basic Usage:**
./xlstoconfig.py examples/cisco_ios.xlsx -t examples/cisco_ios.j2






