**xlstoconfig**
-----------

**Overview** 
xlstoconfig takes data from an excel file and presents it for use with a Jinja2 template. 

**Components:**

 - Xlstoconfig.py: python script
 - Jinja2 template
 - Configuration file (yaml file)
   - Pointer to data within an excel file
   - Defines the attributes and structure of an object that is then passed to Jinja2.
 - Excel file: this holds data the you want to use in your Jinja2 template

Xlstoconfig creates an object called *config* and passes this to Jinja2. The attributes of this object are defined/named in the yaml configuration file. The attributes can hold scalar or list objects. List objects are most useful. The data that's read from an Excel file will always be represented as a list object where each element of the list corresponds to a row within a table. 
Example use cases:

 - Generating Cisco config
 - Creating a xml file for Forcepoint SMC
 - Generate text for change records for cookie cutter type changes

This can also be imported as a module in other Python code. For example: you're writing a script that configures a device using REST API. All the configuration parameters can be read in from the Excel file. In this case you wouldn't use a Jinja2 template. 


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
Let's walk through the cisco_ios example that can be found in the *examples* directory.

**cisco_ios.xlsx**
The is a holding data about vlans and switchport parameters. We want to generate some Cisco IOS configuration based on this data.
In general, there isn't any need to structure the workbook to cater for the script.

**cisco_ios.yml**
This mapping defines the path to the Excel workbook. It is possible to reference multiple workbooks:

    workbook: examples/cisco_ios.xlsx

The *values* node allows the user to define scalar variables to be used by the Jinja2 template. The Jinja2 template is always passed a single object called *config*. Any mappings defined under *values* becomes an attribute of the *config* object.

    values:
        hostname: ACMEME1EXTSW01

Here's how to use this mapping in the template:

    {{config.hostname}}

The *children* node allows the user to define lists:

    children:
        vlans:
            worksheet: VLANs
            range: A2:D28
            columns: #the first column is 0
                name: 1
                id: 0
        interfaces:    
            worksheet: VLANs
            range: B36:F91
            autocolumn: 1

In the above example, two lists will be created: one named *vlans* and one named *interfaces*. In the spreadsheet there are two tables. Each of the lists correspond to one of these tables. Each element of the list corresponds to a row in one of these tables.
Each element of the lists is an object of class Config. You can think of each of these elements as children of the parent *config* node. 
It possible for the child nodes under the parent to have their own child nodes, but this usually isn't necessary.

This mapping defines the spreadsheet where the data resides:

    worksheet: VLANs
    
This mapping is the range within the spreadsheet where the data resides:

    range: B36:F91

Don't include the header row in this range definition.
If the spreadsheet has a single table then it's possible to omit the *range* mapping and let the range be discovered automatically. This feature hasn't been tested extensively and it's recommended to explicitly set the range.


        columns: #the first column is 0
            name: 1
            id: 0

Recall that rows in a table are stored as elements within a list and that these elements are objects of class Config. Each of these objects have attributes corresponding to a column in the table. The mappings under the *columns* node define the names of these attributes and the column number (indexing starts at zero).

It is mandatory for at least one column to be mapped to *name*. In the above example you might not want to use *name* for the 2nd column. In this case you create another mapping to the second column and keep the *name* mapping.


        columns: #the first column is 0
            name: 1
			description: 1
            id: 0


Here's to use these mapping in a template:

    {% for vlan in config.vlans %}
    vlan {{ vlan.id }}
        name "{{ vlan.name }}"
    {% endfor %}


The above Jinja2 template iterates through a list (*config.vlans*). In each iteration of the *for* loop, the template is using the *id* and *name* attributes/columns of the table row.

    interfaces:    
        worksheet: VLANs
        range: B36:F91
        autocolumn: 1

For the *interfaces* node, the row attributes/columns are not explicitly defined. Instead *autocolumn: 1* is used. When autocolumn is used, the header row is used for the row attributes/column names.

**Basic Usage:**
./xlstoconfig.py examples/cisco_ios.yml -t examples/cisco_ios.j2






