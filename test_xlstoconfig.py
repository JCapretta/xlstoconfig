#!/usr/bin/python
import unittest
from xlstoconfig import *
import xlstoconfig
import os

argparse_help = '''
usage: test_xlstoconfig.py [-h] [-f FILE] [-t TEMPLATE] [-j]

John's config generator

optional arguments:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  yaml file
  -t TEMPLATE, --template TEMPLATE
                        Jinja2 template file
  -j                    Generate a skeleton for building a jinja2 template
'''


class Test_args(unittest.TestCase):

    def setUp(self):
        pass
    def test_no_args(self):
        #result = get_args()
        with self.assertRaises(SystemExit) as cm:
            get_args([])
        print cm.exception
        self.assertEqual(cm.exception.code,1)
    def test_j_flag(self):
        args=get_args(['file','-j'])
        self.assertEqual(args.create_skeleton,True)
    def test_template_arg(self):
        args=get_args(['file'])
        self.assertEqual(args.file,'file')
        
class Test_Runner(unittest.TestCase):
    def Setup(self):
        pass
    def test_load_yaml(self):
        with self.assertRaises(SystemExit) as cm:
            runner=Runner(['test'])
            runner.load_yaml()
        print cm.exception
        self.assertEqual(cm.exception.code,4)
        
        
       
		
		
		
if __name__ == '__main__':
    unittest.main()
