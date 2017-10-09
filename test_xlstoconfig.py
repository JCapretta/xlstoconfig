#!/usr/bin/python
import unittest
from xlstoconfig import jinjaSkeleton

class Test_jinjaSkeleton(unittest.TestCase):
    def setUp(self):
        pass
    def test_parent_name_type(self):
        self.assertEqual('blah',jinjaSkeleton('blah','blah'))

if __name__ == '__main__':
    unittest.main()
