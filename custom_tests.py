import re

"""
This module contains custom jinja2 filters to be used with make_config.py
"""


def search(string,pattern):
    return re.search(pattern, string)

