from sys import argv

filename = argv[1]
finaldir = argv[2]
print filename
print finaldir

import os
import zipfile
import shutil
import re
import xml.etree.ElementTree as ET

def convert_footnotes(self):
    # https://docs.python.org/2/library/xml.etree.elementtree.html
    # read document.xml
    # grab and parse each footnote object
    # create new object and remap to required footnote formatting
    # add to footnote.xml file
    tree = ET.parse(self)
    root = tree.getroot()

    return

convert_footnotes( filename )
