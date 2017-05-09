from sys import argv
# make sure to insall lxml: sudo pip install lxml
from lxml import etree
from lxml.builder import E
from lxml.builder import ElementMaker

filename = argv[1]

import os
import shutil
import re
import uuid

import xml.etree.ElementTree as ET

def fix_types(self):
    # parse the incoming XML
    tree = etree.parse(self)
    root = tree.getroot()

    # set the namespace for the builder
    E = ElementMaker(namespace="http://schemas.openxmlformats.org/package/2006/content-types")

    # creating the new Relationship object
    OVER = E.Override

    newoverride = OVER()

    newoverride.attrib['PartName'] = '/word/footnotes.xml'
    newoverride.attrib['ContentType'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'

    # insert the new Override object
    root.append(newoverride)

    # write the new data to a file
    newfile = open('[Content_Types].xml', 'w')
    newfile.write(etree.tostring(root, encoding="UTF-8", standalone=True, xml_declaration=True))
    newfile.close()

    return

fix_types( filename )