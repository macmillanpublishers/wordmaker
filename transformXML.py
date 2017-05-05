from sys import argv
from lxml import etree

filename = argv[1]

import os
import shutil
import re
import uuid

ns = {'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def generate_textid(counter):
    idbase = uuid.uuid4().hex
    idshort = idbase[:8]
    iduniq = idshort + str(counter)
    iduniq = iduniq[-8:]

    return str(iduniq)

def generate_rsid(counter):
    idbase = uuid.uuid4().hex
    idshort = idbase[:8]
    iduniq = idshort + str(counter)
    iduniq = "00" + iduniq[-6:]

    return str(iduniq)

def convert_footnotes(self):
    # https://docs.python.org/2/library/xml.etree.elementtree.html
    # read document.xml
    # grab and parse each footnote object
    # create new object and remap to required footnote formatting
    # add to footnote.xml file
    footnotecounter = 1

    tree = etree.parse(self)
    root = tree.getroot()
    print(root.tag)
    for para in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val='footnotetext']") :
      myparent = para.getparent().getparent()
      textid = generate_textid(footnotecounter)
      rsid = generate_rsid(footnotecounter)
      myparent.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}paraId'] = textid
      myparent.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}textId'] = '77777777'
      myparent.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR'] = rsid
      myparent.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault'] = rsid
      parentstring = etree.tostring(myparent)
      footnotecounter += 1
      print parentstring

    return

convert_footnotes( filename )
