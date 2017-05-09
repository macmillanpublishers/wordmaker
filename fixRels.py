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

def fix_rels(self):
    # parse the incoming XML
    tree = etree.parse(self)
    root = tree.getroot()

    # get the total number of relationship elements
    rels = root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
    relscount = sum(1 for p in rels)

    # set the namespace for the builder
    E = ElementMaker(namespace="http://schemas.openxmlformats.org/package/2006/relationships")

    # creating the new Relationship object
    REL = E.Relationship

    newrel = REL()

    # Add the required attributes
    newrelid = "rId" + str(relscount + 1)

    newrel.attrib['Id'] = newrelid
    newrel.attrib['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    newrel.attrib['Target'] = "footnotes.xml"

    # insert the new Relationship object
    root.append(newrel)

    # write the new data to a file
    newfile = open('document.xml.rels', 'w')
    newfile.write(etree.tostring(root, encoding="UTF-8", standalone=True, xml_declaration=True))
    newfile.close()

    return

fix_rels( filename )