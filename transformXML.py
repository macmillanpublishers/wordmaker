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
    # create new footnotes object
    # add the intial child items
    footnotecounter = 1

    WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w = "{%s}" % WORD_NAMESPACE

    WORD14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"
    w14 = "{%s}" % WORD_NAMESPACE

    NSMAP = {None : WORD_NAMESPACE}

    tree = etree.parse(self)
    root = tree.getroot()
    
    footnotesholder = etree.Element(w + "footnotes", nsmap=NSMAP)
    #print footnotesholder
    
    # loop through each footnote para
    for para in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val='footnotetext']") :
      myparent = para.getparent().getparent()
      # generate the required ids and add them
      textid = generate_textid(footnotecounter)
      rsid = generate_rsid(footnotecounter)
      myparent.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}paraId'] = textid
      myparent.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}textId'] = '77777777'
      myparent.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR'] = rsid
      myparent.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault'] = rsid

      # check to see if it's part of a multi-para note
      divid = para.getparent().find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}divId").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
      prevdivid = ''

      # get the divid of the previous sibling
      prevdiv = myparent.getprevious().find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}divId")

      if footnotesholder.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote") is not None:
        print "YES"
        prevdivid = footnotesholder.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote[last()]").find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}divId").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
        #print prevdivstring
        #prevdivid = prevdiv.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
      else:
        print "Nope"

      print divid
      print prevdivid
      
      if divid == prevdivid:
        footnotesholder.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote[last()]").append(myparent)

        print "CONTINUE"
      else:
        newfootnote = etree.Element(w + "footnote", nsmap=NSMAP)
        newfootnote.append(myparent)
        footnotesholder.append(newfootnote)

        print "NEW"

      # parentstring = etree.tostring(myparent)
      # print parentstring

      footnotestring = etree.tostring(footnotesholder)
      print footnotestring

      # if div id matches prev, add to prev footnote object
      # else wrap in a new footnote parent
      # add new parent to larger footnote block

      footnotecounter += 1

    return

convert_footnotes( filename )
