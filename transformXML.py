from sys import argv
from lxml import etree
from lxml.builder import E
from lxml.builder import ElementMaker

filename = argv[1]

import os
import shutil
import re
import uuid

import xml.etree.ElementTree as ET

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
    
    # create a new parent element to collect our footnotes
    footnotesholder = etree.Element(w + "footnotes", nsmap=NSMAP)
    
    # adjust each footnote para to add attributes etc.
    for para in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val='footnotetext']") :
      # get the footnote para
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

      # get the original html div id of the previous para
      if footnotesholder.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote") is not None:
        prevdivid = footnotesholder.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote[last()]").find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}divId").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

      # print divid
      # print prevdivid
      
      # compare the div ids of the current para to the prev para
      if divid == prevdivid:
        # if they match, then this is a multi-para note.
        # Add the current para to the previous footnote object we created
        footnotesholder.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote[last()]").append(myparent)
        print "CONTINUE"
      else:
        # if they don't match, then this is the start of a new footnote.
        # Create a new footnote object.
        newfootnote = etree.Element(w + "footnote", nsmap=NSMAP)
        newfootnote.append(myparent)
        footnotesholder.append(newfootnote)
        print "NEW"

      # increment the counter for unique id generation
      footnotecounter += 1

    # Add footnoteref element to first para of each footnote
    for footnote in footnotesholder.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote") :
      # create the footnote reference child
      refrun = etree.Element(w + "r", nsmap=NSMAP)
      refpr = etree.Element(w + "rPr", nsmap=NSMAP)
      refstyle = etree.Element(w + "rStyle", nsmap=NSMAP)
      ref = etree.Element(w + "footnoteRef", nsmap=NSMAP)
      refstyle.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}val'] = 'FootnoteReference'
      refpr.append(refstyle)
      refrun.append(refpr)
      refrun.append(ref)

      # add the footnoteref child to the footnote as the first r element
      firstpara = footnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
      firstpara.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r").addprevious(refrun)

    # Add the extra required IDs to footnotes with multiple paragraphs
    for para in footnotesholder.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p") :
      if para.getnext() is not None:
        if para.getprevious() is not None:
          rsidRPr = para.getprevious().attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr']
        else:
          rsidRPr = generate_rsid(footnotecounter)
        para.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}rsidRPr'] = rsidRPr
        para.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}rsidP'] = rsidRPr

    # insert the first two required blank footnote children
    E = ElementMaker(namespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                     nsmap={'w' : "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    E14 = ElementMaker(namespace="http://schemas.microsoft.com/office/word/2010/wordml",
                     nsmap={'w14' : "http://schemas.microsoft.com/office/word/2010/wordml"})

    # <w:footnote w:type="separator" w:id="-1">
    # <w:p w14:paraId="68D87997" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRDefault="00A51F92">
    # <w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
    # <w:r><w:separator/></w:r></w:p></w:footnote>

    NOTE = E.footnote
    PARA = E.p
    PPR = E.pPr
    SPACING = E.spacing
    RUN = E.r
    SEP = E.separator

    note_sep = NOTE(E.type('separator'), E.id('-1'),
      PARA(E14.paraId('12345678'), E14.textId('77777777'), E.rsidR('12345678'), E.rsidRDefault('12345678'),
        PPR(
          SPACING(E.line("240"), E.lineRule("auto"))
        ),
        RUN(
          SEP()
        )
      )
    )

    print(etree.tostring(note_sep, pretty_print=True))

    newfile = open('footnotes.xml', 'w')
    newfile.write(etree.tostring(footnotesholder, encoding="utf-8", standalone=True, xml_declaration=True))
    newfile.close()

    return

convert_footnotes( filename )
