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

    # namespace declarations for the element method we'll use later
    WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    w = "{%s}" % WORD_NAMESPACE

    WORD14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"
    w14 = "{%s}" % WORD_NAMESPACE

    NSMAP = {None : WORD_NAMESPACE}

    # parse the incoming XML
    tree = etree.parse(self)
    root = tree.getroot()
    
    # create a new parent element to collect our footnotes
    # and insert the first two required blank footnote children
    # we'll use the builder method to create this.
    # this is the namespace declaration for the builder.
    E = ElementMaker(namespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                     nsmap={'w' : "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                            'w14' : "http://schemas.microsoft.com/office/word/2010/wordml",
                            'wpc' : "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                            'mc' : "http://schemas.openxmlformats.org/markup-compatibility/2006",
                            'o' : "urn:schemas-microsoft-com:office:office",
                            'r' : "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                            'm' : "http://schemas.openxmlformats.org/officeDocument/2006/math",
                            'v' : "urn:schemas-microsoft-com:vml",
                            'wp14' : "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
                            'wp' : "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                            'w10' : "urn:schemas-microsoft-com:office:word",
                            'wpg' : "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                            'wpi' : "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
                            'wne' : "http://schemas.microsoft.com/office/word/2006/wordml",
                            'wps' : "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"})

    # creating the footnotes object
    NOTES = E.footnotes

    footnotes = NOTES()

    footnotes.attrib['{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable'] = 'w14 wp14'
   
    # adjust each footnote para to add attributes etc.
    for para in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val='footnotetext']") :
      # get the footnote para
      myparent = para.getparent().getparent()
      # generate the required ids and add them
      paraid = generate_textid(footnotecounter)
      rsid = generate_rsid(footnotecounter)
      myparent.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}paraId'] = paraid
      myparent.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}textId'] = '77777777'
      myparent.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR'] = rsid
      myparent.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault'] = rsid

      # check to see if it's part of a multi-para note
      divid = para.getparent().find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}divId").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
      prevdivid = ''

      # get the original html div id of the previous para
      if footnotes.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote") is not None:
        prevdivid = footnotes.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote[last()]").find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}divId").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
      
      # compare the div ids of the current para to the prev para
      if divid == prevdivid:
        # if they match, then this is a multi-para note.
        # Add the current para to the previous footnote object we created
        footnotes.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote[last()]").append(myparent)
      else:
        # if they don't match, then this is the start of a new footnote.
        # Create a new footnote object via the Element method
        newfootnote = etree.Element(w + "footnote", nsmap=NSMAP)
        newfootnote.append(myparent)
        footnotes.append(newfootnote)

      # increment the counter for unique id generation
      footnotecounter += 1

    # Add footnoteref element to first para of each footnote
    for footnote in footnotes.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote") :
      # create the footnote reference child via the Element method
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
    for para in footnotes.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p") :
      if para.getnext() is not None:
        if para.getprevious() is not None:
          rsidRPr = para.getprevious().attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr']
        else:
          rsidRPr = generate_rsid(footnotecounter)
        para.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}rsidRPr'] = rsidRPr
        para.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}rsidP'] = rsidRPr

    # Add the required first footnote children to the main footnotes object
    NOTE = E.footnote
    PARA = E.p
    PPR = E.pPr
    SPACING = E.spacing
    RUN = E.r
    SEP = E.separator
    CONSEP = E.continuationSeparator

    firstnote = NOTE(
      PARA(
        PPR(
          SPACING()
        ),
        RUN(
          SEP()
        )
      )
    )
    
    secondnote = NOTE(
      PARA(
        PPR(
          SPACING()
        ),
        RUN(
          CONSEP()
        )
      )
    )

    paraid = generate_textid(footnotecounter)
    rsid = generate_rsid(footnotecounter)
    firstnote.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type'] = 'separator'
    firstnote.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id'] = '-1'
    firstnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.microsoft.com/office/word/2010/wordml}paraId'] = paraid
    firstnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.microsoft.com/office/word/2010/wordml}textId'] = '77777777'
    firstnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR'] = rsid
    firstnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault'] = rsid
    firstnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line'] = '240'
    firstnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule'] = 'auto'

    paraid = generate_textid(footnotecounter)
    rsid = generate_rsid(footnotecounter)
    secondnote.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type'] = 'continuationSeparator'
    secondnote.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id'] = '0'
    secondnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.microsoft.com/office/word/2010/wordml}paraId'] = paraid
    secondnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.microsoft.com/office/word/2010/wordml}textId'] = '77777777'
    secondnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR'] = rsid
    secondnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRDefault'] = rsid
    secondnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line'] = '240'
    secondnote.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing").attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule'] = 'auto'

    footnotes.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote").addprevious(secondnote)
    footnotes.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote").addprevious(firstnote)

    # Now print it all to a file!
    newfile = open('footnotes.xml', 'w')
    newfile.write(etree.tostring(footnotes, encoding="utf-8", standalone=True, xml_declaration=True))
    newfile.close()

    return

convert_footnotes( filename )
