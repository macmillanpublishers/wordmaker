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

# To do: zip the edited contents of the extracted archive (NOT nested in a root file),
# and rename the ext to .docx
# Then relink the footnotes/endnotes

Styles = {}
Styles['copyrighttextdoublespacecrtxd'] = "CopyrightTextdoublespacecrtxd"
Styles['titlepagebooktitletit'] = "TitlepageBookTitletit"

worddir = os.path.join(finaldir, "word")
newxml = os.path.join(finaldir, "document.xml")

def rename_styles(self):
    # read the contents of the main document.xml file
    xmlfile = open(filename, 'r')
    xml_content = xmlfile.read()

    # replace the stylenames in document.xml with the correctly capitalized names
    for k,v in Styles.items():
        print str(k)
        print str(v)
        xml_content = re.sub(str(k), str(v), xml_content)

    # write the re-styled content to a new document.xml file
    newfile = open(newxml, 'w')
    with newfile as f:
        f.write(xml_content)
        f.close()

    # copy the new document.xml file into the extracted zip, replacing the old version
    shutil.copy(newxml, worddir)
    # copy the style definitions into the extracted zip, replacing the old version
    #shutil.copy('Desktop/html-to-doc-tests/styles.xml', 'Desktop/html-to-doc-tests/newzip/word/')

    return

rename_styles( filename )
