from sys import argv

script, filename = argv

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

workingpath = os.path.abspath('Desktop/html-to-doc-tests/newzip/')
print workingpath

def convert_manuscript(self):

    # must be .docx or .docm
    extension = os.path.splitext(self)[1]

    if extension in ('.docx', '.docm', '.doc'):
        print "converting %s" % filename
        # get the contents of the Word file
        document = zipfile.ZipFile(self, 'a')
        document.extractall(workingpath)
        document.close()

        # read the contents of the main document.xml file
        xmlfile = open('Desktop/html-to-doc-tests/newzip/word/document.xml', 'r')
        xml_content = xmlfile.read()

        # replace the stylenames in document.xml with the correctly capitalized names
        for k,v in Styles.items():
            print str(k)
            print str(v)
            xml_content = re.sub(str(k), str(v), xml_content)

        # write the re-styled content to a new document.xml file
        newfile = open('Desktop/html-to-doc-tests/document.xml', 'w')
        with newfile as f:
            f.write(xml_content)
            f.close()

        # copy the new document.xml file into the extracted zip, replacing the old version
        shutil.copy('Desktop/html-to-doc-tests/document.xml', 'Desktop/html-to-doc-tests/newzip/word/')
        # copy the style definitions into the extracted zip, replacing the old version
        shutil.copy('Desktop/html-to-doc-tests/styles.xml', 'Desktop/html-to-doc-tests/newzip/word/')

        # zip all the contents of the extracted zip,
        # WITHOUT nesting them inside a root folder.
        # def zipdir(path, ziph):
        #     # ziph is zipfile handle
        #     for root, dirs, files in os.walk(path):
        #         for file in files:
        #             ziph.write(file)

        # if __name__ == '__main__':
        #     zipf = zipfile.ZipFile('Python.zip', 'w', zipfile.ZIP_DEFLATED)
        #     zipdir('Desktop/html-to-doc-tests/newzip/', zipf)
        #     zipf.close()

        return

def convert_footnotes(self):
    # https://docs.python.org/2/library/xml.etree.elementtree.html
    # read document.xml
    # grab and parse each footnote object
    # create new object and remap to required footnote formatting
    # add to footnote.xml file

    return

convert_manuscript( filename )
