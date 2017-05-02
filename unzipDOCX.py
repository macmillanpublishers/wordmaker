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

def unzip_manuscript(self):

    # must be .docx or .docm
    extension = os.path.splitext(self)[1]

    if extension in ('.docx', '.docm', '.doc'):
        print "unzipping %s" % filename
        # get the contents of the Word file
        document = zipfile.ZipFile(self, 'a')
        document.extractall(finaldir)
        document.close()

        return

unzip_manuscript( filename )
