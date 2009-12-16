#!/usr/bin/python3.1
# anti-docx, a .docx-to-text converter written by Albin Stjerna
#
# Please, feel free to do what you want with this code. It is way too short
# for a proper license. :)
import zipfile, sys, textwrap
from xml.sax import parse, ContentHandler
from optparse import OptionParser

def extract_document(fn):
    """return a file pointer to the XML file in the .docx 
    containing the actual document"""
    wf=zipfile.ZipFile(fn, "r")
    docfile=wf.open("word/document.xml", "r")
    wf.close()
    return docfile

def get_word_paragraphs(xmlfile):
    """Return a list of paragraphs from
       the docx-formatted xml file at xmlfile"""
    class tagHandler(ContentHandler):
        def __init__(self):
            self.paragraphMarker = "w:p"
            self.textMarker = "w:t"
            self.paragraphs = []
            self.string = ""
            self.inText = False
            self.inParagraph = False
        def startElement(self, name, attr):
            if name == self.textMarker:
                self.inText = True
            elif name == self.paragraphMarker:
                self.inParagraph = True
        def endElement(self, name):
            if name == self.textMarker:
                self.inText = False
            elif name == self.paragraphMarker:
                self.inParagraph == False
                self.paragraphs.append(self.string)
                self.string = ""
        def characters(self, ch):
            if self.inText:
                self.string+=ch

    handler = tagHandler()
    parse(xmlfile, handler)
    return handler.paragraphs


def main():
    parser = OptionParser()

    parser.add_option("-w", "--wrap-lines",
                      action="store_true", dest="wraplines", default=False,
                      help="wrap long lines in output")

    (options, args) = parser.parse_args()

    if len(args) != 1:
        parser.error("incorrect number of arguments (did you specify a file name?)")

    filename = args[0]

    wrapper = textwrap.TextWrapper()

    for a in get_word_paragraphs(extract_document(filename)):
        if options.wraplines:
            print(wrapper.fill(a))
        else:
            print(a)

if __name__ == "__main__":
    main()







