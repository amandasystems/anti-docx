#!/usr/bin/python3.1
# anti-docx, a .docx-to-text converter written by Albin Stjerna
#
# Please, feel free to do what you want with this code. It is way too short
# for a proper license. :)
import zipfile, sys
from xml.sax import parse, ContentHandler

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

wordfile= sys.argv[1]

for a in get_word_paragraphs(extract_document(wordfile)):
    print(a)



