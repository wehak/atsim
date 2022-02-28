from xmlHelpers import getXLSXfileList, createXML

"""
Finner alle xlsx-filer i stien "folderPath" og forsøker å gjøre de om til xml-filer.
Scriptet vil bare lese blad nr "0" i arbeidsboken, dvs første bladet.
"""

folderPath = r"C:\Users\weyhak\Desktop\grorud sim 22\scenario"

for filename in getXLSXfileList(folderPath):
    createXML(filename)