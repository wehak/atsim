from atsim_func import getXLSXfileList, createXML

"""
Finner alle xlsx-filer i stien "folderPath" og forsøker å gjøre de om til xml-filer.
Scriptet vil bare forsøke å lese blad nr "0" i arbeidsboken, dvs første bladet.
"""

folderPath = r"C:\Users\weyhak\OneDrive - Bane NOR\Dokumenter\DIV\lysaker sim\scenarioer"

for filename in getXLSXfileList(folderPath):
    createXML(filename)