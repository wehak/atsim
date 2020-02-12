# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 14:10:14 2019

@author: WEYHAK

Diverse funksjoner til atsim

"""

# tar en bokstavkode, gir plass i alfabetet
def alphabet_number(some_char):
    return ord(some_char.upper())-64

def col_name(letter_string):
    sum = 0
    for idx, c in enumerate(reversed(letter_string)):
        sum += 26**idx*alphabet_number(c)
    return sum - 1

# lager et regneark som lar bruker definere segmenter
def definer_segmenter(baliseoversikt_obj):
    
    import pandas as pd
    
    pd_import = []
    for ktab in baliseoversikt_obj.alle_ktab:
        for balise in ktab.balise_group_obj_list:
            pd_import.append({
                "Sign./type" : balise.sign_type,
                "Sted" : balise.id1,
                "ID" : balise.id2,
                "KM" : balise.km,
                "Retning" : balise.retning,
                "Tegning" : balise.s_nr,
                "Rad nr." : "{}" .format(balise.first_row+1),
                "Segment" : ""
            })
    
    balise_df = pd.DataFrame(pd_import)
    balise_df = balise_df[['Retning', 'Sign./type', 'Sted', 'ID', 'KM', 'Tegning', 'Rad nr.', 'Segment']]
    return(balise_df)

# vasker kodeord for å presenteres i excel    
def rens_kodeord(kodeliste):
    
    kodeliste = set(kodeliste)
    
    if "-" in kodeliste:
        kodeliste.remove("-")
        if len(kodeliste) == 0:
            return 1
    
    if len(kodeliste) == 1:
        return kodeliste.pop()
    else: 
        return ', '.join(map(str, kodeliste))


def lagReferanse(row, col):
    return str(chr(65+col)) + str(row)



# Lager excelark med baliser
def skrivBaliseliste(ktabList, wbName):
    import xlsxwriter

    # Lager liste med dictionaries
    baliseDictList = []
    for ktab in ktabList.alle_ktab:
        for bgruppe in ktab.balise_group_obj_list:
            for balise in bgruppe.baliser:                
                baliseDictList.append({
                        "Retning": bgruppe.retning,
                        "Sign./Type": bgruppe.sign_type,
                        "Type": bgruppe.type,
                        "ID_sted": bgruppe.id1, 
                        "ID_type": bgruppe.id2, 
                        "KM_prosjektert": balise.km,
                        "KM_simulering": 0,
                        # "KM_simulering": "=" + lagReferanse(len(baliseDictList)+2, 6-1), # rad og kolonne det skal refereres til
                        # "Segment": evaluerSegment(balise, len(baliseDictList)+2, 7), # for å gjøre P, B, C til referanse
                        "Segment": bgruppe.sim_segment,                        
                        "Rang": balise.rang, 
                        "X-ord": rens_kodeord(balise.x_reg), 
                        "Y-ord": rens_kodeord(balise.y_reg), 
                        "Z-ord": rens_kodeord(balise.z_reg),
                        "Tegning": bgruppe.s_nr, 
                        "Rad nr.": bgruppe.first_row + 1
                        })

    # Lage workbook-objekt
    workbook  = xlsxwriter.Workbook(wbName)
    baliseWorksheet = workbook.add_worksheet("Balisegrupper")

    # Estetikk
    listContent = workbook.add_format({"align": "center"})
    tableHeader = workbook.add_format({"bold": True, "border": True, "align": "center"})

    # Definer tabell
    data = makeListOfLists(baliseDictList)
    baliseWorksheet.add_table(0,0, len(data), len(data[0])-1, {
        "data": data,
        "columns": makeHeaders(baliseDictList)
        # "header_row": True
        })

    # # Skriv skilter
    # skiltCols = [
    #     "Retning",
    #     "Sign./Type",
    #     "Type",
    #     "ID_sted",
    #     "ID_type",
    #     "KM_prosjektert" # må være siste kolonne
    #     ]

    # # Kolonneoverskrifter
    # skiltWorksheet = workbook.add_worksheet("Skilt")
    # for col, content in enumerate(skiltCols):
    #     skiltWorksheet.write(0, col, content, tableHeader)
    
    # # Tabelldata
    # skiltDataTbl = []
    # for i, baliseDict in enumerate(baliseDictList):
    #     row = []
    #     if baliseDict["Rang"] == "A":
    #         for j, key in enumerate(skiltCols):
    #             if (key == "KM_prosjektert"):
    #                 row.append("=Balisegrupper!" + lagReferanse(i+2, len(skiltCols)))
    #             else: 
    #                 row.append(baliseDict[key])
    #         skiltDataTbl.append(row)
    
    # skiltWorksheet.add_table(0,0, len(skiltDataTbl), len(skiltDataTbl[0])-1, {
    #     "data": skiltDataTbl,
    #     "columns": makeHeadersList(skiltCols)
    #     # "header_row": True
    #     })
    
    # Opprydding
    workbook.close()
    return

"""
                tally = 0
    skiltDataTbl = []
    for row, baliseDict in enumerate(baliseDictList):
        if baliseDict["Rang"] == "A":
            for i, col in enumerate(skiltCols):
                if (col == "KM_prosjektert"):
                    skiltWorksheet.write(
                        tally+1,
                        i,
                        "=Balisegrupper!" + lagReferanse(row+2, len(skiltCols)),
                        listContent
                        )
                else: 
                    skiltWorksheet.write(tally+1, i, baliseDict[col], listContent)
            tally += 1
"""
    


# For alle baliser som ikke er A-balise peker segment-celle til A-balisa
def evaluerSegment(bgruppeObj, row, col):
    if bgruppeObj.rang == "P":
        return "=" + str(chr(65+col)) + str(row+1)
    if bgruppeObj.rang != "A":
        return "=" + str(chr(65+col)) + str(row-1)
    else:
        return "?"

# Finner alle .xlsx filer i angitt mappe
def getXLSXfileList(folder_path):
    from os import walk
    fileList = []
    (_, _, filenames) = next(walk(folder_path))

    for file in filenames:
        if file.lower().endswith(".xlsx"):
            fileList.append(folder_path + "\\" + file)
    print("Antall .XLSX-filer funnet: {}" .format(len(fileList)))
    return fileList

""" Utgår? Bruker annet bibliotek
# Lager en XML-fil av hver inputfil
def createXML(excelFilename):
    import xml.etree.ElementTree as ET
    import xlrd
    from atsim_class import XMLbalise
    
    # Navn på XML-fil som skal genereres
    xmlFilename = excelFilename[:-4] + "xml"

    # Sti til excelark med baliser
    # excelFilename = "lysaker1.xlsx"

    # Navn på kolonner som skal leses inn fra excelark
    searchPatterns = [
        "ID_sted",
        "ID_type",
        "KM_simulering",
        "Rang",
        "X-ord",
        "Y-ord",
        "Z-ord"
        ]


    # Åpnder excelark
    wb = xlrd.open_workbook(excelFilename) # åpner excel workbook
    ws = wb.sheet_by_index(0) # aktiverer sheet nr 0

    # Finner colonne med relevant data
    antallPatterns = len(searchPatterns)
    headerColumnDict = {}
    for i, header in enumerate(ws.row_values(0)):
        for j, pattern in enumerate(searchPatterns):
            if header == pattern:
                headerColumnDict.update({header:i})
                searchPatterns.pop(j)
                
    # Sjekker om alt er funnet
    if len(headerColumnDict) == antallPatterns:
        print("OK")
    else:
        print("Mangler verdier")

    baliser = []
    for i in range(1, ws.nrows):
        baliser.append(
            XMLbalise(
                ws.cell_value(i, headerColumnDict["ID_sted"]),
                ws.cell_value(i, headerColumnDict["ID_type"]),
                ws.cell_value(i, headerColumnDict["Rang"]),
                ws.cell_value(i, headerColumnDict["KM_simulering"]),
                ws.cell_value(i, headerColumnDict["X-ord"]),
                ws.cell_value(i, headerColumnDict["Y-ord"]),
                ws.cell_value(i, headerColumnDict["Z-ord"])
            )
        )


    # Første tag
    root = ET.Element("TrackConnectedObjectListXML")

    # Start KM
    startKM = ET.SubElement(root, "KmInfoXML")
    ET.SubElement(startKM, "KmOffsetXML").text = "0"

    for balise in baliser:
        balise.toXML(root)

    tree = ET.ElementTree(root)
    tree.write(xmlFilename, encoding="UTF-8", xml_declaration=True, default_namespace=None, method="xml")
"""    



def makeHeaders(DictList):
    return [{"header": "{}" .format(key)} for key in DictList[0]]


def makeHeadersList(inputList):
    return [{"header": "{}" .format(item)} for item in inputList]

def makeListOfLists(DictList):
    return [list(dictionary.values()) for dictionary in DictList]

def createXML(excelFilename):
    import xml.etree.ElementTree as etree
    import xlrd
    from atsim_class import XMLbalise
    
    # Navn på XML-fil som skal genereres
    xmlFilename = excelFilename[:-4] + "xml"

    # Sti til excelark med baliser
    # excelFilename = "lysaker1.xlsx"

    # Navn på kolonner som skal leses inn fra excelark
    searchPatterns = [
        "Sign./Type",
        "ID_sted",
        "ID_type",
        "KM_simulering",
        "Rang",
        "X-ord",
        "Y-ord",
        "Z-ord"
        ]


    # Åpnder excelark
    wb = xlrd.open_workbook(excelFilename) # åpner excel workbook
    ws = wb.sheet_by_index(0) # aktiverer sheet nr 0

    # Finner colonne med relevant data
    antallPatterns = len(searchPatterns)
    headerColumnDict = {}
    for i, header in enumerate(ws.row_values(0)):
        for j, pattern in enumerate(searchPatterns):
            if header == pattern:
                headerColumnDict.update({header:i})
                searchPatterns.pop(j)
                
    # Sjekker om alt er funnet
    if len(headerColumnDict) == antallPatterns:
        print("OK")
    else:
        print("Mangler verdier")

    baliser = []
    for i in range(1, ws.nrows):
        baliser.append(
            XMLbalise(
                ws.cell_value(i, headerColumnDict["Sign./Type"]),
                ws.cell_value(i, headerColumnDict["ID_sted"]),
                ws.cell_value(i, headerColumnDict["ID_type"]),
                ws.cell_value(i, headerColumnDict["Rang"]),
                ws.cell_value(i, headerColumnDict["KM_simulering"]),
                ws.cell_value(i, headerColumnDict["X-ord"]),
                ws.cell_value(i, headerColumnDict["Y-ord"]),
                ws.cell_value(i, headerColumnDict["Z-ord"])
            )
        )


    # Første tag
    root = etree.Element("TCO-balises")
    root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    TCOlist = etree.SubElement(root, "TrackConnectedObjectListXML")

    # Start KM
    startKM = etree.SubElement(TCOlist, "KmInfoXML")
    etree.SubElement(startKM, "KmOffsetXML").text = "0"

    for balise in baliser:
        balise.toXML(TCOlist)

    tree = etree.ElementTree(root)
    tree.write(xmlFilename, encoding="UTF-8", xml_declaration=True, default_namespace=None, method="xml")