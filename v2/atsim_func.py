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
                        "Segment": evaluerSegment(balise),
                        # "Segment": bgruppe.sim_segment,                        
                        "Rang": balise.rang, 
                        "X-ord": rens_kodeord(balise.x_reg), 
                        "Y-ord": rens_kodeord(balise.y_reg), 
                        "Z-ord": rens_kodeord(balise.z_reg),
                        "Tegning": bgruppe.s_nr, 
                        "Rad nr.": bgruppe.first_row + 1
                        })

    # Lage workbook-objekt
    workbook  = xlsxwriter.Workbook(wbName)
    worksheet = workbook.add_worksheet("Balisegrupper")

    # Estetikk
    listContent = workbook.add_format({"align": "center"})
    tableHeader = workbook.add_format({"bold": True, "border": True, "align": "center"})

    # Skriv overskrifter
    for i, key in enumerate(baliseDictList[0]):
        worksheet.write(0, i, key, tableHeader)
    
    # Skriv innhold
    for row, baliseDict in enumerate(baliseDictList):
        for col, key in enumerate(baliseDict):
            worksheet.write(row+1, col, baliseDict[key], listContent)
        lastRow = row
    
    # Opprydding
    workbook.close()
    return

# For alle baliser som ikke er A-balise peker segment-celle til A-balisa
def evaluerSegment(bgruppeObj):
    if bgruppeObj.rang == "P":
        return "=INDIRECT(ADDRESS(ROW()+1,COLUMN()))"
    if bgruppeObj.rang != "A":
        return "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))"
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


# Lager en XML-fil av hver inputfil
def createXML(excelFilename):
    import xml.etree.ElementTree as ET
    import xlrd
    from atsim_class import XMLbalise
    
    # Navn på XML-fil som skal genereres
    xmlFilename = excelFilename[:-4] + "xlm"

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