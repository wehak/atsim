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