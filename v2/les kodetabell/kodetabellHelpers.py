# -*- coding: utf-8 -*-
"""
Created on Thu Jul  4 13:55:16 2019

@author: Håkon Weydahl (weyhak@banenor.no)

Inneholder klasser som kan ta en mengde kodetabeller og hente ut informasjonen. 
Bibliotektet som snakker med excel (xlrd) er testet kun på .xls-filer. 
Dersom kodetabellen er i det nyere .xlsx-formatet kan kodetabellen lagres på
nytt i gammelt format.

Objekter:
    -   Baliseoversikt(): "Permen" med alle kodetabellene du er interessert i. 
        Innholder en liste over alle kodetabellene
    -   Kodetabell(): Hvert enkelt regneark, inneholder en liste over alle 
        balisegruppene på arket
    -   Balisegruppe(): Den enkelte bgruppe, inneholder en liste over alle 
        balisene i gruppa
    -   Balise(): En enkelt balise
    
"""

from pathlib import Path
import re
import xlrd

###########
# Klasser #
###########

class Baliseoversikt:
    def __init__(self):
        self.alle_ktab = []
        
    def ny_mappe(self, folder_path):
        for file in self.__getXLSfileList(folder_path):
            self.alle_ktab.append(Kodetabell(file))
            
    # hvordan oversikten printes
    def __str__(self):
        balisegrupper_df = PD_table(self.alle_ktab)
        print(balisegrupper_df.balise_df)
        return ""

    # Finner alle .xls filer i angitt mappe
    def __getXLSfileList(self, folder_path):
        # import os
        # xls_files = []
        # (_, _, filenames) = next(os.walk(folder_path))
    
        # for file in filenames:
        #     if file.lower().endswith(".xls"):
        #         xls_files.append(folder_path + "\\" + file)

        xls_files = list(Path(folder_path).rglob("*.xls"))

        # sjekker om filer er funnet, slutter hvis ikke
        if len(xls_files) == 0:
            print(f"Ingen .XLS-filer funnet i '{folder_path}'")
            exit()
        else:
            print("Antall .XLS-filer funnet: {}" .format(len(xls_files)))
            return xls_files
    

class Kodetabell:
    def __init__(self, filepath):
        self.filepath = filepath
        self.balise_group_obj_list = [] # liste over alle balisegrupper på arket
        
        # Hvilke kolonner i excel-arket som definerer en tilstand
        # <navn> : <excel-kolonne>
        self.ktab_cols = {
                "H" : "F",
                "F/H" : "G",
                "F" : "H",
                "kjor" : "I",
                "vent" : "J",
                "p-avstand" : "K",
                "b-avstand" : "L",
                "fall" : "M",
                "PX" : "AP", "PY" : "AQ", "PZ" : "AR", # p-balise
                "AX" : "AS", "AY" : "AV", "AZ" : "AX", # a-balise
                "BX" : "AZ", "BY" : "BA", "BZ" : "BC", # b-balise
                "CX" : "BE", "CY" : "BF", "CZ" : "BG", # c-balise
                "NX" : "BH", "NY" : "BI", "NZ" : "BJ", # n-balise
                "motr_type" : "BP",
                "motr_hast" : "BQ"
                }
             
        # initiering starter her
        self.__les_kodetabell()
    
    def __les_kodetabell(self):
        print(self.filepath)
        self.wbook = xlrd.open_workbook(self.filepath) # åpner excel workbook
        self.wb_sheet = self.wbook.sheet_by_index(0) # aktiverer sheet nr 0
        
        self.__definer_balisegrupper() # lager balise_group_obj_list
        
        for bgruppe in self.balise_group_obj_list:
            bgruppe = self.__definer_tilstander(bgruppe)
            bgruppe.kodere = self.__tell_kodere(bgruppe)
            # print(bgruppe.id2, "\n", bgruppe.tilstander) # kun for debugging. printer output
    
    
    # søker etter balisegrupper i kodetabellen
    def __definer_balisegrupper(self):
        for group_row in range(5,42):            
            # Lager balise objekt med __init__ info 
            if (self.wb_sheet.cell(group_row,1).ctype==0) or \
            (self.wb_sheet.cell(group_row,2).ctype==0 and
             self.wb_sheet.cell(group_row,3).ctype==0):# and
            #  self.wb_sheet.cell(group_row,4).ctype==0):
                continue
            else:
                self.balise_group_obj_list.append(Balisegruppe(
                    self.wb_sheet.cell_value(group_row,1), # sign_type
                    self.wb_sheet.cell_value(group_row,2), # id1
                    self.wb_sheet.cell_value(group_row,3), # id2
                    self.__clean_KM(self.wb_sheet.cell_value(group_row,4)), # km
                    self.wb_sheet.cell_value(5,0), # ktab retning
                    self.wb_sheet.cell_value(50,90), # s_nr
                    group_row, # første rad nr
                    self.__last_row(group_row) # siste rad nr
                ))

    # finner alle definerte tilstander for en balisegruppe
    def __definer_tilstander(self, group_obj):        
        # search_col, returnerer en liste for hver kolonne        
        kolonne_dict = {}
        for key in self.ktab_cols:
            value = self.__search_col(
                            col_name(self.ktab_cols[key]),
                            group_obj
                            )
            if value != None:
                kolonne_dict.update({key : value})
                
        # lager en linje per tilstand
        tilstand_list = []
        row_span = group_obj.last_row - group_obj.first_row + 1
        for row in range(row_span):            
            tilstand_linje = []
            for key in kolonne_dict:
                # print(row, key, kolonne_dict[key][row]) # debugging
                tilstand_linje.append({key : kolonne_dict[key][row]})
            togvei_celle = self.wb_sheet.cell_value(
                    group_obj.first_row + row,
                    col_name("CB")
                    )
            
            # kopier over eventuelt innhold fra celle med togvei info
            if togvei_celle != "":                
                tilstand_linje.append({"togvei" : self.wb_sheet.cell_value(
                        group_obj.first_row + row,
                        col_name("CB")
                        )})
            
            tilstand_list.append(tilstand_linje)
        group_obj.tilstander = tilstand_list
        
        # lager Balise objekt med info om koding
        for litra in ["P", "A", "B", "C"]:
            if litra + "X" in kolonne_dict:
                group_obj.baliser.append(Balise(
                        litra,
                        kolonne_dict[litra + "X"],
                        kolonne_dict[litra + "Y"],
                        kolonne_dict[litra + "Z"]
                        ))
                
        # sette km på balisene        
        if ("A" in group_obj.retning):
            retning = -1
        else:
            retning = 1
            
        offset = 6 # hvor mange meter fra hsign til første balise
        
        for balise in group_obj.baliser:
            if group_obj.type == "H.sign":
                egen_gruppe = [balise.rang for balise in group_obj.baliser]
                for i, bokstav in enumerate(egen_gruppe[::-1]):
                    if bokstav == balise.rang:
                        balise.km = group_obj.km + (offset + 3 * i) * retning
            else:
                egen_gruppe = [balise.rang for balise in group_obj.baliser]
                for i, bokstav in enumerate(egen_gruppe[::-1]):
                    if bokstav == balise.rang:
                        balise.km = group_obj.km + 3 * i * retning                    
                # if balise.rang is "P":
                #     balise.km = group_obj.km - 3 * retning
                # else:
                #     for i, bokstav in enumerate(["A", "B", "C"]):
                #         if bokstav == balise.rang:
                #             balise.km = group_obj.km + 3 * i * retning
        # def slutt
        return group_obj
    
    # leter i kommentarfeltet etter gyldige koderbenevninger, returnerer liste
    def __tell_kodere(self, group_obj):
        
        # gyldige navn på kodere:
        koder_benevning = (
        "FSK[1-9]*"
        "|HSK[1-9]*"
        "|DSK[1-9]*"
        "|VK[ZY1-9]*"
        "|PK[ZY1-9]*"
        "|BK[ZY1-9]*"
        "|CK[ZY1-9]*"
        "|REP\.*K[1-9]*"
        "|RSK[1-9]*"
        )
        
        koder_list = []        
        for row in range(group_obj.first_row, group_obj.last_row + 1):
            kommentar_celle = self.wb_sheet.cell_value(row, col_name("CA"))
            if kommentar_celle == "":
                continue
            else:
                match_obj = re.findall(koder_benevning, kommentar_celle, re.I | re.X)
                if match_obj:
                    [koder_list.append(item) for item in match_obj]
        return koder_list
            
            
        
    # leser en kolonne fra top til bunn og kopierer innhold
    # returner liste dersom normal, returner None dersom kolonna er tom
    def __search_col(self, col, group_obj):
        
        row_code = []
        row = group_obj.first_row # første rad i siste balise-objekt fra liste
        if (self.wb_sheet.cell(row, col).ctype == 2) or (self.wb_sheet.cell_value(row, col) != ""): # hvis har innhold
            row_code.append(
                    self.wb_sheet.cell_value(row, col) # les kode fra celle
                    )
        else: # hvis ikke innhold
            # row_code.append(None) # returner liste med None per rad
            return None # returner None i stedet for en liste
        
        if group_obj.first_row == group_obj.last_row:
            return self.__make_int(row_code)
        else:
            for row in range(group_obj.first_row + 1, group_obj.last_row + 1):
                if (self.wb_sheet.cell(row, col).ctype == 2) or (self.wb_sheet.cell_value(row, col) != ""): # hvis har innhold
                    row_code.append(
                            self.wb_sheet.cell_value(row, col) # les kode fra celle
                            )
                else: # hvis ikke innhold
                    row_code.append(row_code[-1]) # kopierer kode fra forrige linje
            return self.__make_int(row_code)
              
    # finner antall rader en balisegruppe strekker seg over
    def __last_row(self, first_row):
        last_row = first_row      
        for key in self.ktab_cols:
            col = col_name(self.ktab_cols[key])
            row = first_row
            while True:
                if (self.wb_sheet.cell(row + 1, col).ctype == 2) or (self.wb_sheet.cell_value(row + 1, col) == "-"): # hvis cellen ikke er tom
                    row += 1
                else:
                    break
            if row > last_row:
                last_row = row
        return last_row

    # del av search_col()
    def __make_int (self, aList):
        newList = []
        for string in aList:
            try:
                newList.append(int(string))
            except:
                newList.append(string)
        if len(aList) != len(newList):
            print("__make_int error")
        return newList
    
    # fjerner rusk fra KM og returnerer en int
    def __clean_KM(self, KM_str):
        from re import findall
        KM_str = str(KM_str)
        if KM_str.isdigit():
            return KM_str
        else:
            try:
                KM_str = "".join(findall("[0-9]", KM_str))
                return int(KM_str)
            except:
                print(KM_str)
                print(findall("[0-9]", KM_str))
                return -1.0


class Balisegruppe:
    def __init__(self, sign_type, id1, id2, km, ktab_retning, s_nr, first_row, last_row):
        self.sign_type = sign_type
        self.id1 = id1
        self.id2 = id2
        self.km = km
        self.ktab_retning = ktab_retning
        self.s_nr = s_nr
        self.first_row = first_row
        self.last_row = last_row
        self.tilstander = None
        self.kodere = []
        self.sim_segment = None # segment dersom den skal brukes i ATC sim
        self.baliser = []
        
        self.finn_retning()
        self.finn_type()
        
        # setter retning avhengig av om id2 er odde er partall
    def finn_retning(self):
        m = re.match("\d+", self.id2[::-1])
        try:
            nr = int(m.group(0)[::-1])
            if nr % 2 == 0:
                self.retning = "B"
            else:
                self.retning = "A"
        except:
            self.retning = "?"

    # klassifiserer etter type balisegruppe        
    def finn_type(self):
        # https://trv.banenor.no/wiki/Signal/Prosjektering/ATC#Baliseidentitet
        tabell_12 = {
                "H.sign": ["_", "M", "O", "S", "Y", "Æ", "Å", "L", "N", "P", "T", "X", "Ø"],
                "D.sign": ["m", "o", "s", "y", "æ", "å", "l", "n", "p", "t", "x", "ø"],
                "F.sign": ["F"],
                "FF": ["Z"],
                # "Rep.": ["R", "U", "V", "W"],
                "Rep.": ["R", "U", "W"], # V er for SVG
                "L": ["L"],
                "SVG/RVG": ["V", "v"],
                "SH": ["S"],
                "H/H(K1)/H(K2)": ["H"],
                "ERH/EH/SEH": ["E"],
                "GMD/GMO/HG/BU/SU": ["G"]
                }
        for key in tabell_12:
            if self.id2[0] in tabell_12[key] or self.id2[1] in tabell_12[key]:
                self.type = key
        
    
    def __str__(self):
        self_str = "{}\t{} {}\t{}\t" .format(self.sign_type, self.id1, self.id2, self.km)
        return self_str


class Balise:
    def __init__(self, rang, x_reg, y_reg, z_reg):
        self.rang = rang # P, A, B, C eller N-balise
        self.x_reg = x_reg # X-ord
        self.y_reg = y_reg
        self.z_reg = z_reg
        self.km = 0
        
    def __str__(self):
        return ("{0}X: {1}\t{0}Y: {2}\t{0}Z: {3}" .format(
                self.rang, 
                self.x_reg, 
                self.y_reg, 
                self.z_reg
                ))    


class PD_table:
    import pandas as pd
    def __init__(self, ktab_liste):
        self.ktab_liste = ktab_liste
        
        self.pd_import = []
        for ktab in self.ktab_liste:
            for balise in ktab.balise_group_obj_list:
                self.pd_import.append({
                    "Sign./type" : balise.sign_type,
                    "Sted" : balise.id1,
                    "ID" : balise.id2,
                    "KM" : balise.km,
                    "Retning" : balise.retning,
                    "Tegning" : balise.s_nr,
                    "Rad nr." : "{}-{}" .format(balise.first_row+1, balise.last_row+1),
                    "Kodere" : len(balise.kodere)
                })
        
        self.balise_df = pd.DataFrame(self.pd_import)
        self.balise_df = self.balise_df[['Retning', 'Sign./type', 'Sted', 'ID', 'KM', 'Tegning', 'Rad nr.', 'Kodere']]
        
    def lagre_excel(self):
        self.balise_df.to_excel("gruppeliste.xlsx")
        
    def print_df(self):
        print(self.balise_df)



##############
# Funksjoner #
##############

# vasker kodeord for å presenteres i excel    
def rens_kodeord(kodeliste):
    
    kodeliste = set(kodeliste)
    
    if "-" in kodeliste:
        kodeliste.remove("-")
        if len(kodeliste) == 0:
            return 1 # kode "1" dersom koding er vilkårlig
    
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
    # listContent = workbook.add_format({"align": "center"})
    # tableHeader = workbook.add_format({"bold": True, "border": True, "align": "center"})

    # Definer tabell
    data = makeListOfLists(baliseDictList)
    baliseWorksheet.add_table(0,0, len(data), len(data[0])-1, {
        "data": data,
        "columns": makeHeaders(baliseDictList)
        # "header_row": True
        })
    
    # Opprydding
    workbook.close()
    return

# tar en bokstavkode, gir plass i alfabetet
def alphabet_number(some_char):
    return ord(some_char.upper())-64

def col_name(letter_string):
    sum = 0
    for idx, c in enumerate(reversed(letter_string)):
        sum += 26**idx*alphabet_number(c)
    return sum - 1

def makeListOfLists(DictList):
    return [list(dictionary.values()) for dictionary in DictList]

def makeHeaders(DictList):
    return [{"header": "{}" .format(key)} for key in DictList[0]]
