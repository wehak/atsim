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