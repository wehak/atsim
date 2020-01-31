# -*- coding: utf-8 -*-
"""
Created on Jul  7 09:40:18 2019

@author: WEYHAK

Leser et regneark laget i del 1. Assosierer alle balisegruppene med et segment
og lager et nytt regneark med alle de individuelle balisene.

"""

import pandas as pd

from atsim_class import Baliseoversikt
from atsim_func import rens_kodeord


# Input 1 : mappesti kodetabeller hentes i fra. må være samme som i del 1
sti_til_kodetabeller = r"C:\Users\weyhak\Desktop\temp\grorud"

# Input 2:  sti til fil med definisjon av segmener 
#           (default er samme sted somi del 1)
segmentfil = r"C:\Users\weyhak\OneDrive - Bane NOR\Dokumenter\Tools\Python\definer_segmenter_test.xlsx"

# Output:   sti til filen som skal lagres. Hvis ingen sti er angitt lagres fila
#           i samme mappe scriptet kjøres fra
baliselistefil = "baliseliste.xlsx"

# Leser kodetabeller
alle_ark = Baliseoversikt()
alle_ark.ny_mappe(sti_til_kodetabeller)

# søker etter matchene balise ID og legger segment til Balisegruppe() objekt
baliseoversikt_df = pd.read_excel(segmentfil)
for ktab in alle_ark.alle_ktab:
    for bgruppe in ktab.balise_group_obj_list:
        bgruppe.sim_segment = baliseoversikt_df.loc[(baliseoversikt_df['Sted'] == bgruppe.id1) & (baliseoversikt_df['ID'] == bgruppe.id2)]["Segment"].values[0]
segmenter = set([bgruppe.sim_segment for ktab in alle_ark.alle_ktab for bgruppe in ktab.balise_group_obj_list])

# lager dataframe som kan sendes til excel
to_pd = []
for ktab in alle_ark.alle_ktab:
    for bgruppe in ktab.balise_group_obj_list:
        for balise in bgruppe.baliser:
            to_pd.append({
                    "Retning": bgruppe.retning,
                    "Sign./Type": bgruppe.sign_type,
                    "Type": bgruppe.type,
                    "ID_sted": bgruppe.id1, 
                    "ID_type": bgruppe.id2, 
                    "Rang": balise.rang, 
                    "KM_prosjektert": balise.km, 
                    "KM_simulering": 0, 
                    "Tegning": bgruppe.s_nr, 
                    "Rad nr.": bgruppe.first_row + 1, 
                    "Segment": bgruppe.sim_segment,
                    "X-ord": rens_kodeord(balise.x_reg), 
                    "Y-ord": rens_kodeord(balise.y_reg), 
                    "Z-ord": rens_kodeord(balise.z_reg)
                    })

# definerer rekkefølge på kolonner
ut_df = pd.DataFrame(to_pd)            
ut_df = ut_df[[
        "Retning",
        "Sign./Type",
        "Type",
        "ID_sted", 
        "ID_type", 
        "Rang", 
        "KM_prosjektert", 
        "KM_simulering",
        "Tegning",
        "Rad nr.", 
        "Segment", 
        "X-ord",
        "Y-ord", 
        "Z-ord"
        ]]

# output
print(ut_df.sort_values(by = ["ID_sted", "ID_type"]))
#print(ut_df[["ID_sted", "ID_type", "KM_prosjektert"]].sort_values(by = ["ID_sted", "ID_type"]))
ut_df.sort_values(by = ["ID_sted", "ID_type"]).to_excel(baliselistefil, index = False)

