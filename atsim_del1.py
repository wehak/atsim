# -*- coding: utf-8 -*-
"""
Created on Jun  7 11:34:42 2019

@author: Haakon Weydahl, weyhak@banenor.no

Python 3.7

Leser kodetabeller i .xls format og lager et excel-ark som lar bruker
assosiere hver balisegruppe med et segment. Dette skal være input for neste
trinn 


"""

from atsim_class import Baliseoversikt
from atsim_func import definer_segmenter

# Input: mappe kodetabeller hentes i fra
sti_til_kodetabeller = r"C:\Users\weyhak\Desktop\temp\grorud"

# Output: filnavn og plassering
filnavn = r"definer_segmenter.xlsx"

# Leser kodetabeller
alle_ark = Baliseoversikt()
alle_ark.ny_mappe(sti_til_kodetabeller)

# lager dataframe og skriver til excel fil
baliseoversikt_df = definer_segmenter(alle_ark)
baliseoversikt_df.to_excel(filnavn, index = False)

#os.startfile(filnavn)
#print("Rediger regnearket og trykk lagre.")
#input("'ENTER' for å fortsette")
#baliseoversikt_df = pd.read_excel(filnavn) #, sheet_name='Sheet3')
