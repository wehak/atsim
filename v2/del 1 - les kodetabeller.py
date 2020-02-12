# -*- coding: utf-8 -*-
"""
Created on Jul  7 09:40:18 2019

@author: WEYHAK

Del 1:
Finner alle kodertabeller i "sti_til_kodetabeller" og forsøker å lage en samlet baliseliste i xlsx-format (excel)

"""


from atsim_class import Baliseoversikt
#from atsim_func import rens_kodeord
from atsim_func import skrivBaliseliste


# Mappe hvor kodetabeller finnes
sti_til_kodetabeller = r"C:\Users\weyhak\OneDrive - Bane NOR\Dokumenter\DIV\lysaker sim\kodetabeller\TTG2"

# Leser kodetabeller
alle_ark = Baliseoversikt()
alle_ark.ny_mappe(sti_til_kodetabeller)

# Lager excelark
skrivBaliseliste(alle_ark, "lysaker_ttg2.xlsx")