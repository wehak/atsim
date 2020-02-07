# -*- coding: utf-8 -*-
"""
Created on Jul  7 09:40:18 2019

@author: WEYHAK

Del 1:
Leser kodetabeller og lager et regneark med koding

"""


from atsim_class import Baliseoversikt
from atsim_func import rens_kodeord
from atsim_func import skrivBaliseliste


# Mappe hvor kodetabeller finnes
sti_til_kodetabeller = r"C:\Users\weyhak\Desktop\temp\lysaker1"

# Leser kodetabeller
alle_ark = Baliseoversikt()
alle_ark.ny_mappe(sti_til_kodetabeller)

# Lager excelark
skrivBaliseliste(alle_ark, "lysaker1.xlsx")