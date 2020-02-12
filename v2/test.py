

egen_gruppe = ["P", "A", "B", "C"]
retning = 1
offset = 6
print(egen_gruppe)
print(egen_gruppe[::-1])
for i, bokstav in enumerate(egen_gruppe[::-1]):
  print(i, bokstav)
  balise_km = 100 + (offset + 3 * i) * retning
  print(bokstav, balise_km)