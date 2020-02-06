import xml.etree.ElementTree as ET
import xlrd
from atsim_class import XMLbalise



# Navn på XML-fil som skal genereres
xmlFilename = "simulering test.xml"

# Sti til excelark med baliser
excelFilename = "baliser test.xlsx"

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