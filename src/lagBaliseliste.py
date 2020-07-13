from kodetabellHelpers import Baliseoversikt

mypath = r"C:\Users\weyhak\Desktop\temp\Ny mappe (7)"
    
# Leser kodetabeller
alle_ark = Baliseoversikt()
alle_ark.ny_mappe(mypath)
alle_ark.makeSQL("oslo_s.db")