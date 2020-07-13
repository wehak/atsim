def clean_KM(KM_str):
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


testTall = 3855
print(clean_KM(testTall))

print(type(1.0) is float)