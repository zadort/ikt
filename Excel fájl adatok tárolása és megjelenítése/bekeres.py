import pandas as pd

# Adatok bekérése a felhasználótól
nevek = []
korok = []

while True:
    nev = input("Add meg a nevet (vagy nyomj Entert a kilépéshez): ")
    if not nev:
        break
    kor = input(f"Add meg {nev} korát: ")
    
    nevek.append(nev)
    korok.append(kor)

# Adatok tárolása DataFrame-ben
adatok = {'Név': nevek, 'Kor': korok}
df = pd.DataFrame(adatok)

# Adatok mentése Excel fájlba
file_nev = "adatok.xlsx"
df.to_excel(file_nev, index=False)

print(f"Az adatok sikeresen elmentve az {file_nev} fájlba.")

# Excel fájl beolvasása
beolvas = pd.read_excel(file_nev)

#Beolvasott adatok megjelenítése
print(beolvas)