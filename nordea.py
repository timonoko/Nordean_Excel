import pandas as pd
import warnings,sys

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

osakkeita=19

df = pd.read_excel(sys.argv[1])

subset = df.iloc[2:, [10, 23]].copy()
#subset.iloc[:, 1] = pd.to_numeric(subset.iloc[:, 1].astype(str).str.replace(',', '.').str.replace(r'\s+', '', regex=True), errors='coerce')
result = subset.dropna().sort_values(by=subset.columns[1], ascending=False)
print(result)

summa=0
nettosumma=0
print()
print(" Osake                          Arvo     Osto      Voitto  Hankintaolettama")
print()
for i in range(3,3+osakkeita):
    hinta=df.iloc[i,23]
    summa+=hinta
    osto=df.iloc[i,21]
    olettama=''
    if osto==0.0: # jos ostohintaa ei ole se peritty
        osto=hinta*0.4
        olettama='40%'
    elif osto/hinta<0.4: # muut ovat aina 5 vuotta vanhoja
        osto=hinta*0.2
        olettama='20%'
    netto=hinta-osto
    nettosumma+=netto
    print(f"{df.iloc[i,10]:27} {df.iloc[i,23]:9.2f} {df.iloc[i,21]:9.2f} {netto:9.2f}   {olettama} ")
    
print()
print(f"Yhteensä    = {summa:.2f}")
verotettu=summa-nettosumma*0.34
print(f"Todellisuus = {verotettu:.2f} {(summa-verotettu)/summa*100:.2f}%")
cash=df.iloc[24,ord('F')-ord('A')]+df.iloc[1,ord('F')-ord('A')]
print(f"Käteistä    = {cash:.2f}")     
print(f"Brutto      = {summa+cash:.2f}")     
