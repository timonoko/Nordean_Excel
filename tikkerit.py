#! /usr/bin/python3

# Lisää tämä koodin alkuun
MAPPING = {
    "Huhtamäki Oyj": "HUH1V.HE",
    "Wärtsilä Oyj Abp": "WRT1V.HE",
    "Orion Oyj B": "ORNBV.HE",
    "Lassila & Tikanoja": "LASTIK.HE",
    "Volvo B": "VOLV-B.ST",
    "KONE Oyj":"KNEBV.HE",
    "Elisa Oyj": "ELISA.HE",
    "Suominen": "SUY1V.HE",
    "Luotea Plc": "LUOTEA.HE"
    }

    # ... jatkuu kuten aiemmin ...

import pandas as pd
import yfinance as yf
import requests
import warnings
import time

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

url = f"https://api.frankfurter.dev/v1/latest"
VALUUTAT = requests.get(url).json()['rates']

def get_live_rates(curr):
    if curr=='EUR': return 1.0
    try:
        return 1.0/VALUUTAT[curr]
    except:
            print("EI TOIMI")
            # Varasuunnitelma jos Yahoo takkuaa
            fallbacks = {'SEK': 0.088, 'NOK': 0.086, 'USD': 0.92, 'DKK': 0.134}
            rates = fallbacks[curr]
    return rates


def get_ticker_and_price(name):
    name = name.strip()
    
    # 1. Tarkistetaan löytyykö nimi suoraan mappaus-listalta
    if name in MAPPING:
        symbol = MAPPING[name]
        try:
            ticker = yf.Ticker(symbol)
            return symbol, ticker.fast_info['lastPrice']
        except:
            pass

    # 2. Jos ei löytynyt, yritetään hakua (parannettu haku)
    import urllib.parse
    clean_name = name.replace("Oyj", "").replace("Abp", "").strip()
    search_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={urllib.parse.quote(clean_name)}"
    """Etsii ticker-symbolin nimen perusteella ja hakee hinnan."""
    try:
        # 1. Etsi symboli
        search_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={name}"
        headers = {'User-Agent': 'Mozilla/5.0'}
        search_results = requests.get(search_url, headers=headers).json()
        
        if not search_results['quotes']:
            return None, None
            
        symbol = search_results['quotes'][0]['symbol']
        
        # 2. Hae hinta symbolilla
        ticker = yf.Ticker(symbol)
        # Käytetään 'fast_info' tai history riippuen saatavuudesta
        #price = ticker.fast_info['lastPrice']
        data = ticker.history(period="1mo")
        if not data.empty:
            price = data['Close'].iloc[-1]        
        return symbol, price
    except Exception:
        return None, None

# --- DATAN LATAUS ---
df = pd.read_excel('Document.vnd.openxmlformats-officedocument.spreadsheetml.sheet.xlsx')
# K=10 (Nimi), X=23 (Määrä/Arvo - tässä oletetaan että sarakkeessa X on osakkeiden määrä jos haluat laskea nykyhinnan mukaan)
# Jos sarakkeessa X on jo euroarvo, käytetään sitä suoraan.
stocks = df.iloc[2:, [10, 11, 9]].dropna()

results = []
total_market_value = 0

print(f"{'Nimi':<25} {'Ticker':<12} {'Hinta':>10} {'Yhteensä':>12}")
print("-" * 65)

for _, row in stocks.iterrows():
    name = str(row.iloc[0]).strip()
    amount = row.iloc[1] # Oletetaan tässä että tämä on kpl-määrä
    symbol, price = get_ticker_and_price(name)
    if symbol in ("LT5.SG","NHYDY"):
        valuutta=get_live_rates('USD')
    else: valuutta= get_live_rates(row.iloc[2])

    if price:
        price2 = price * valuutta
        current_value = amount * price2
        if valuutta==1.0: huom=""
        else: huom=f"{price:>10.2f}*{valuutta:.4f}"
        total_market_value += current_value
        print(f"{name[:23]:<25} {symbol:<12} {price2:>10.2f}€ {current_value:>10.2f}€ {huom}")
        results.append({'Nimi': name, 'Arvo': current_value})
    else:
        print(f"{name[:23]:<25} {'EI LÖYTYNYT':<12}")
    
    time.sleep(0.2) # Pieni tauko ettei Yahoo suutu


print("-" * 65)
print(f"Salkun markkina-arvo:  {total_market_value:>12.2f} €")

