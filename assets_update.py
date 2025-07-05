import yfinance as yf
import pandas as pd

# 1. Wczytaj plik Excel (zachowując wszystkie arkusze!)
xls = pd.ExcelFile("portfel.xlsx")
df = xls.parse("assets")

# 2. Sprawdź, czy kolumna 'cena' istnieje
if 'cena' not in df.columns:
    raise ValueError("Brakuje kolumny 'cena' w arkuszu 'assets'.")

# 3. Aktualizuj kolumnę 'cena'
for i, row in df.iterrows():
    symbol = row['ticker']
    if symbol == "PLN":
        pass
    else:

        ticker = yf.Ticker(symbol)

        try:
            data = ticker.history(period="1d")
            price = data['Close'].iloc[-1]
            currency=str(ticker.info.get("currency", "Brak danych"))

            if currency == 'PLN':
                ratio = 1
            else:
                curr_ticker = yf.Ticker(currency+"PLN=X")
                data = curr_ticker.history(period="1d")
                ratio = data['Close'].iloc[-1]


            df.at[i, 'cena'] = price
            df.at[i, 'waluta'] = currency
            df.at[i, 'kurs/pln'] = ratio

            print(f"{symbol}: {price:.2f} : {ratio}")
        except Exception as e:
            print(f"Błąd dla {symbol}: {e}")
            df.at[i, 'cena'] = None  # lub zostaw bez zmian

# 4. Zapisz dane z powrotem do tego samego pliku i arkusza
with pd.ExcelWriter("portfel.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="assets", index=False)







#
# ###
# # Propozycja chata GPT na zaciągnięcie kursów walut
# ###
#
# # Lista par walutowych względem PLN
# waluty = ["USDPLN=X", "EURPLN=X"]  # PLNPLN=X dla PLN względem PLN = 1.0
# kursy = {}
# for symbol in waluty:
#     ticker = yf.Ticker(symbol)
#     data = ticker.history(period="1d")
#     try:
#         kurs = data['Close'].iloc[-1]
#     except Exception:
#         kurs = None
#     kursy[symbol] = kurs
#
# # Zamiana nazw na czytelne (możesz dostosować)
# kursy_czytelne = {
#     "USD": kursy["USDPLN=X"],
#     "EUR": kursy["EURPLN=X"],
#     "PLN": 1.0
# }
#
# # Wczytaj plik Excel
# xls = pd.ExcelFile("portfel.xlsx")
#
# # Załaduj arkusz "kursy" (jeśli nie istnieje, stworzymy DataFrame od nowa)
# try:
#     df_kursy = xls.parse("kursy")
# except ValueError:
#     df_kursy = pd.DataFrame()
#
# # Utwórz DataFrame z kursami
# df_kursy = pd.DataFrame(list(kursy_czytelne.items()), columns=["Waluta", "Kurs_do_PLN"])
#
# print(df_kursy)
#
# # Zapisz z powrotem do pliku Excel, nadpisując arkusz "kursy"
# with pd.ExcelWriter("portfel.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
#     df_kursy.to_excel(writer, sheet_name="kursy", index=False)
#
#
