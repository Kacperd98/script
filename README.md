"""
Automatyzacja Raportu Excel z SAP
WERSJA POPRAWIONA - Używa win32com zamiast openpyxl (bezpieczniejsze dla tabel!)

Wymagania: 
pip install pandas holidays pywin32
"""

import pandas as pd
from datetime import datetime, timedelta
import shutil
import os
import holidays
import win32com.client
import pythoncom

# ============================================
# KONFIGURACJA - DOSTOSUJ TE ŚCIEŻKI
# ============================================

SCIEZKA_GLOWNA = r"C:\Raporty"
SCIEZKA_EXPORT_SAP = r"C:\Exports\export.xlsx"
NAZWA_ARKUSZA_GLOWNEGO = "Arkusz 1"
NAZWA_TABELI_SAP = "DaneSAP"

# KOLUMNY Z PLIKU SAP
KOLUMNA_REFERENCJA = "Kod referencyjny 1"
KOLUMNA_KWOTA = "Kwota w walucie krajowej"

# DOZWOLONE KODY W KOLEJNOŚCI (WAŻNE!)
DOZWOLONE_KODY_KOLEJNOSC = [
    "INV01P6885",
    "INV01P5918",
    "BREO2P1025",
    "BREO2P1026"
]

# Kolumny docelowe w tabeli DaneSAP
KOLUMNA_CEL_REFERENCJA = "I"  # Kolumna I - Referencje
KOLUMNA_CEL_KWOTA = "J"        # Kolumna J - Suma kwoty

# ============================================
# FUNKCJE - DNI ROBOCZE I ŚWIĘTA
# ============================================

def pobierz_poprzedni_dzien_roboczy():
    """
    Zwraca poprzedni dzień roboczy
    Pomija weekendy I ŚWIĘTA PAŃSTWOWE W POLSCE
    """
    dzisiaj = datetime.now()
    poprzedni = dzisiaj - timedelta(days=1)
    
    # Pobierz święta w Polsce
    pl_holidays = holidays.Poland(years=[poprzedni.year, dzisiaj.year])
    
    # Cofaj się jeśli to weekend lub święto
    while poprzedni.weekday() >= 5 or poprzedni in pl_holidays:
        if poprzedni in pl_holidays:
            print(f"   ⚠️ Pomijam święto: {poprzedni.date()} - {pl_holidays.get(poprzedni)}")
        poprzedni -= timedelta(days=1)
    
    return poprzedni

def pobierz_ostatni_dzien_roboczy_miesiaca(rok, miesiac):
    """Zwraca ostatni dzień roboczy danego miesiąca"""
    if miesiac == 12:
        ostatni_dzien = datetime(rok, miesiac, 31)
    else:
        nastepny_miesiac = datetime(rok, miesiac + 1, 1)
        ostatni_dzien = nastepny_miesiac - timedelta(days=1)
    
    pl_holidays = holidays.Poland(years=[rok])
    
    while ostatni_dzien.weekday() >= 5 or ostatni_dzien in pl_holidays:
        ostatni_dzien -= timedelta(days=1)
    
    return ostatni_dzien

def czy_nowy_miesiac(dzisiaj, wczoraj):
    """Sprawdza czy jest pierwszy dzień roboczy miesiąca"""
    return dzisiaj.month != wczoraj.month

def formatuj_date_plik(data):
    """Formatuje datę do formatu ddMM"""
    return data.strftime("%d%m")

def formatuj_date_miesiac(data):
    """Formatuje datę do formatu MMyyyy"""
    return data.strftime("%m%Y")

# ============================================
# ROZPOZNAWANIE WIERSZY SUMY
# ============================================

def znajdz_wiersze_sumy(df_sap):
    """
    Znajduje wiersze sumy dla każdego dozwolonego kodu.
    Zwraca w OKREŚLONEJ KOLEJNOŚCI: INV01P6885, INV01P5918, BREO2P1025, BREO2P1026
    """
    print("\n" + "="*60)
    print("🔍 ROZPOZNAWANIE WIERSZY SUMY")
    print("="*60)
    
    # Sprawdź czy kolumny istnieją
    if KOLUMNA_REFERENCJA not in df_sap.columns:
        print(f"❌ BŁĄD: Nie znaleziono kolumny '{KOLUMNA_REFERENCJA}'")
        print(f"Dostępne kolumny: {df_sap.columns.tolist()}")
        return []
    
    if KOLUMNA_KWOTA not in df_sap.columns:
        print(f"❌ BŁĄD: Nie znaleziono kolumny '{KOLUMNA_KWOTA}'")
        print(f"Dostępne kolumny: {df_sap.columns.tolist()}")
        return []
    
    print(f"✅ Znaleziono kolumny:")
    print(f"   - Referencja: {KOLUMNA_REFERENCJA}")
    print(f"   - Kwota: {KOLUMNA_KWOTA}")
    
    # Grupuj wiersze według pierwszych 10 znaków kodu referencyjnego
    grupy = {}
    
    for idx, row in df_sap.iterrows():
        ref = str(row[KOLUMNA_REFERENCJA]) if pd.notna(row[KOLUMNA_REFERENCJA]) else ""
        
        # Usuń białe znaki
        ref = ref.strip()
        
        if len(ref) < 10:
            continue
        
        # Pobierz pierwsze 10 znaków i upewnij się że nie ma białych znaków
        kod = ref[:10].strip()
        
        # Debug - pokaż co znalazło
        if any(dozwolony in kod for dozwolony in DOZWOLONE_KODY_KOLEJNOSC):
            print(f"   🔍 Znaleziono potencjalny kod: '{kod}' (długość: {len(kod)})")
        
        # Sprawdź dokładne dopasowanie
        if kod not in DOZWOLONE_KODY_KOLEJNOSC:
            continue
        
        if kod not in grupy:
            grupy[kod] = []
        
        grupy[kod].append({
            'indeks': idx,
            'referencja': ref,
            'kwota': row[KOLUMNA_KWOTA],
            'wiersz': row
        })
    
    print(f"\n📊 Znaleziono grupy:")
    for kod in DOZWOLONE_KODY_KOLEJNOSC:
        if kod in grupy:
            print(f"   {kod}: {len(grupy[kod])} wierszy")
        else:
            print(f"   {kod}: BRAK DANYCH ⚠️")
    
    # Dla każdej grupy weź ostatni wiersz (sumy)
    wyniki = {}
    
    for kod in DOZWOLONE_KODY_KOLEJNOSC:
        if kod not in grupy:
            print(f"   ⚠️ {kod}: BRAK - dodaję puste wartości")
            wyniki[kod] = {
                'kod': kod,
                'referencja': kod,  # Wpisz sam kod jeśli brak danych
                'kwota': 0,
                'indeks': -1
            }
            continue
        
        wiersze_grupy = grupy[kod]
        
        # Sprawdź wiersze od końca grupy (ostatni wiersz)
        znaleziono = False
        for wiersz_data in reversed(wiersze_grupy):
            wiersz = wiersz_data['wiersz']
            
            # Sprawdź czy inne kolumny są puste
            kolumny_do_sprawdzenia = [col for col in df_sap.columns 
                                     if col not in [KOLUMNA_REFERENCJA, KOLUMNA_KWOTA]]
            
            czy_puste = all(
                pd.isna(wiersz[col]) or str(wiersz[col]).strip() == '' 
                for col in kolumny_do_sprawdzenia
            )
            
            if czy_puste:
                wyniki[kod] = {
                    'kod': kod,
                    'referencja': wiersz_data['referencja'],
                    'kwota': wiersz_data['kwota'],
                    'indeks': wiersz_data['indeks']
                }
                print(f"   ✅ {kod}: wiersz {wiersz_data['indeks'] + 2} (Excel) - kwota: {wiersz_data['kwota']}")
                znaleziono = True
                break
        
        if not znaleziono:
            # Weź ostatni wiersz
            ostatni = wiersze_grupy[-1]
            wyniki[kod] = {
                'kod': kod,
                'referencja': ostatni['referencja'],
                'kwota': ostatni['kwota'],
                'indeks': ostatni['indeks']
            }
            print(f"   ⚠️ {kod}: wiersz {ostatni['indeks'] + 2} (ostatni w grupie) - kwota: {ostatni['kwota']}")
    
    # Zwróć w OKREŚLONEJ KOLEJNOŚCI
    wiersze_w_kolejnosci = [wyniki[kod] for kod in DOZWOLONE_KODY_KOLEJNOSC]
    
    print(f"\n✅ Dane do wklejenia (w kolejności):")
    for i, w in enumerate(wiersze_w_kolejnosci, 1):
        print(f"   {i}. {w['kod']}: Ref='{w['referencja'][:20]}...' Kwota={w['kwota']}")
    
    return wiersze_w_kolejnosci

# ============================================
# AKTUALIZACJA POWER QUERY
# ============================================

def aktualizuj_daty_power_query(plik_excel, data_start, data_end):
    """Aktualizuje daty w arkuszu Parametry dla Power Query - używa COM API"""
    print(f"\n📅 Aktualizacja dat Power Query...")
    
    # Inicjalizuj COM
    pythoncom.CoInitialize()
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(plik_excel)
        
        # Sprawdź czy arkusz Parametry istnieje
        arkusz_istnieje = False
        for ws in wb.Worksheets:
            if ws.Name == "Parametry":
                arkusz_istnieje = True
                ws_param = ws
                break
        
        if not arkusz_istnieje:
            print("⚠️ Tworzę arkusz 'Parametry'...")
            ws_param = wb.Worksheets.Add()
            ws_param.Name = "Parametry"
            ws_param.Cells(1, 1).Value = "DataStart"
            ws_param.Cells(1, 2).Value = "DataEnd"
        
        # Zapisz daty
        ws_param.Cells(2, 1).Value = data_start.strftime("%Y-%m-%d")
        ws_param.Cells(2, 2).Value = data_end.strftime("%Y-%m-%d")
        
        wb.Save()
        wb.Close()
        excel.Quit()
        
        print(f"✅ Zaktualizowano daty Power Query:")
        print(f"   DataStart: {data_start.date()}")
        print(f"   DataEnd: {data_end.date()}")
        
    finally:
        pythoncom.CoUninitialize()

# ============================================
# GŁÓWNA LOGIKA
# ============================================

def main():
    print("="*60)
    print("🚀 AUTOMATYZACJA RAPORTU EXCEL Z SAP")
    print("="*60)
    
    # KROK 1: Obliczanie dat
    dzisiaj = datetime.now()
    wczoraj = pobierz_poprzedni_dzien_roboczy()
    
    nazwa_dzis = formatuj_date_plik(dzisiaj)
    nazwa_wczoraj = formatuj_date_plik(wczoraj)
    nazwa_miesiac = formatuj_date_miesiac(dzisiaj)
    
    plik_wczoraj = os.path.join(SCIEZKA_GLOWNA, f"{nazwa_wczoraj}.xlsx")
    plik_dzis = os.path.join(SCIEZKA_GLOWNA, f"{nazwa_dzis}.xlsx")
    plik_miesiac = os.path.join(SCIEZKA_GLOWNA, f"Raport_{nazwa_miesiac}.xlsx")
    
    print(f"\n📅 Dzisiaj: {dzisiaj.strftime('%Y-%m-%d (%A)')}")
    print(f"📅 Poprzedni dzień roboczy: {wczoraj.strftime('%Y-%m-%d (%A)')}")
    
    # KROK 2: Sprawdzenie czy nowy miesiąc
    if czy_nowy_miesiac(dzisiaj, wczoraj):
        print("\n🎉 WYKRYTO NOWY MIESIĄC!")
        
        data_start = pobierz_ostatni_dzien_roboczy_miesiaca(wczoraj.year, wczoraj.month)
        
        if dzisiaj.month == 12:
            data_end = datetime(dzisiaj.year + 1, 1, 1)
        else:
            data_end = datetime(dzisiaj.year, dzisiaj.month + 1, 1)
        
        if os.path.exists(plik_wczoraj):
            shutil.copy2(plik_wczoraj, plik_miesiac)
            print(f"✅ Utworzono nowy plik miesięczny: {plik_miesiac}")
            aktualizuj_daty_power_query(plik_miesiac, data_start, data_end)
        else:
            print(f"❌ BŁĄD: Nie znaleziono pliku źródłowego: {plik_wczoraj}")
            return
        
        plik_zrodlowy = plik_miesiac
    else:
        print("\n📁 Normalny dzień - kopiuję wczorajszy plik")
        plik_zrodlowy = plik_wczoraj
    
    # KROK 3: Kopiowanie pliku
    if not os.path.exists(plik_zrodlowy):
        print(f"❌ BŁĄD: Nie znaleziono pliku źródłowego: {plik_zrodlowy}")
        return
    
    shutil.copy2(plik_zrodlowy, plik_dzis)
    print(f"✅ Skopiowano plik: {os.path.basename(plik_zrodlowy)} → {os.path.basename(plik_dzis)}")
    
    # KROK 4: Otwórz Excel przez COM API (bezpieczne dla tabel!)
    print(f"\n📂 Otwieram plik przez Excel COM API: {plik_dzis}")
    
    # Inicjalizuj COM
    pythoncom.CoInitialize()
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Ukryty Excel
        excel.DisplayAlerts = False  # Bez alertów
        
        wb = excel.Workbooks.Open(os.path.abspath(plik_dzis))
        
        # KROK 5: Sprawdź czy arkusz już istnieje i usuń go
        print(f"\n📊 Przygotowuję arkusz z danymi SAP: {nazwa_dzis}")
        
        for ws in wb.Worksheets:
            if ws.Name == nazwa_dzis:
                print(f"   ⚠️ Usuwam istniejący arkusz {nazwa_dzis}")
                ws.Delete()
                break
        
        # Dodaj nowy arkusz
        ws_nowy = wb.Worksheets.Add()
        ws_nowy.Name = nazwa_dzis
        
        # KROK 6: Wczytaj dane z export.xlsx
        if not os.path.exists(SCIEZKA_EXPORT_SAP):
            print(f"\n❌ BŁĄD: Nie znaleziono pliku SAP: {SCIEZKA_EXPORT_SAP}")
            print("⚠️ Uruchom najpierw flow SAP do pobrania danych!")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return
        
        print(f"\n📥 Wczytuję dane z SAP...")
        df_sap = pd.read_excel(SCIEZKA_EXPORT_SAP)
        print(f"✅ Wczytano {len(df_sap)} wierszy, {len(df_sap.columns)} kolumn")
        
        # KROK 7: Wklej dane SAP do nowego arkusza przez COM
        print(f"📋 Kopiuję dane do arkusza {nazwa_dzis}...")
        
        # Nagłówki
        for c_idx, col_name in enumerate(df_sap.columns, start=1):
            ws_nowy.Cells(1, c_idx).Value = col_name
        
        # Dane - wklej całą tabelę naraz (szybsze!)
        dane_array = df_sap.values.tolist()
        
        if len(dane_array) > 0:
            # Określ zakres
            end_row = len(dane_array) + 1
            end_col = len(df_sap.columns)
            
            # Konwertuj dane na format dla Excel
            for r_idx, row in enumerate(dane_array, start=2):
                for c_idx, value in enumerate(row, start=1):
                    # Konwertuj NaN na None (puste komórki)
                    if pd.isna(value):
                        ws_nowy.Cells(r_idx, c_idx).Value = None
                    else:
                        ws_nowy.Cells(r_idx, c_idx).Value = value
        
        print(f"✅ Wklejono wszystkie dane SAP do arkusza '{nazwa_dzis}'")
        
        # KROK 8: Znajdź wiersze sumy
        wiersze_sumy = znajdz_wiersze_sumy(df_sap)
        
        if len(wiersze_sumy) == 0:
            print("\n⚠️ UWAGA: Nie znaleziono żadnych wierszy sumy!")
            wb.Save()
            wb.Close()
            excel.Quit()
            return
        
        # KROK 9: Znajdź arkusz główny
        print(f"\n📝 Szukam arkusza '{NAZWA_ARKUSZA_GLOWNEGO}'...")
        
        ws_glowny = None
        for ws in wb.Worksheets:
            if ws.Name == NAZWA_ARKUSZA_GLOWNEGO:
                ws_glowny = ws
                break
        
        if ws_glowny is None:
            print(f"❌ BŁĄD: Arkusz '{NAZWA_ARKUSZA_GLOWNEGO}' nie istnieje!")
            wb.Save()
            wb.Close()
            excel.Quit()
            return
        
        # KROK 10: Znajdź pierwszy wolny wiersz (zaczynając od wiersza 4)
        pierwszy_wolny = 4
        
        while ws_glowny.Cells(pierwszy_wolny, ord(KOLUMNA_CEL_REFERENCJA) - ord('A') + 1).Value is not None:
            pierwszy_wolny += 1
        
        print(f"✅ Pierwszy wolny wiersz w tabeli DaneSAP: {pierwszy_wolny}")
        
        # KROK 11: Wklej dane w określonej kolejności
        print(f"\n✍️ Wklejam dane do kolumn {KOLUMNA_CEL_REFERENCJA} i {KOLUMNA_CEL_KWOTA}...")
        print(f"   Kolejność: {' → '.join(DOZWOLONE_KODY_KOLEJNOSC)}")
        
        col_ref_num = ord(KOLUMNA_CEL_REFERENCJA) - ord('A') + 1
        col_kwota_num = ord(KOLUMNA_CEL_KWOTA) - ord('A') + 1
        
        for idx, wiersz in enumerate(wiersze_sumy):
            wiersz_cel = pierwszy_wolny + idx
            
            # Wpisz referencję
            ws_glowny.Cells(wiersz_cel, col_ref_num).Value = wiersz['referencja']
            
            # Wpisz kwotę
            ws_glowny.Cells(wiersz_cel, col_kwota_num).Value = wiersz['kwota']
            
            ref_display = wiersz['referencja'][:20] + "..." if len(wiersz['referencja']) > 20 else wiersz['referencja']
            print(f"   Wiersz {wiersz_cel}: {wiersz['kod']} → Ref: {ref_display} | Kwota: {wiersz['kwota']}")
        
        print(f"\n✅ Wklejono {len(wiersze_sumy)} wierszy do tabeli DaneSAP")
        
        # KROK 12: Zapisz i zamknij
        print(f"\n💾 Zapisuję plik...")
        wb.Save()
        wb.Close()
        excel.Quit()
        
        print(f"✅ Zapisano: {plik_dzis}")
        
    except Exception as e:
        print(f"\n❌ BŁĄD podczas pracy z Excel: {str(e)}")
        try:
            excel.Quit()
        except:
            pass
        raise
    
    finally:
        pythoncom.CoUninitialize()
    
    print("\n" + "="*60)
    print("🎉 RAPORT WYGENEROWANY POMYŚLNIE!")
    print("="*60)
    
    print(f"\n📊 Podsumowanie:")
    print(f"   📁 Plik: {os.path.basename(plik_dzis)}")
    print(f"   📋 Nowy arkusz: {nazwa_dzis}")
    print(f"   📈 Dane w tabeli DaneSAP: {len(wiersze_sumy)} wierszy (od wiersza {pierwszy_wolny})")
    
    print(f"\n⚠️ OSTATNI KROK - RĘCZNIE:")
    print(f"   1. Otwórz plik: {plik_dzis}")
    print(f"   2. Naciśnij Ctrl+Alt+F5 (Odśwież wszystko)")
    print(f"   3. Zapisz (Ctrl+S)")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ BŁĄD: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nNaciśnij Enter aby zamknąć...")
