"""
Automatyzacja Raportu Excel z SAP
WERSJA FINALNA - z obsługą świąt i precyzyjnym rozpoznawaniem wierszy sumy

Wymagania: 
pip install openpyxl pandas holidays
"""

import openpyxl
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime, timedelta
import shutil
import os
import holidays

# ============================================
# KONFIGURACJA - DOSTOSUJ TE ŚCIEŻKI
# ============================================

SCIEZKA_GLOWNA = r"C:\Raporty"
SCIEZKA_EXPORT_SAP = r"C:\Exports\export.xlsx"
NAZWA_ARKUSZA_GLOWNEGO = "Arkusz 1"

# KOLUMNY Z PLIKU SAP
KOLUMNA_REFERENCJA = "Kod referencyjny 1"
KOLUMNA_KWOTA = "Kwota w walucie krajowej"

# DOZWOLONE KODY REFERENCYJNE (pierwsze 10 znaków)
DOZWOLONE_KODY = [
    "BREO2P1025",
    "BREO2P1026",
    "INV01P5918",
    "INV01P6885"
]

# Kolumny docelowe w arkuszu głównym (gdzie wklejamy dane)
KOLUMNA_CEL_REFERENCJA = "A"
KOLUMNA_CEL_KWOTA = "B"

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
    """
    Zwraca ostatni dzień roboczy danego miesiąca
    """
    # Znajdź ostatni dzień miesiąca
    if miesiac == 12:
        ostatni_dzien = datetime(rok, miesiac, 31)
    else:
        nastepny_miesiac = datetime(rok, miesiac + 1, 1)
        ostatni_dzien = nastepny_miesiac - timedelta(days=1)
    
    pl_holidays = holidays.Poland(years=[rok])
    
    # Cofaj się jeśli to weekend lub święto
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
    Znajduje wiersze sumy według reguł:
    1. Pierwsze 10 znaków kolumny "Kod referencyjny 1" = jeden z dozwolonych kodów
    2. Wszystkie inne kolumny są puste (oprócz Kod referencyjny 1 i Kwota w walucie krajowej)
    3. To ostatni wiersz w grupie o tym samym kodzie referencyjnym
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
    print(f"\n🎯 Szukam kodów: {', '.join(DOZWOLONE_KODY)}")
    
    # Grupuj wiersze według pierwszych 10 znaków kodu referencyjnego
    grupy = {}
    
    for idx, row in df_sap.iterrows():
        ref = str(row[KOLUMNA_REFERENCJA]) if pd.notna(row[KOLUMNA_REFERENCJA]) else ""
        
        # Sprawdź czy ma minimum 10 znaków
        if len(ref) < 10:
            continue
        
        # Pobierz pierwsze 10 znaków
        kod = ref[:10]
        
        # Sprawdź czy to jeden z dozwolonych kodów
        if kod not in DOZWOLONE_KODY:
            continue
        
        # Dodaj do grupy
        if kod not in grupy:
            grupy[kod] = []
        
        grupy[kod].append({
            'indeks': idx,
            'referencja': ref,
            'kwota': row[KOLUMNA_KWOTA],
            'wiersz': row
        })
    
    print(f"\n📊 Znaleziono grupy:")
    for kod, wiersze in grupy.items():
        print(f"   {kod}: {len(wiersze)} wierszy")
    
    # Dla każdej grupy weź ostatni wiersz (gdzie inne kolumny są puste)
    wiersze_sumy = []
    
    for kod, wiersze_grupy in grupy.items():
        # Sprawdź każdy wiersz od końca grupy
        for wiersz_data in reversed(wiersze_grupy):
            wiersz = wiersz_data['wiersz']
            
            # Sprawdź czy inne kolumny są puste
            kolumny_do_sprawdzenia = [col for col in df_sap.columns 
                                     if col not in [KOLUMNA_REFERENCJA, KOLUMNA_KWOTA]]
            
            czy_puste = all(
                pd.isna(wiersz[col]) or str(wiersz[col]).strip() == '' 
                for col in kolumny_do_sprawdzenia
            )
            
            # Jeśli znaleziono wiersz sumy (wszystkie inne kolumny puste)
            if czy_puste:
                wiersze_sumy.append({
                    'kod': kod,
                    'referencja': wiersz_data['referencja'],
                    'kwota': wiersz_data['kwota'],
                    'indeks': wiersz_data['indeks']
                })
                print(f"   ✅ {kod}: wiersz {wiersz_data['indeks'] + 2} (Excel)")
                break  # Bierzemy tylko ostatni wiersz z grupy
        else:
            # Jeśli nie znaleziono wiersza z pustymi kolumnami, weź ostatni
            ostatni = wiersze_grupy[-1]
            wiersze_sumy.append({
                'kod': kod,
                'referencja': ostatni['referencja'],
                'kwota': ostatni['kwota'],
                'indeks': ostatni['indeks']
            })
            print(f"   ⚠️ {kod}: wiersz {ostatni['indeks'] + 2} (ostatni w grupie)")
    
    print(f"\n✅ Znaleziono {len(wiersze_sumy)} wierszy sumy")
    
    return wiersze_sumy

# ============================================
# AKTUALIZACJA POWER QUERY
# ============================================

def aktualizuj_daty_power_query(plik_excel, data_start, data_end):
    """
    Aktualizuje daty w arkuszu Parametry dla Power Query
    WAŻNE: Musisz mieć arkusz 'Parametry' i w Power Query odwoływać się do:
    - komórki A2 jako DataStart
    - komórki B2 jako DataEnd
    """
    wb = load_workbook(plik_excel)
    
    if 'Parametry' not in wb.sheetnames:
        print("⚠️ UWAGA: Arkusz 'Parametry' nie istnieje. Tworzę go...")
        ws = wb.create_sheet('Parametry', 0)  # Dodaj jako pierwszy arkusz
        ws['A1'] = "DataStart"
        ws['B1'] = "DataEnd"
    else:
        ws = wb['Parametry']
    
    # Zapisz daty w wierszu 2
    ws['A2'] = data_start.strftime("%Y-%m-%d")
    ws['B2'] = data_end.strftime("%Y-%m-%d")
    
    wb.save(plik_excel)
    print(f"\n✅ Zaktualizowano daty Power Query:")
    print(f"   DataStart (A2): {data_start.date()}")
    print(f"   DataEnd (B2): {data_end.date()}")
    print(f"\n💡 PAMIĘTAJ: W Power Query użyj:")
    print(f"   let")
    print(f"       DataStart = Excel.CurrentWorkbook(){{[Name=\"Parametry\"]}}[Content]{{1}}[DataStart],")
    print(f"       DataEnd = Excel.CurrentWorkbook(){{[Name=\"Parametry\"]}}[Content]{{1}}[DataEnd],")
    print(f"       ...")

# ============================================
# GŁÓWNA LOGIKA
# ============================================

def main():
    print("="*60)
    print("🚀 AUTOMATYZACJA RAPORTU EXCEL Z SAP")
    print("="*60)
    
    # KROK 1: Obliczanie dat (z uwzględnieniem świąt!)
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
        
        # Oblicz daty dla Power Query
        data_start = pobierz_ostatni_dzien_roboczy_miesiaca(wczoraj.year, wczoraj.month)
        
        # Data_End: pierwszy dzień następnego miesiąca (wyłącznie)
        if dzisiaj.month == 12:
            data_end = datetime(dzisiaj.year + 1, 1, 1)
        else:
            data_end = datetime(dzisiaj.year, dzisiaj.month + 1, 1)
        
        # Skopiuj poprzedni plik jako bazę
        if os.path.exists(plik_wczoraj):
            shutil.copy2(plik_wczoraj, plik_miesiac)
            print(f"✅ Utworzono nowy plik miesięczny: {plik_miesiac}")
            
            aktualizuj_daty_power_query(plik_miesiac, data_start, data_end)
        else:
            print(f"❌ BŁĄD: Nie znaleziono pliku źródłowego: {plik_wczoraj}")
            print(f"⚠️ Tworzę nowy plik od zera...")
            # Tutaj można dodać logikę tworzenia nowego pliku
            return
        
        plik_zrodlowy = plik_miesiac
    else:
        print("\n📁 Normalny dzień - kopiuję wczorajszy plik")
        plik_zrodlowy = plik_wczoraj
    
    # KROK 3: Kopiowanie pliku
    if not os.path.exists(plik_zrodlowy):
        print(f"❌ BŁĄD: Nie znaleziono pliku źródłowego: {plik_zrodlowy}")
        print("💡 Sprawdź czy wczoraj był dzień roboczy i czy plik istnieje")
        return
    
    shutil.copy2(plik_zrodlowy, plik_dzis)
    print(f"✅ Skopiowano plik: {os.path.basename(plik_zrodlowy)} → {os.path.basename(plik_dzis)}")
    
    # KROK 4: Otwarcie pliku Excel
    wb = load_workbook(plik_dzis)
    
    # KROK 5: Dodanie nowego arkusza
    ws_nowy = wb.create_sheet(nazwa_dzis)
    print(f"✅ Utworzono nowy arkusz: {nazwa_dzis}")
    
    # KROK 6: Import danych z export.xlsx
    if not os.path.exists(SCIEZKA_EXPORT_SAP):
        print(f"\n❌ BŁĄD: Nie znaleziono pliku SAP: {SCIEZKA_EXPORT_SAP}")
        print("⚠️ Upewnij się, że:")
        print("   1. Uruchomiłeś flow SAP do pobrania danych")
        print("   2. Plik został zapisany w poprawnej lokalizacji")
        wb.save(plik_dzis)
        return
    
    print(f"\n📥 Wczytuję dane z SAP...")
    df_sap = pd.read_excel(SCIEZKA_EXPORT_SAP)
    print(f"✅ Wczytano {len(df_sap)} wierszy, {len(df_sap.columns)} kolumn")
    
    # Zapisz dane SAP do nowego arkusza (z nagłówkami)
    for c_idx, col_name in enumerate(df_sap.columns, start=1):
        ws_nowy.cell(row=1, column=c_idx, value=col_name)
    
    for r_idx, row in enumerate(df_sap.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws_nowy.cell(row=r_idx, column=c_idx, value=value)
    
    print(f"✅ Wklejono dane SAP do arkusza '{nazwa_dzis}'")
    
    # KROK 7: Znajdź wiersze sumy
    wiersze_sumy = znajdz_wiersze_sumy(df_sap)
    
    if len(wiersze_sumy) == 0:
        print("\n⚠️ UWAGA: Nie znaleziono żadnych wierszy sumy!")
        print("Sprawdź czy:")
        print(f"   1. Kolumny '{KOLUMNA_REFERENCJA}' i '{KOLUMNA_KWOTA}' istnieją")
        print(f"   2. Dane zawierają kody: {', '.join(DOZWOLONE_KODY)}")
        wb.save(plik_dzis)
        return
    
    print(f"\n📋 Szczegóły znalezionych wierszy:")
    for i, wiersz in enumerate(wiersze_sumy, 1):
        print(f"   {i}. {wiersz['kod']}")
        print(f"      Referencja: {wiersz['referencja']}")
        print(f"      Kwota: {wiersz['kwota']}")
    
    # KROK 8: Wklej dane do arkusza głównego
    if NAZWA_ARKUSZA_GLOWNEGO not in wb.sheetnames:
        print(f"\n⚠️ UWAGA: Arkusz '{NAZWA_ARKUSZA_GLOWNEGO}' nie istnieje!")
        print("Tworzę nowy arkusz...")
        ws_glowny = wb.create_sheet(NAZWA_ARKUSZA_GLOWNEGO)
    else:
        ws_glowny = wb[NAZWA_ARKUSZA_GLOWNEGO]
    
    pierwszy_wolny = ws_glowny.max_row + 1
    
    print(f"\n📝 Wklejam dane do arkusza '{NAZWA_ARKUSZA_GLOWNEGO}'...")
    for idx, wiersz in enumerate(wiersze_sumy):
        wiersz_cel = pierwszy_wolny + idx
        ws_glowny[f"{KOLUMNA_CEL_REFERENCJA}{wiersz_cel}"] = wiersz['referencja']
        ws_glowny[f"{KOLUMNA_CEL_KWOTA}{wiersz_cel}"] = wiersz['kwota']
        print(f"   Wiersz {wiersz_cel}: {wiersz['kod']} → {wiersz['kwota']}")
    
    print(f"✅ Wklejono {len(wiersze_sumy)} wierszy (od wiersza {pierwszy_wolny})")
    
    # KROK 9: Zapisz plik
    wb.save(plik_dzis)
    print(f"\n💾 Zapisano plik: {plik_dzis}")
    
    print("\n" + "="*60)
    print("🎉 RAPORT WYGENEROWANY POMYŚLNIE!")
    print("="*60)
    
    print("\n⚠️ OSTATNI KROK - ODŚWIEŻENIE POWER QUERY:")
    print("1. Otwórz plik w Excel")
    print("2. Naciśnij Ctrl+Alt+F5 (lub Dane → Odśwież wszystko)")
    print("3. Poczekaj na zakończenie odświeżania")
    print("4. Zapisz plik (Ctrl+S)")
    print("\n💡 TIP: Możesz to zautomatyzować przez Power Automate Desktop!")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ BŁĄD: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nNaciśnij Enter aby zamknąć...")
