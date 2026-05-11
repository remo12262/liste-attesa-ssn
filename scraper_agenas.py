"""
Scraper Liste di Attesa SSN — AGENAS PNLA
Legge dati da portali open data e li carica su Google Sheets
"""
import os, requests, json, time, logging
from datetime import datetime
from io import StringIO
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
log = logging.getLogger(__name__)

SOGLIE = {"U": 3, "B": 10, "D": 30, "P": 120}
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

def connetti_sheets():
    try:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        log.error(f"Errore Google Sheets: {e}")
        return None

def scarica_dati_gov():
    """Cerca dataset liste attesa su dati.gov.it"""
    frames = []
    try:
        r = requests.get(
            "https://www.dati.gov.it/opendata/api/3/action/package_search",
            params={"q": "liste attesa sanitarie", "rows": 10}, timeout=15)
        pkgs = r.json().get("result", {}).get("results", [])
        log.info(f"Trovati {len(pkgs)} dataset su dati.gov.it")
        for pkg in pkgs[:3]:
            for res in pkg.get("resources", []):
                if res.get("format", "").upper() == "CSV":
                    try:
                        r2 = requests.get(res["url"], timeout=20)
                        if r2.status_code == 200 and len(r2.content) > 500:
                            df = pd.read_csv(StringIO(r2.text), sep=None, engine="python")
                            df["fonte"] = f"dati.gov.it/{pkg.get('name','')}"
                            frames.append(df)
                            log.info(f"  OK: {pkg.get('name')} — {len(df)} righe")
                            break
                    except Exception as e:
                        log.warning(f"  Errore download: {e}")
                    time.sleep(0.5)
    except Exception as e:
        log.warning(f"Errore dati.gov.it: {e}")
    return frames

def scarica_lazio():
    """Dati PS Lazio (open data in tempo reale)"""
    try:
        r = requests.get(
            "https://dati.lazio.it/catalog/api/3/action/datastore_search",
            params={"resource_id": "d5f94776-4c88-4f53-8ee5-e3e7df15f0ea", "limit": 1000},
            timeout=20)
        records = r.json().get("result", {}).get("records", [])
        if records:
            df = pd.DataFrame(records)
            df["regione"] = "Lazio"
            df["fonte"] = "OpenData-Lazio"
            log.info(f"Lazio: {len(df)} record")
            return df
    except Exception as e:
        log.warning(f"Errore Lazio: {e}")
    return None

def normalizza(df, fonte=None, regione=None):
    """Converte qualsiasi DataFrame nel formato standard"""
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    # Rinomina colonne comuni
    alias = {
        "regione": ["regione","region","Regione","REGIONE"],
        "prestazione": ["prestazione","Prestazione","descrizione","servizio","nome_prestazione"],
        "classe": ["classe","Classe","priorita","priority"],
        "tempo_medio_gg": ["tempo_medio_gg","tempo_medio","giorni","days","attesa_media"],
        "perc_nei_tempi": ["perc_nei_tempi","percentuale","pct","compliance"],
        "volume": ["volume","Volume","prestazioni","count","totale"],
        "mese": ["mese","Mese","periodo","month"],
    }
    for standard, nomi in alias.items():
        for n in nomi:
            if n in df.columns and standard not in df.columns:
                df = df.rename(columns={n: standard})
                break
    # Valori default
    if "regione" not in df.columns: df["regione"] = regione or "N/D"
    if "prestazione" not in df.columns: df["prestazione"] = "N/D"
    if "classe" not in df.columns: df["classe"] = "D"
    if "mese" not in df.columns: df["mese"] = datetime.now().strftime("%Y-%m")
    if "volume" not in df.columns: df["volume"] = 0
    if "tempo_medio_gg" not in df.columns: df["tempo_medio_gg"] = None
    if "perc_nei_tempi" not in df.columns: df["perc_nei_tempi"] = None
    # Normalizza classe e soglia
    df["classe"] = df["classe"].astype(str).str.upper().str.strip()
    df["classe"] = df["classe"].apply(lambda x: x if x in SOGLIE else "D")
    df["soglia_gg"] = df["classe"].map(SOGLIE)
    # Calcola % nei tempi se mancante
    for col in ["tempo_medio_gg", "perc_nei_tempi"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    mask = df["perc_nei_tempi"].isna() & df["tempo_medio_gg"].notna()
    df.loc[mask, "perc_nei_tempi"] = df.loc[mask].apply(
        lambda r: max(0, min(100, 100*(1-max(0,r["tempo_medio_gg"]-r["soglia_gg"])/max(1,r["soglia_gg"])))), axis=1)
    df["volume"] = pd.to_numeric(df["volume"], errors="coerce").fillna(0).astype(int)
    df["fonte"] = fonte or df.get("fonte", "N/D")
    df["aggiornato_il"] = datetime.now().isoformat()
    # Stato semaforo
    df["stato"] = df["perc_nei_tempi"].apply(
        lambda p: "CRITICO" if p is not None and p < 50 else "ATTENZIONE" if p is not None and p < 75 else "OK")
    cols = ["regione","mese","prestazione","classe","soglia_gg","tempo_medio_gg","perc_nei_tempi","volume","stato","fonte","aggiornato_il"]
    for c in cols:
        if c not in df.columns: df[c] = None
    return df[cols].dropna(subset=["regione","prestazione"])

def carica_su_sheets(client, df):
    if not client or not SPREADSHEET_ID:
        log.warning("Google Sheets non configurato")
        return
    try:
        wb = client.open_by_key(SPREADSHEET_ID)
        try: ws = wb.worksheet("dati_live")
        except: ws = wb.add_worksheet("dati_live", rows=10000, cols=20)
        ws.clear()
        dati = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
        ws.update(dati, value_input_option="USER_ENTERED")
        log.info(f"✅ {len(df)} righe scritte su Google Sheets")
        # Metadati
        try: meta = wb.worksheet("_meta")
        except: meta = wb.add_worksheet("_meta", rows=10, cols=2)
        meta.update([["chiave","valore"],["aggiornato",datetime.now().isoformat()],["righe",len(df)]])
    except Exception as e:
        log.error(f"Errore scrittura Sheets: {e}")

def main():
    log.info("=" * 50)
    log.info("🏥 SCRAPER LISTE DI ATTESA SSN")
    log.info(f"   {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    log.info("=" * 50)

    frames = []

    # 1. dati.gov.it
    log.info("\n[1/2] dati.gov.it...")
    for df in scarica_dati_gov():
        n = normalizza(df)
        if not n.empty:
            frames.append(n)

    # 2. Lazio open data
    log.info("\n[2/2] Regione Lazio...")
    lazio = scarica_lazio()
    if lazio is not None:
        n = normalizza(lazio, fonte="OpenData-Lazio", regione="Lazio")
        if not n.empty:
            frames.append(n)

    if not frames:
        log.warning("Nessun dato scaricato dalle fonti online.")
        log.warning("→ Carica manualmente il CSV da AGENAS nella dashboard.")
        return

    # Unisci e deduplicа
    df_finale = pd.concat(frames, ignore_index=True)
    df_finale = df_finale.drop_duplicates(subset=["regione","mese","prestazione","classe"], keep="first")

    # Salva CSV locale
    os.makedirs("output", exist_ok=True)
    df_finale.to_csv("output/latest.csv", index=False, encoding="utf-8-sig")
    log.info(f"\n✅ CSV salvato: output/latest.csv ({len(df_finale)} record)")

    # Carica su Google Sheets
    client = connetti_sheets()
    carica_su_sheets(client, df_finale)

    log.info("\n📊 RIEPILOGO:")
    log.info(f"   Record totali: {len(df_finale)}")
    log.info(f"   Regioni: {df_finale['regione'].nunique()}")
    critici = (df_finale["stato"] == "CRITICO").sum()
    log.info(f"   Record critici: {critici}")

if __name__ == "__main__":
    main()
