import json
import re
import time
from datetime import datetime, date
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://punainenristi.fi"
INPUT_FILE = "osastojen_QA_tarkistuslista.xlsx"

DATE_PATTERNS = [
    r"\b(\d{1,2}\.\d{1,2}\.\d{4})\b",
    r"\b(\d{4}-\d{2}-\d{2})\b",
    r"\b(\d{1,2}/\d{1,2}/\d{4})\b",
]

OLD_YEAR_LIMIT = 2023

NEWS_HINTS = [
    "ajankohtaista",
    "uutiset",
    "news",
    "aktuellt",
    "nyheter",
]

CONTACT_HINTS = [
    "yhteystiedot",
    "contact",
    "contacts",
    "kontakt",
    "kontaktuppgifter",
]

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "SPR-osastotarkistus/1.0"
})


def find_first_existing_column(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col
    raise ValueError(f"Yhtään sarakkeista ei löytynyt: {candidates}")


def siisti_polku(arvo):
    if pd.isna(arvo):
        return ""
    polku = str(arvo).strip()
    if not polku:
        return ""
    return polku.split()[0]


def bool_to_k_e(value):
    return "K" if value else "E"


def medal_from_score(score):
    if score >= 8:
        return "Kulta"
    if score >= 6:
        return "Hopea"
    if score >= 4:
        return "Pronssi"
    return "Kehitettävä"


def parse_date_string(text):
    text = text.strip()
    formats = ["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def extract_dates_from_text(text):
    found_dates = []
    for pattern in DATE_PATTERNS:
        matches = re.findall(pattern, text)
        for match in matches:
            parsed = parse_date_string(match)
            if parsed:
                found_dates.append(parsed)
    return sorted(set(found_dates))


def extract_meta_modified_date(soup):
    meta_candidates = [
        ("property", "article:modified_time"),
        ("property", "article:published_time"),
        ("name", "last-modified"),
        ("name", "modified"),
        ("name", "date"),
    ]
    for attr_name, attr_value in meta_candidates:
        tag = soup.find("meta", attrs={attr_name: attr_value})
        if tag and tag.get("content"):
            content = tag["content"].strip()
            try:
                return datetime.fromisoformat(content.replace("Z", "+00:00")).date()
            except ValueError:
                pass
            for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S%z"):
                try:
                    return datetime.strptime(content, fmt).date()
                except ValueError:
                    pass
    return None


def extract_latest_date(soup):
    text = soup.get_text(" ", strip=True)
    meta_date = extract_meta_modified_date(soup)
    text_dates = extract_dates_from_text(text)
    latest_text_date = max(text_dates) if text_dates else None
    if meta_date and latest_text_date:
        return max(meta_date, latest_text_date)
    return meta_date or latest_text_date


def detect_old_years(text):
    years = re.findall(r"\b(19\d{2}|20\d{2})\b", text)
    years = sorted(set(int(y) for y in years))
    return [y for y in years if y <= OLD_YEAR_LIMIT]


def days_since(d):
    if not d:
        return ""
    return (date.today() - d).days


def format_date(d):
    return d.isoformat() if d else ""


def safe_get(url):
    try:
        return SESSION.get(url, timeout=12)
    except Exception:
        return None


def is_same_domain(url):
    try:
        parsed = urlparse(url)
        return parsed.netloc in {"punainenristi.fi", "www.punainenristi.fi"}
    except Exception:
        return False


def find_first_matching_link(soup, base_url, hints):
    for a in soup.find_all("a", href=True):
        href = (a.get("href") or "").strip()
        text = a.get_text(" ", strip=True).lower()
        href_lower = href.lower()

        if any(hint in text or hint in href_lower for hint in hints):
            full_url = urljoin(base_url, href)
            if is_same_domain(full_url):
                return full_url
    return ""


def count_news_items(soup):
    count = 0
    articles = soup.find_all("article")
    count += len(articles)

    news_like_classes = soup.find_all(
        class_=lambda c: c and any(x in str(c).lower() for x in ["news", "article", "post"])
    )
    count = max(count, len(news_like_classes))

    news_headings = soup.find_all(
        string=lambda s: s and any(x in s.lower() for x in ["ajankohtaista", "uutiset", "news", "aktuellt", "nyheter"])
    )
    if news_headings and count == 0:
        count = 1

    return count


def count_activity_boxes(soup):
    selectors = [
        lambda c: c and "activity" in str(c).lower(),
        lambda c: c and "toiminta" in str(c).lower(),
        lambda c: c and "card" in str(c).lower(),
        lambda c: c and "box" in str(c).lower(),
    ]

    counts = []
    for selector in selectors:
        matches = soup.find_all(class_=selector)
        counts.append(len(matches))

    texts = soup.find_all(
        string=lambda s: s and any(x in s.lower() for x in ["toiminta", "verksamhet", "activities"])
    )
    if texts:
        counts.append(1)

    return max(counts) if counts else 0


def fetch_page_details(url):
    result = {
        "url": url,
        "status_code": "",
        "ok": False,
        "latest_date": "",
        "age_days": "",
        "old_years": "",
        "error": "",
    }

    if not url:
        return result

    r = safe_get(url)
    if r is None:
        result["error"] = "HTTP request failed"
        return result

    result["status_code"] = r.status_code
    result["ok"] = (r.status_code == 200)

    if not result["ok"]:
        return result

    soup = BeautifulSoup(r.text, "html.parser")
    latest = extract_latest_date(soup)
    text = soup.get_text(" ", strip=True)
    old_years = detect_old_years(text)

    result["latest_date"] = format_date(latest)
    result["age_days"] = days_since(latest)
    result["old_years"] = ", ".join(str(y) for y in old_years[:10])

    return result


def laske_pisteet(tulos):
    pisteet = 0
    if tulos["linkki_toimii"]:
        pisteet += 1
    if tulos["yhteystiedot_loytyi_automaattisesti"]:
        pisteet += 1
    if tulos["uutisia_loytyi_automaattisesti"]:
        pisteet += 1
    if tulos["toimintoja_loytyi_automaattisesti"]:
        pisteet += 1
    if tulos["etusivun_paivitys_loytyi"]:
        pisteet += 1
    if tulos["uutissivun_paivitys_loytyi"]:
        pisteet += 1
    if tulos["yhteystietosivun_paivitys_loytyi"]:
        pisteet += 1
    if tulos["uutisten_lukumaara"] >= 3:
        pisteet += 1
    if tulos["sinisten_laatikoiden_maara"] >= 3:
        pisteet += 1
    return pisteet


def paata_tila(tulos):
    if not tulos["linkki_toimii"]:
        return "Virhe"
    if tulos["pisteet"] >= 7:
        return "OK"
    return "Puutteellinen"


def tarkista_osasto(url_polku):
    polku = siisti_polku(url_polku)
    url = urljoin(BASE + "/", polku.lstrip("/"))

    tulos = {
        "kaytetty_polku": polku,
        "osasto_url": url,
        "status_koodi": "",
        "linkki_toimii": False,
        "yhteystiedot_loytyi_automaattisesti": False,
        "uutisia_loytyi_automaattisesti": False,
        "toimintoja_loytyi_automaattisesti": False,
        "uutisten_lukumaara": 0,
        "sinisten_laatikoiden_maara": 0,
        "etusivu_url": url,
        "etusivun_viimeisin_paivitys": "",
        "etusivun_paivitys_ika_paivina": "",
        "etusivun_paivitys_loytyi": False,
        "uutissivu_url": "",
        "uutissivun_viimeisin_paivitys": "",
        "uutissivun_paivitys_ika_paivina": "",
        "uutissivun_paivitys_loytyi": False,
        "yhteystietosivu_url": "",
        "yhteystietosivun_viimeisin_paivitys": "",
        "yhteystietosivun_paivitys_ika_paivina": "",
        "yhteystietosivun_paivitys_loytyi": False,
        "sisalto_vaikuttaa_vanhalta": False,
        "vanhat_vuodet": "",
        "virhe": "",
    }

    if not polku:
        tulos["virhe"] = "Markkinointiosoite puuttuu"
        return tulos

    r = safe_get(url)
    if r is None:
        tulos["virhe"] = "HTTP request failed"
        return tulos

    tulos["status_koodi"] = r.status_code
    tulos["linkki_toimii"] = (r.status_code == 200)

    if not tulos["linkki_toimii"]:
        return tulos

    soup = BeautifulSoup(r.text, "html.parser")
    text = soup.get_text(" ", strip=True)

    tulos["yhteystiedot_loytyi_automaattisesti"] = bool(
        soup.find("a", href=lambda h: h and "mailto:" in h)
        or soup.find("a", href=lambda h: h and "tel:" in h)
    )

    tulos["uutisia_loytyi_automaattisesti"] = bool(
        soup.find("article")
        or soup.find(class_=lambda c: c and "news" in str(c).lower())
        or soup.find(string=lambda s: s and "ajankohtaista" in s.lower())
        or soup.find(string=lambda s: s and "aktuellt" in s.lower())
        or soup.find(string=lambda s: s and "news" in s.lower())
    )

    tulos["toimintoja_loytyi_automaattisesti"] = bool(
        soup.find(class_=lambda c: c and "activity" in str(c).lower())
        or soup.find(string=lambda s: s and "toiminta" in s.lower())
        or soup.find(string=lambda s: s and "verksamhet" in s.lower())
        or soup.find(string=lambda s: s and "activities" in s.lower())
    )

    tulos["uutisten_lukumaara"] = count_news_items(soup)
    tulos["sinisten_laatikoiden_maara"] = count_activity_boxes(soup)

    etusivu_latest = extract_latest_date(soup)
    tulos["etusivun_viimeisin_paivitys"] = format_date(etusivu_latest)
    tulos["etusivun_paivitys_ika_paivina"] = days_since(etusivu_latest)
    tulos["etusivun_paivitys_loytyi"] = etusivu_latest is not None

    old_years = detect_old_years(text)
    tulos["vanhat_vuodet"] = ", ".join(str(y) for y in old_years[:10])
    tulos["sisalto_vaikuttaa_vanhalta"] = bool(old_years)

    uutissivu_url = find_first_matching_link(soup, url, NEWS_HINTS)
    tulos["uutissivu_url"] = uutissivu_url
    if uutissivu_url:
        news_info = fetch_page_details(uutissivu_url)
        tulos["uutissivun_viimeisin_paivitys"] = news_info["latest_date"]
        tulos["uutissivun_paivitys_ika_paivina"] = news_info["age_days"]
        tulos["uutissivun_paivitys_loytyi"] = bool(news_info["latest_date"])

    yhteystietosivu_url = find_first_matching_link(soup, url, CONTACT_HINTS)
    tulos["yhteystietosivu_url"] = yhteystietosivu_url
    if yhteystietosivu_url:
        contact_info = fetch_page_details(yhteystietosivu_url)
        tulos["yhteystietosivun_viimeisin_paivitys"] = contact_info["latest_date"]
        tulos["yhteystietosivun_paivitys_ika_paivina"] = contact_info["age_days"]
        tulos["yhteystietosivun_paivitys_loytyi"] = bool(contact_info["latest_date"])

    return tulos


def rakenna_leaderboard(raportti):
    leaderboard = raportti.copy()
    leaderboard["sijoitusperuste"] = list(
        zip(
            -leaderboard["pisteet"],
            -leaderboard["linkki_toimii"].astype(int),
            -leaderboard["uutisten_lukumaara"].fillna(0).astype(int),
            -leaderboard["sinisten_laatikoiden_maara"].fillna(0).astype(int),
            -leaderboard["etusivun_paivitys_loytyi"].astype(int),
            -leaderboard["uutissivun_paivitys_loytyi"].astype(int),
            -leaderboard["yhteystietosivun_paivitys_loytyi"].astype(int),
            leaderboard["Osaston nimi"].astype(str),
        )
    )
    leaderboard = leaderboard.sort_values("sijoitusperuste").drop(columns=["sijoitusperuste"])
    leaderboard.insert(0, "sijoitus", range(1, len(leaderboard) + 1))
    leaderboard["score_text"] = leaderboard["pisteet"].astype(str) + "/9"
    return leaderboard


def update_history(raportti, district_col):
    dashboard_dir = Path("dashboard")
    dashboard_dir.mkdir(exist_ok=True)
    history_path = dashboard_dir / "history.json"

    run_date = datetime.utcnow().strftime("%Y-%m-%d")

    history_rows = []
    for _, row in raportti.iterrows():
        history_rows.append({
            "run_date": run_date,
            "piiri": row.get(district_col, ""),
            "osasto": row.get("Osaston nimi", ""),
            "pisteet": int(row.get("pisteet", 0)),
            "tila": row.get("tila", ""),
            "medal": row.get("medal", ""),
            "uutisten_lukumaara": int(row.get("uutisten_lukumaara", 0)),
            "sinisten_laatikoiden_maara": int(row.get("sinisten_laatikoiden_maara", 0)),
            "etusivun_paivitys_ika_paivina": row.get("etusivun_paivitys_ika_paivina", ""),
            "linkki_toimii": bool(row.get("linkki_toimii", False)),
        })

    existing = []
    if history_path.exists():
        try:
            existing = json.loads(history_path.read_text(encoding="utf-8"))
        except Exception:
            existing = []

    # Poista saman päivän duplikaatit ennen uuden snapshotin lisäämistä
    existing = [
        r for r in existing
        if not (r.get("run_date") == run_date and r.get("osasto") in set(raportti["Osaston nimi"].astype(str)))
    ]

    combined = existing + history_rows
    history_path.write_text(json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8")
    return combined


def build_summary(raportti, leaderboard, district_col):
    total = len(raportti)
    rikki = int((raportti["tila"] == "Virhe").sum())
    puutteellinen = int((raportti["tila"] == "Puutteellinen").sum())
    ok = int((raportti["tila"] == "OK").sum())
    vanhentuneet = int((pd.to_numeric(raportti["etusivun_paivitys_ika_paivina"], errors="coerce") > 365).sum())
    puuttuu_yhteystiedot = int((~raportti["yhteystiedot_loytyi_automaattisesti"]).sum())
    puuttuu_uutiset = int((~raportti["uutisia_loytyi_automaattisesti"]).sum())

    district_summary = (
        raportti.groupby(district_col, dropna=False)
        .agg(
            osastoja=("tila", "size"),
            ok=("tila", lambda s: int((s == "OK").sum())),
            puutteellinen=("tila", lambda s: int((s == "Puutteellinen").sum())),
            rikki=("tila", lambda s: int((s == "Virhe").sum())),
            keskipisteet=("pisteet", "mean"),
            kultaa=("medal", lambda s: int((s == "Kulta").sum())),
            hopeaa=("medal", lambda s: int((s == "Hopea").sum())),
            pronssia=("medal", lambda s: int((s == "Pronssi").sum())),
        )
        .reset_index()
        .fillna("")
    )
    district_summary["keskipisteet"] = district_summary["keskipisteet"].round(2)

    return {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "totals": {
            "tarkistettu": total,
            "rikki": rikki,
            "puutteellinen": puutteellinen,
            "ok": ok,
            "vanhentuneet": vanhentuneet,
            "puuttuu_yhteystiedot": puuttuu_yhteystiedot,
            "puuttuu_uutiset": puuttuu_uutiset,
        },
        "highlights": [
            "Tekninen toimivuus on pääosin hyvä, mutta sisältöjen ajantasaisuudessa ja aktiivisuudessa on eroja.",
            "Medal-luokitus auttaa näkemään nopeasti, mitkä osastot ansaitsevat kehuja ja missä tarvitaan tukea.",
        ],
        "top5": leaderboard.head(5)[["Osaston nimi", "score_text", "tila", "medal"]].to_dict(orient="records"),
        "bottom5": leaderboard.tail(5).iloc[::-1][["Osaston nimi", "score_text", "tila", "medal"]].to_dict(orient="records"),
        "districts": district_summary.to_dict(orient="records"),
    }


def build_trends(history_rows):
    if not history_rows:
        return {"district_trends": [], "osasto_trends": [], "medal_history": [], "latest_changes": []}

    hist = pd.DataFrame(history_rows)

    district_trends = (
        hist.groupby(["run_date", "piiri"], dropna=False)
        .agg(
            keskipisteet=("pisteet", "mean"),
            osastoja=("osasto", "size"),
            kultaa=("medal", lambda s: int((s == "Kulta").sum())),
            hopeaa=("medal", lambda s: int((s == "Hopea").sum())),
            pronssia=("medal", lambda s: int((s == "Pronssi").sum())),
        )
        .reset_index()
    )
    district_trends["keskipisteet"] = district_trends["keskipisteet"].round(2)

    osasto_trends = (
        hist.sort_values(["osasto", "run_date"])
        .groupby("osasto", dropna=False)
        .tail(12)
        .reset_index(drop=True)
    )

    medal_history = (
        hist.groupby(["run_date", "medal"], dropna=False)
        .size()
        .reset_index(name="count")
    )

    latest_dates = sorted(hist["run_date"].dropna().unique())
    latest_changes = []

    if len(latest_dates) >= 2:
        prev_date = latest_dates[-2]
        curr_date = latest_dates[-1]

        prev_df = hist[hist["run_date"] == prev_date][["osasto", "pisteet"]].rename(columns={"pisteet": "prev_pisteet"})
        curr_df = hist[hist["run_date"] == curr_date][["osasto", "pisteet", "medal", "piiri"]].rename(columns={"pisteet": "curr_pisteet"})

        merged = curr_df.merge(prev_df, on="osasto", how="left")
        merged["delta"] = merged["curr_pisteet"] - merged["prev_pisteet"].fillna(0)
        latest_changes = merged.sort_values("delta", ascending=False).to_dict(orient="records")

    return {
        "district_trends": district_trends.to_dict(orient="records"),
        "osasto_trends": osasto_trends.to_dict(orient="records"),
        "medal_history": medal_history.to_dict(orient="records"),
        "latest_changes": latest_changes,
    }


def write_dashboard_files(raportti, leaderboard, summary, trends, district_col):
    dashboard_dir = Path("dashboard")
    dashboard_dir.mkdir(exist_ok=True)

    export_cols = [
        district_col,
        "Osaston nimi",
        "osasto_url",
        "tila",
        "pisteet",
        "Score",
        "medal",
        "uutisten_lukumaara",
        "sinisten_laatikoiden_maara",
        "etusivun_paivitys_ika_paivina",
        "uutissivun_paivitys_ika_paivina",
        "yhteystietosivun_paivitys_ika_paivina",
        "linkki_toimii",
        "yhteystiedot_loytyi_automaattisesti",
        "uutisia_loytyi_automaattisesti",
        "toimintoja_loytyi_automaattisesti",
        "sisalto_vaikuttaa_vanhalta",
        "virhe",
    ]
    export_cols = [c for c in export_cols if c in raportti.columns]

    raportti[export_cols].to_json(
        dashboard_dir / "data.json",
        orient="records",
        force_ascii=False,
        indent=2,
    )

    leaderboard.head(100).to_json(
        dashboard_dir / "leaderboard.json",
        orient="records",
        force_ascii=False,
        indent=2,
    )

    with open(dashboard_dir / "summary.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    with open(dashboard_dir / "trends.json", "w", encoding="utf-8") as f:
        json.dump(trends, f, ensure_ascii=False, indent=2)


def main():
    print(f"Luetaan tiedosto: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)
    df.columns = df.columns.str.strip()

    name_col = find_first_existing_column(
        df,
        ["Osaston nimi", "Avdelning", "Avdelningens namn", "Branch Name", "Department Name", "Bransh Name"],
    )
    path_col = find_first_existing_column(
        df,
        ["Markkinointiosoite", "Avdelning URL", "Branch URL", "Department URL", "Bransh URL"],
    )
    district_col = find_first_existing_column(
        df,
        ["Piiri", "Distrikt", "District"],
    )

    if "Osaston nimi" not in df.columns:
        df["Osaston nimi"] = df[name_col]

    tarkistettavat = df.dropna(subset=[path_col]).copy()
    print(f"Tarkistettavia rivejä: {len(tarkistettavat)}")

    tulokset = []

    for _, rivi in tarkistettavat.iterrows():
        nimi = str(rivi[name_col]).strip()
        print(f"Tarkistetaan: {nimi}")

        tarkistus = tarkista_osasto(rivi[path_col])
        yhdistetty = rivi.to_dict()

        for k, v in tarkistus.items():
            yhdistetty[k] = v

        yhdistetty["pisteet"] = laske_pisteet(yhdistetty)
        yhdistetty["tila"] = paata_tila(yhdistetty)
        yhdistetty["medal"] = medal_from_score(yhdistetty["pisteet"])
        yhdistetty["Score"] = f'{yhdistetty["pisteet"]}/9'

        yhdistetty["Auto: Linkki toimii"] = bool_to_k_e(yhdistetty["linkki_toimii"])
        yhdistetty["Auto: Yhteystiedot OK"] = bool_to_k_e(yhdistetty["yhteystiedot_loytyi_automaattisesti"])
        yhdistetty["Auto: Ajankohtaisia uutisia"] = bool_to_k_e(yhdistetty["uutisia_loytyi_automaattisesti"])
        yhdistetty["Auto: Toimintoja sivulla"] = bool_to_k_e(yhdistetty["toimintoja_loytyi_automaattisesti"])

        tulokset.append(yhdistetty)
        time.sleep(0.3)

    raportti = pd.DataFrame(tulokset)

    rikki = raportti[raportti["tila"] == "Virhe"].copy()
    puutteellinen = raportti[raportti["tila"] == "Puutteellinen"].copy()
    ok = raportti[raportti["tila"] == "OK"].copy()

    vanhentuneet = raportti[
        (
            pd.to_numeric(raportti["etusivun_paivitys_ika_paivina"], errors="coerce") > 365
        ) | (
            pd.to_numeric(raportti["uutissivun_paivitys_ika_paivina"], errors="coerce") > 365
        ) | (
            pd.to_numeric(raportti["yhteystietosivun_paivitys_ika_paivina"], errors="coerce") > 365
        )
    ].copy()

    leaderboard = rakenna_leaderboard(raportti)
    history_rows = update_history(raportti, district_col)
    summary = build_summary(raportti, leaderboard, district_col)
    trends = build_trends(history_rows)
    write_dashboard_files(raportti, leaderboard, summary, trends, district_col)

    tiedostonimi = f"osasto_tarkistus_{datetime.now().strftime('%Y%m%d')}.xlsx"

    with pd.ExcelWriter(tiedostonimi, engine="openpyxl") as writer:
        raportti.to_excel(writer, sheet_name="Kaikki", index=False)
        leaderboard.to_excel(writer, sheet_name="Leaderboard", index=False)
        rikki.to_excel(writer, sheet_name="Rikki", index=False)
        puutteellinen.to_excel(writer, sheet_name="Puutteellinen", index=False)
        ok.to_excel(writer, sheet_name="OK", index=False)
        vanhentuneet.to_excel(writer, sheet_name="Vanhentuneet", index=False)

    print("")
    print(f"Valmis! Raportti tallennettu tiedostoon: {tiedostonimi}")
    print(f"Kaikki rivejä: {len(raportti)}")
    print(f"Rikki: {len(rikki)}")
    print(f"Puutteellinen: {len(puutteellinen)}")
    print(f"OK: {len(ok)}")
    print(f"Vanhentuneet: {len(vanhentuneet)}")


if __name__ == "__main__":
    main()
