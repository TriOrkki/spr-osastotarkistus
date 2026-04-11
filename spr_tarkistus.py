import time
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://punainenristi.fi"
INPUT_FILE = "osastojen_QA_tarkistuslista.xlsx"


def siisti_polku(markkinointiosoite):
    if pd.isna(markkinointiosoite):
        return ""

    polku = str(markkinointiosoite).strip()
    if not polku:
        return ""

    # Jos samassa solussa on useampi polku, käytetään ensimmäistä
    return polku.split()[0]


def paata_tila(tulos):
    if not tulos["linkki_toimii"]:
        return "Virhe"

    puutteita = 0
    if not tulos["yhteystiedot"]:
        puutteita += 1
    if not tulos["uutisia"]:
        puutteita += 1
    if not tulos["toimintoja"]:
        puutteita += 1

    if puutteita == 0:
        return "OK"
    return "Puutteellinen"


def tarkista_osasto(url_polku):
    polku = siisti_polku(url_polku)
    url = urljoin(BASE + "/", polku.lstrip("/"))

    tulos = {
        "osasto_url": url,
        "kaytetty_polku": polku,
        "linkki_toimii": False,
        "status_koodi": "",
        "yhteystiedot": False,
        "uutisia": False,
        "toimintoja": False,
        "tila": "",
        "virhe": "",
    }

    if not polku:
        tulos["virhe"] = "Markkinointiosoite puuttuu"
        tulos["tila"] = "Virhe"
        return tulos

    try:
        r = requests.get(url, timeout=10)
        tulos["status_koodi"] = r.status_code
        tulos["linkki_toimii"] = (r.status_code == 200)

        if tulos["linkki_toimii"]:
            soup = BeautifulSoup(r.text, "html.parser")

            tulos["yhteystiedot"] = bool(
                soup.find("a", href=lambda h: h and "mailto:" in h)
                or soup.find("a", href=lambda h: h and "tel:" in h)
            )

            tulos["uutisia"] = bool(
                soup.find("article")
                or soup.find(class_=lambda c: c and "news" in str(c).lower())
                or soup.find(string=lambda s: s and "ajankohtaista" in s.lower())
            )

            tulos["toimintoja"] = bool(
                soup.find(class_=lambda c: c and "activity" in str(c).lower())
                or soup.find(string=lambda s: s and "toiminta" in s.lower())
            )

    except Exception as e:
        tulos["virhe"] = str(e)

    tulos["tila"] = paata_tila(tulos)
    return tulos


def main():
    print(f"Luetaan tiedosto: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)
    df.columns = df.columns.str.strip()

    print("Löydetyt sarakkeet:")
    print(list(df.columns))
    print(f"Rivejä tiedostossa: {len(df)}")

    pakolliset = ["Osaston nimi", "Markkinointiosoite"]
    puuttuvat = [sarake for sarake in pakolliset if sarake not in df.columns]
    if puuttuvat:
        raise ValueError(f"Excelistä puuttuvat sarakkeet: {puuttuvat}")

    osastot = df.dropna(subset=["Markkinointiosoite"]).copy()
    print(f"Tarkistettavia rivejä: {len(osastot)}")

    tulokset = []

    for _, rivi in osastot.iterrows():
        osaston_nimi = str(rivi["Osaston nimi"]).strip()
        alkuperainen_polku = rivi["Markkinointiosoite"]
        kaytettava_polku = siisti_polku(alkuperainen_polku)

        print(f"Tarkistetaan: {osaston_nimi}")
        print(f"  Polku: {kaytettava_polku}")

        tarkistus = tarkista_osasto(alkuperainen_polku)

        yhdistetty = rivi.to_dict()
        yhdistetty["kaytetty_polku"] = tarkistus["kaytetty_polku"]
        yhdistetty["osasto_url"] = tarkistus["osasto_url"]
        yhdistetty["tila"] = tarkistus["tila"]
        yhdistetty["status_koodi"] = tarkistus["status_koodi"]
        yhdistetty["linkki_toimii"] = tarkistus["linkki_toimii"]
        yhdistetty["yhteystiedot_löytyi_automaattisesti"] = tarkistus["yhteystiedot"]
        yhdistetty["uutisia_löytyi_automaattisesti"] = tarkistus["uutisia"]
        yhdistetty["toimintoja_löytyi_automaattisesti"] = tarkistus["toimintoja"]
        yhdistetty["virhe"] = tarkistus["virhe"]

        tulokset.append(yhdistetty)
        time.sleep(0.5)

    raportti = pd.DataFrame(tulokset)

    # Siistimpi sarakejärjestys: alkuperäiset sarakkeet ensin, sitten automaattiset tulokset
    alkuperaiset = list(df.columns)
    lisasarakkeet = [
        "kaytetty_polku",
        "osasto_url",
        "tila",
        "status_koodi",
        "linkki_toimii",
        "yhteystiedot_löytyi_automaattisesti",
        "uutisia_löytyi_automaattisesti",
        "toimintoja_löytyi_automaattisesti",
        "virhe",
    ]
    raportti = raportti[alkuperaiset + lisasarakkeet]

    rikki = raportti[raportti["tila"] == "Virhe"].copy()
    puutteellinen = raportti[raportti["tila"] == "Puutteellinen"].copy()
    ok = raportti[raportti["tila"] == "OK"].copy()

    tiedostonimi = f"tarkistus_{datetime.now().strftime('%Y%m%d')}.xlsx"

    with pd.ExcelWriter(tiedostonimi, engine="openpyxl") as writer:
        raportti.to_excel(writer, sheet_name="Kaikki", index=False)
        rikki.to_excel(writer, sheet_name="Rikki", index=False)
        puutteellinen.to_excel(writer, sheet_name="Puutteellinen", index=False)
        ok.to_excel(writer, sheet_name="OK", index=False)

    print("")
    print(f"Valmis! Raportti tallennettu tiedostoon: {tiedostonimi}")
    print(f"Kaikki rivejä: {len(raportti)}")
    print(f"Rikki: {len(rikki)}")
    print(f"Puutteellinen: {len(puutteellinen)}")
    print(f"OK: {len(ok)}")


if __name__ == "__main__":
    main()
