#!/usr/bin/env python3
"""
Wertet die Seite "Detaillierter Ladenachweis" von WWZ E-Mobilitätsrechnungen
(PDF) aus und exportiert die Tabelle als CSV und Excel. Es können ein
einzelnes PDF, mehrere PDFs oder ein ganzer Ordner voller PDFs auf einmal
verarbeitet werden - die Ergebnisse landen in einer gemeinsamen Tabelle.

Funktionsweise:
Der Text in diesem PDF-Typ ist nicht sauber extrahierbar (die Schrift ist
so eingebettet, dass normale Textextraktion nur Zeichensalat liefert).
Darum wird die gewünschte Seite zuerst als Bild gerendert und per OCR
(Tesseract) gelesen. Aus dem erkannten Text werden dann die einzelnen
Ladevorgänge per Muster-Erkennung (Regex) herausgefiltert.

Voraussetzungen (einmalig installieren):
    pip install pandas openpyxl --break-system-packages
    apt-get install poppler-utils tesseract-ocr tesseract-ocr-deu

Verwendung:
    # Einzelne Datei
    python3 wwz_ladenachweis_auswerten.py rechnung.pdf --seite 3

    # Mehrere Dateien
    python3 wwz_ladenachweis_auswerten.py jan.pdf feb.pdf maerz.pdf --out auswertung

    # Ganzer Ordner (alle *.pdf darin)
    python3 wwz_ladenachweis_auswerten.py ./rechnungen/ --out auswertung
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path

import pandas as pd


ROW_RE = re.compile(
    r"(\d{2}\.\d{2}\.\d{2})\s+(\d{2}:\d{2}:\d{2})\s+"
    r"(\d{2}\.\d{2}\.\d{2})\s+(\d{2}:\d{2}:\d{2})\s+"
    r"([\d.]+)\s+(.*?)\s+([\d.]+)\s*$"
)
LADESCHLUESSEL_RE = re.compile(r"Ladeschl(?:ü|ue?)ssel", re.IGNORECASE)
# Fall A: Ladeschlüssel-ID steht direkt hinter dem Label auf derselben Zeile
LABEL_AND_ID_RE = re.compile(r"Ladeschl(?:ü|ue?)ssel\s+(\S+)", re.IGNORECASE)
# Fall B: offizielles WWZ-Format "CH-XXX-..." irgendwo im (Folge-)Text
KEY_ID_RE = re.compile(r"CH-[A-Z0-9\-]+")
RECHNUNGSNR_RE = re.compile(r"Rechnung\s*Nr\.?\s*(\d+)", re.IGNORECASE)
STOP_WORDS = ("zwischentotal", "total")

# --- Muster für Seite "Kostendetails" ---
BEZUG_LADESTROM_RE = re.compile(r"Bezug Ladestrom.*?([\d.]+)\s*$", re.MULTILINE | re.IGNORECASE)
BETRIEBSDIENST_RE = re.compile(r"Betriebsdienstleistung.*?([\d.]+)\s*$", re.MULTILINE | re.IGNORECASE)
TOTAL_EXKL_RE = re.compile(r"Total exkl\.?\s*MwSt\.?\s+([\d.]+)", re.IGNORECASE)
MWST_RE = re.compile(r"Mehrwertsteuer\s+([\d,]+)\s*%\s+([\d.]+)", re.IGNORECASE)
TOTAL_INKL_RE = re.compile(r"Total inkl\.?\s*MwSt\.?\s+([\d.]+)", re.IGNORECASE)


def rasterize_page(pdf_path: Path, page: int, dpi: int, out_prefix: Path) -> Path:
    """Rendert eine PDF-Seite als JPEG mit pdftoppm."""
    subprocess.run(
        [
            "pdftoppm", "-jpeg", "-r", str(dpi),
            "-f", str(page), "-l", str(page),
            str(pdf_path), str(out_prefix),
        ],
        check=True,
    )
    matches = sorted(out_prefix.parent.glob(f"{out_prefix.name}-*.jpg"))
    if not matches:
        raise FileNotFoundError("Seite konnte nicht gerendert werden.")
    return matches[0]


def ocr_image(image_path: Path, lang: str = "deu") -> str:
    """Liest den Text eines Bildes per Tesseract-OCR."""
    result = subprocess.run(
        ["tesseract", str(image_path), "-", "--psm", "6", "-l", lang],
        check=True, capture_output=True, text=True,
    )
    return result.stdout


def parse_ladenachweis(text: str) -> pd.DataFrame:
    """Extrahiert die Ladevorgänge aus dem OCR-Text als Tabelle."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    rechnungsnr_match = RECHNUNGSNR_RE.search(text)
    rechnungsnr = rechnungsnr_match.group(1) if rechnungsnr_match else None

    records = []
    current_key = None

    for i, line in enumerate(lines):
        m = ROW_RE.search(line)
        if not m:
            continue

        start_datum, start_zeit, ende_datum, ende_zeit, verbrauch, adresse, betrag = m.groups()
        prefix = line[:m.start()]

        # Fall A: "Ladeschlüssel <ID>" steht direkt vor den Datenspalten dieser Zeile
        label_match = LABEL_AND_ID_RE.search(prefix)
        if label_match:
            current_key = label_match.group(1)

        # Zeilen ohne Datenzeilen-Muster, die danach folgen, an die Adresse anhängen
        # (Adresse und Ladeschlüssel-ID stehen bei diesem Layout oft in der Folgezeile)
        extra = ""
        j = i + 1
        while j < len(lines) and not ROW_RE.search(lines[j]) and not lines[j].lower().startswith(STOP_WORDS):
            extra += " " + lines[j]
            j += 1

        full_adresse = (adresse + extra).strip()

        # Fall B: ID im offiziellen WWZ-Format, meist in der Folgezeile bei der Adresse
        if not label_match:
            key_match = KEY_ID_RE.search(full_adresse)
            if key_match:
                current_key = key_match.group()
                full_adresse = full_adresse.replace(current_key, "").strip()
        full_adresse = re.sub(r"\s{2,}", " ", full_adresse)

        records.append({
            "Rechnungsnummer": rechnungsnr,
            "Ladeschlüssel": current_key,
            "Start Datum": start_datum,
            "Start Zeit UTC": start_zeit,
            "Ende Datum": ende_datum,
            "Ende Zeit UTC": ende_zeit,
            "Verbrauch kWh": float(verbrauch),
            "Adresse": full_adresse,
            "Betrag CHF exkl. MwSt.": float(betrag),
        })

    return pd.DataFrame(records)


def parse_kostendetails(text: str) -> dict:
    """Extrahiert Bezug Ladestrom, Betriebsdienstleistung, MwSt und Totale von der 'Kostendetails'-Seite."""
    def _num(m, group=1):
        return float(m.group(group).replace(",", ".")) if m else None

    bezug = BEZUG_LADESTROM_RE.search(text)
    betrieb = BETRIEBSDIENST_RE.search(text)
    total_exkl = TOTAL_EXKL_RE.search(text)
    mwst = MWST_RE.search(text)
    total_inkl = TOTAL_INKL_RE.search(text)

    return {
        "Bezug Ladestrom CHF": _num(bezug),
        "Betriebsdienstleistung CHF": _num(betrieb),
        "Total exkl. MwSt. CHF": _num(total_exkl),
        "MwSt-Satz %": _num(mwst, 1),
        "MwSt CHF": _num(mwst, 2),
        "Total inkl. MwSt. CHF": _num(total_inkl),
    }


def allocate_costs(detail_df: pd.DataFrame, kosten: dict) -> pd.DataFrame:
    """Verteilt Betriebsdienstleistung und MwSt proportional zum Energieanteil auf die Ladeschlüssel."""
    grouped = (
        detail_df.groupby("Ladeschlüssel", dropna=False)["Betrag CHF exkl. MwSt."]
        .sum()
        .rename("Energie CHF exkl. MwSt.")
        .reset_index()
    )
    total_energie = grouped["Energie CHF exkl. MwSt."].sum()
    betrieb_total = kosten.get("Betriebsdienstleistung CHF") or 0.0
    mwst_satz = (kosten.get("MwSt-Satz %") or 0.0) / 100

    if total_energie:
        grouped["Anteil"] = grouped["Energie CHF exkl. MwSt."] / total_energie
    else:
        grouped["Anteil"] = 0.0

    grouped["Betriebsdienstleistung CHF (verteilt)"] = (grouped["Anteil"] * betrieb_total).round(2)
    grouped["Zwischensumme exkl. MwSt. CHF"] = (
        grouped["Energie CHF exkl. MwSt."] + grouped["Betriebsdienstleistung CHF (verteilt)"]
    ).round(2)
    grouped["MwSt CHF (verteilt)"] = (grouped["Zwischensumme exkl. MwSt. CHF"] * mwst_satz).round(2)
    grouped["Total inkl. MwSt. CHF"] = (
        grouped["Zwischensumme exkl. MwSt. CHF"] + grouped["MwSt CHF (verteilt)"]
    ).round(2)

    grouped = grouped.drop(columns=["Anteil"])
    return grouped


def process_pdf(pdf_path: Path, seite: int, dpi: int, lang: str, tmp_dir: Path,
                 seite_kosten: int | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Rastert die Ladenachweis- (und optional Kostendetails-)Seite, macht OCR und liefert
    (Detailtabelle, Kostenverteilung pro Ladeschlüssel) zurück."""
    tmp_prefix = tmp_dir / f"_{pdf_path.stem}_page{seite}"
    image_path = rasterize_page(pdf_path, seite, dpi, tmp_prefix)
    try:
        text = ocr_image(image_path, lang)
    finally:
        image_path.unlink(missing_ok=True)

    df = parse_ladenachweis(text)
    df.insert(0, "Quelldatei", pdf_path.name)

    summary = pd.DataFrame()
    if seite_kosten and not df.empty:
        tmp_prefix_k = tmp_dir / f"_{pdf_path.stem}_page{seite_kosten}"
        image_path_k = rasterize_page(pdf_path, seite_kosten, dpi, tmp_prefix_k)
        try:
            text_k = ocr_image(image_path_k, lang)
        finally:
            image_path_k.unlink(missing_ok=True)
        kosten = parse_kostendetails(text_k)
        summary = allocate_costs(df, kosten)
        summary.insert(0, "Quelldatei", pdf_path.name)
        rechnungsnr = df["Rechnungsnummer"].iloc[0] if "Rechnungsnummer" in df else None
        summary.insert(1, "Rechnungsnummer", rechnungsnr)

    return df, summary


def collect_pdfs(paths: list[Path]) -> list[Path]:
    """Löst Ordner zu allen enthaltenen *.pdf-Dateien auf; Dateien bleiben wie sie sind."""
    pdfs = []
    for p in paths:
        if p.is_dir():
            pdfs.extend(sorted(p.glob("*.pdf")))
        else:
            pdfs.append(p)
    return pdfs


def main():
    parser = argparse.ArgumentParser(description="WWZ-Ladenachweis aus einem oder mehreren PDFs auswerten")
    parser.add_argument(
        "pdfs", type=Path, nargs="+",
        help="Ein oder mehrere PDF-Pfade, oder ein Ordner voller PDFs",
    )
    parser.add_argument("--seite", type=int, default=3, help="Seitenzahl mit dem Ladenachweis (Standard: 3)")
    parser.add_argument("--seite-kosten", type=int, default=2,
                         help="Seitenzahl mit den Kostendetails, für die Verteilung von Betriebsdienstleistung/MwSt "
                              "auf die Ladeschlüssel (Standard: 2). Mit 0 deaktivieren.")
    parser.add_argument("--dpi", type=int, default=200, help="Auflösung für die Rasterisierung (Standard: 200)")
    parser.add_argument("--lang", type=str, default="deu", help="OCR-Sprache (Standard: deu)")
    parser.add_argument("--out", type=Path, default=Path("ladenachweis"), help="Basisname für die Export-Dateien (ohne Endung)")
    args = parser.parse_args()

    pdf_files = collect_pdfs(args.pdfs)
    if not pdf_files:
        sys.exit("Keine PDF-Dateien gefunden.")

    seite_kosten = args.seite_kosten if args.seite_kosten > 0 else None

    all_dfs = []
    all_summaries = []
    for pdf_path in pdf_files:
        if not pdf_path.exists():
            print(f"  Übersprungen (nicht gefunden): {pdf_path}", file=sys.stderr)
            continue
        print(f"Verarbeite {pdf_path.name} ...")
        try:
            df, summary = process_pdf(pdf_path, args.seite, args.dpi, args.lang,
                                       args.out.parent or Path("."), seite_kosten)
        except Exception as exc:
            print(f"  Fehler bei {pdf_path.name}: {exc}", file=sys.stderr)
            continue
        if df.empty:
            print(f"  Keine Ladevorgänge gefunden in {pdf_path.name} (Seite {args.seite} prüfen).", file=sys.stderr)
            continue
        all_dfs.append(df)
        if not summary.empty:
            all_summaries.append(summary)

    if not all_dfs:
        sys.exit("Aus keiner der PDFs konnten Ladevorgänge extrahiert werden.")

    combined = pd.concat(all_dfs, ignore_index=True)

    csv_path = args.out.with_name(args.out.name + "_ladenachweis.csv")
    xlsx_path = args.out.with_name(args.out.name + "_ladenachweis.xlsx")
    combined.to_csv(csv_path, index=False)

    combined_summary = pd.concat(all_summaries, ignore_index=True) if all_summaries else None
    if combined_summary is not None:
        summary_csv_path = args.out.with_name(args.out.name + "_kosten_pro_schluessel.csv")
        combined_summary.to_csv(summary_csv_path, index=False)
        with pd.ExcelWriter(xlsx_path) as writer:
            combined.to_excel(writer, sheet_name="Ladenachweis", index=False)
            combined_summary.to_excel(writer, sheet_name="Kosten pro Schlüssel", index=False)
    else:
        combined.to_excel(xlsx_path, index=False)

    print("\n" + combined.to_string(index=False))
    print(f"\n{len(pdf_files)} PDF(s) verarbeitet, {len(combined)} Ladevorgänge total.")

    if combined_summary is not None:
        print("\nKostenverteilung pro Ladeschlüssel (Betriebsdienstleistung/MwSt proportional zum Energieanteil):")
        print(combined_summary.to_string(index=False))

    print(f"\nGespeichert als:\n  {csv_path}\n  {xlsx_path}")
    if combined_summary is not None:
        print(f"  {summary_csv_path}")


if __name__ == "__main__":
    main()
