import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.title("Krank-Meldungen Monatsübersicht (ab 2025)")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

german_months = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

wochentage = [
    "Montag", "Dienstag", "Mittwoch", "Donnerstag",
    "Freitag", "Samstag", "Sonntag"
]

if uploaded_files:
    eintraege = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[4:]
            df.columns = range(df.shape[1])

            for _, row in df.iterrows():
                kommentar = str(row[15]) if 15 in row and pd.notnull(row[15]) else ""
                name = row[3] if 3 in row else None
                vorname = row[4] if 4 in row else None
                datum = pd.to_datetime(row[14], errors='coerce') if 14 in row else None

                # NUR ab 2025
                if (
                    "krank" in kommentar.lower()
                    and pd.notnull(name)
                    and pd.notnull(vorname)
                    and pd.notnull(datum)
                    and datum.year >= 2025
                ):
                    kw = datum.isocalendar().week
                    wochentag = wochentage[datum.weekday()]  # 0=Montag
                    datum_kw = datum.strftime("%d.%m.%Y") + f" (KW {kw}, {wochentag})"
                    monat_index = datum.month
                    jahr = datum.year
                    monat_name = german_months[monat_index]
                    eintraege.append({
                        "Nachname": name,
                        "Vorname": vorname,
                        "DatumKW": datum_kw,
                        "Kommentar": kommentar,
                        "Monat": f"{monat_index:02d}-{jahr}_{monat_name} {jahr}"
                    })

        except Exception as e:
            st.error(f"Fehler in Datei {file.name}: {e}")

    if eintraege:
        df_gesamt = pd.DataFrame(eintraege)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for monat_key in sorted(df_gesamt["Monat"].unique()):
                df_monat = df_gesamt[df_gesamt["Monat"] == monat_key]
                zeilen = []
                for (nach, vor), gruppe in df_monat.groupby(["Nachname", "Vorname"]):
                    zeilen.append([f"{vor} {nach}", ""])
                    zeilen.append(["Datum", "Kommentar"])
                    for _, r in gruppe.iterrows():
                        zeilen.append([r["DatumKW"], r["Kommentar"]])
                    zeilen.append(["", ""])

                df_sheet = pd.DataFrame(zeilen, columns=["Spalte A", "Spalte B"])
                sheet_name = monat_key.split("_")[1][:31]
                df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)

                sheet = writer.sheets[sheet_name]
                sheet.row_dimensions[1].hidden = True  # Zeile 1 ausblenden

                thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                              top=Side(style='thin'), bottom=Side(style='thin'))

                orange_fill = PatternFill("solid", fgColor="ffc000")
                header_fill = PatternFill("solid", fgColor="95b3d7")

                for row in sheet.iter_rows():
                    val = str(row[0].value).strip().lower() if row[0].value else ""
                    is_name_row = (
                        str(row[0].value).strip() != ""
                        and (row[1].value is None or row[1].value == "")
                    )

                    for cell in row:
                        cell.font = Font(name="Calibri", size=11)
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        cell.border = thin

                        if is_name_row:
                            cell.font = Font(bold=True, size=12)
                            cell.fill = orange_fill
                        elif row[0].value == "Datum":
                            cell.font = Font(bold=True)
                            cell.fill = header_fill

                # Autobreite auf alle Spalten
                for col_cells in sheet.columns:
                    max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col_cells)
                    col_letter = get_column_letter(col_cells[0].column)
                    sheet.column_dimensions[col_letter].width = int(max_len * 1.2) + 2

        st.download_button("Excel-Datei herunterladen", output.getvalue(), file_name="Krank_Monatsauswertung.xlsx")

    else:
        st.warning("Keine gültigen Krank-Meldungen ab 2025 gefunden.")
