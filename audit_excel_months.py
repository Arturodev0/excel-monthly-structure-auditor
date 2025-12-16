from __future__ import annotations

import argparse
import re
from pathlib import Path
import pandas as pd

PL_SHEETS_TRY_DEFAULT = ["P&L", "P&L by Month"]
BS_SHEET_DEFAULT = "BS by Month Condensed"
DB_SHEET_DEFAULT = "DataBase Result"

COMBINED_FILENAME_DEFAULT = "combined.xlsx"
MONTH_FILE_DEFAULT = "monthly.xlsx"

PL_EXPECT = {"Parent", "Category"}
BS_EXPECT = {"Category"}
DB_EXPECT = {"Account"}

MONTH_FOLDER_RE = re.compile(r"^\d{1,2}\.\d{4}$")


def discover_months(base: Path):
    months = []
    for year_dir in sorted(base.iterdir()):
        if not year_dir.is_dir() or not year_dir.name.isdigit():
            continue
        for mdir in sorted(year_dir.iterdir()):
            if not mdir.is_dir() or not MONTH_FOLDER_RE.match(mdir.name):
                continue
            mm, yyyy = mdir.name.split(".")
            month = f"{int(mm):02d}"
            source = f"{year_dir.name}/{month}.{yyyy}"
            months.append((source, year_dir, mdir))
    return months


def infer_header_row(file_path: Path, sheet: str, expected_cols: set[str], scan_rows: int = 35):
    """
    Lee las primeras N filas sin header y trata de encontrar en qué fila vienen los headers.
    No modifica nada, solo detecta.
    """
    try:
        preview = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=scan_rows, engine="openpyxl")
    except Exception as e:
        return None, None, f"read_preview_failed: {type(e).__name__}: {e}"

    best_idx = None
    best_score = -1

    for i in range(len(preview)):
        row = preview.iloc[i].tolist()
        vals = set()
        for v in row:
            if pd.isna(v):
                continue
            s = str(v).strip()
            if s:
                vals.add(s)
        score = sum(1 for c in expected_cols if c in vals)
        if score > best_score:
            best_score = score
            best_idx = i

    if best_score <= 0:
        return None, None, "header_not_detected"

    return best_idx, best_score, None


def get_columns(file_path: Path, sheet: str, expected_cols: set[str]):
    """
    Intenta leer columnas de forma robusta:
    1) Detecta header row buscando columnas esperadas
    2) Lee nrows=0 con ese header para obtener nombres
    """
    header_row, score, err = infer_header_row(file_path, sheet, expected_cols)
    if err:
        try:
            df0 = pd.read_excel(file_path, sheet_name=sheet, nrows=0, engine="openpyxl")
            return list(df0.columns), 0, f"fallback_header0 ({err})"
        except Exception as e:
            return [], None, f"read_header_failed: {type(e).__name__}: {e}"

    try:
        df0 = pd.read_excel(file_path, sheet_name=sheet, nrows=0, header=header_row, engine="openpyxl")
        cols = [str(c).strip() for c in df0.columns]
        return cols, header_row, None
    except Exception as e:
        return [], header_row, f"read_with_detected_header_failed: {type(e).__name__}: {e}"


def read_sources_from_sheet(xlsx: Path, sheet: str, source_col: str = "Source"):
    try:
        df = pd.read_excel(xlsx, sheet_name=sheet, usecols=[source_col], engine="openpyxl")
        sources = set(df[source_col].dropna().astype(str).str.strip().unique())
        return sources, None
    except Exception as e:
        return set(), f"{type(e).__name__}: {e}"


def parse_args():
    p = argparse.ArgumentParser(prog="audit_excel_months.py")
    p.add_argument("--base-dir", default=None, help="Directorio base que contiene carpetas YYYY/MM.YYYY")
    p.add_argument("--combined", default=COMBINED_FILENAME_DEFAULT, help="Nombre del archivo combined")
    p.add_argument("--monthly", default=MONTH_FILE_DEFAULT, help="Nombre del archivo mensual dentro de cada MM.YYYY")
    p.add_argument("--pl-sheets", nargs="*", default=PL_SHEETS_TRY_DEFAULT, help="Nombres de hojas P&L a intentar")
    p.add_argument("--bs-sheet", default=BS_SHEET_DEFAULT, help="Nombre de hoja BS")
    p.add_argument("--db-sheet", default=DB_SHEET_DEFAULT, help="Nombre de hoja DB")
    p.add_argument("--combined-pl-sheet", default="P&L Combined", help="Nombre de hoja en combined para P&L")
    p.add_argument("--combined-bs-sheet", default="BS Condensed Combined", help="Nombre de hoja en combined para BS")
    p.add_argument("--combined-db-sheet", default="DataBase Combined", help="Nombre de hoja en combined para DB")
    p.add_argument("--source-col", default="Source", help="Nombre de columna Source en combined")
    return p.parse_args()


def main():
    args = parse_args()

    base = Path(args.base_dir).resolve() if args.base_dir else Path(__file__).resolve().parent
    combined = base / args.combined

    print("\n==============================")
    print("AUDIT Excel Monthly Structure (solo lectura)")
    print(f"Base: {base}")
    print("==============================\n")

    months = discover_months(base)
    if not months:
        print("❌ No se encontraron carpetas con patrón YYYY/MM.YYYY.")
        return

    folder_sources = [s for s, _, _ in months]
    folder_sources_set = set(folder_sources)

    print(f"📁 Meses detectados en carpetas: {len(folder_sources_set)}")
    print(f"   Ejemplo: {folder_sources[:5]}\n")

    problems = []
    ok_count = 0

    for source, year_dir, mdir in months:
        month_file = mdir / args.monthly
        if not month_file.exists():
            problems.append((source, "MISSING_FILE", str(month_file)))
            continue

        pl_sheet_used = None
        pl_cols = []
        pl_header = None
        pl_err = None

        for s_try in args.pl_sheets:
            cols, header, err = get_columns(month_file, s_try, PL_EXPECT)
            if cols:
                pl_sheet_used = s_try
                pl_cols = cols
                pl_header = header
                pl_err = err
                break

        if not pl_sheet_used:
            problems.append((source, "MISSING_PL_SHEET", f"tried={args.pl_sheets}"))
        else:
            missing_pl = sorted(list(PL_EXPECT - set(pl_cols)))
            if missing_pl:
                problems.append((source, "PL_MISSING_COLS", f"sheet={pl_sheet_used} header_row={pl_header} missing={missing_pl} note={pl_err}"))

        bs_cols, bs_header, bs_err = get_columns(month_file, args.bs_sheet, BS_EXPECT)
        if not bs_cols:
            problems.append((source, "MISSING_BS_SHEET", args.bs_sheet))
        else:
            missing_bs = sorted(list(BS_EXPECT - set(bs_cols)))
            if missing_bs:
                problems.append((source, "BS_MISSING_COLS", f"header_row={bs_header} missing={missing_bs} note={bs_err}"))

        db_cols, db_header, db_err = get_columns(month_file, args.db_sheet, DB_EXPECT)
        if not db_cols:
            problems.append((source, "MISSING_DB_SHEET", args.db_sheet))
        else:
            missing_db = sorted(list(DB_EXPECT - set(db_cols)))
            if missing_db:
                problems.append((source, "DB_MISSING_COLS", f"header_row={db_header} missing={missing_db} note={db_err}"))

        if pl_sheet_used and bs_cols and db_cols and not (PL_EXPECT - set(pl_cols)) and not (BS_EXPECT - set(bs_cols)) and not (DB_EXPECT - set(db_cols)):
            ok_count += 1

    print(f"✅ Meses OK (archivo + sheets + columnas mínimas): {ok_count}/{len(months)}\n")

    if problems:
        print("⚠️ Problemas detectados (esto suele explicar huecos en el dashboard):")
        for source, code, detail in problems[:50]:
            print(f" - {source} | {code} | {detail}")
        if len(problems) > 50:
            print(f" ... y {len(problems)-50} más\n")
        else:
            print()
    else:
        print("✅ No se detectaron problemas de estructura en los Excels mensuales.\n")

    print("---- AUDIT COMBINED ----")
    if not combined.exists():
        print(f"❌ No existe {combined}")
        return

    pl_sources, pl_err = read_sources_from_sheet(combined, args.combined_pl_sheet, args.source_col)
    bs_sources, bs_err = read_sources_from_sheet(combined, args.combined_bs_sheet, args.source_col)
    db_sources, db_err = read_sources_from_sheet(combined, args.combined_db_sheet, args.source_col)

    if pl_err:
        print(f"❌ No pude leer '{args.combined_pl_sheet}' Sources: {pl_err}")
    if bs_err:
        print(f"❌ No pude leer '{args.combined_bs_sheet}' Sources: {bs_err}")
    if db_err:
        print(f"❌ No pude leer '{args.combined_db_sheet}' Sources: {db_err}")

    if not pl_err:
        missing_in_pl = sorted(list(folder_sources_set - pl_sources))
        extra_in_pl = sorted(list(pl_sources - folder_sources_set))
        print(f"📄 {args.combined_pl_sheet} Sources: {len(pl_sources)}")
        if missing_in_pl:
            print(f"   ⚠️ Carpetas existentes pero faltan en P&L Combined: {missing_in_pl[:20]}{'...' if len(missing_in_pl)>20 else ''}")
        if extra_in_pl:
            print(f"   ⚠️ Sources en P&L Combined que no existen en carpetas: {extra_in_pl[:20]}{'...' if len(extra_in_pl)>20 else ''}")

    if not bs_err:
        missing_in_bs = sorted(list(folder_sources_set - bs_sources))
        print(f"📄 {args.combined_bs_sheet} Sources: {len(bs_sources)}")
        if missing_in_bs:
            print(f"   ⚠️ Carpetas existentes pero faltan en BS Combined: {missing_in_bs[:20]}{'...' if len(missing_in_bs)>20 else ''}")

    if not db_err:
        missing_in_db = sorted(list(folder_sources_set - db_sources))
        print(f"📄 {args.combined_db_sheet} Sources: {len(db_sources)}")
        if missing_in_db:
            print(f"   ⚠️ Carpetas existentes pero faltan en DB Combined: {missing_in_db[:20]}{'...' if len(missing_in_db)>20 else ''}")

    print("\n✅ AUDIT COMPLETADO.\n")
    print("Interpretación rápida:")
    print("- Si un mes falta en Combined (Sources), el dashboard tendrá huecos para ese periodo.")
    print("- Si el mes existe pero hay DB_MISSING_COLS (ej. falta Account), ese mes no aporta DataBase Result.")
    print("- Si no hay problemas aquí, los blancos suelen ser por medidas/filtros en Power BI (no por el Excel).")


if __name__ == "__main__":
    main()
