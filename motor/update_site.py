#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_site.py — вишесезонска верзија
---------------------------------------
Покреће се сваки пут када додаш нови пар Excel фајлова за неку сезону, или
измениш постојеће, да би се сајт (свих сезона) регенерисао.

КАКО ИМЕНОВАТИ EXCEL ФАЈЛОВЕ (скрипта их сама препознаје по имену):

  Termini 202526.xlsx        + Golovi 202526.xlsx        -> Сезона 2025/26
  Termini Leto 2026.xlsx     + Golovi Leto 2026.xlsx     -> Летња развојна лига 2026
  Termini 202627.xlsx        + Golovi 202627.xlsx        -> Сезона 2026/27
  ... и тако даље, без ограничења броја сезона.

  Правило: фајл мора почети са "Termini" тј. "Golovi" (величина слова није
  битна), а остатак имена (нпр. "202526" или "Leto 2026") мора бити ИСТИ за
  оба фајла истог пара - по томе их скрипта спарује.

Шта скрипта ради за сваку пронађену сезону:
  1. Учитава оба Excel фајла те сезоне.
  2. Чисти податке (уклања помоћне/празне колоне, редове без Датума).
  3. Уклања колоне које су ЦЕЛУ сезону 0 (нпр. ако летња лига не прати
     шутеве/додавања/одбране) - тако их сајт аутоматски сакрива, без икаквог
     ручног означавања "ово је летња сезона".
  4. Пише <sezona>/data/termini.js, golovi.js и config.js.
  5. Копира "мотор" сајта (season-template/index.html) у <sezona>/index.html.
     (Ово значи: ако желиш нешто да измениш на сајту - нпр. упишеш Google
     Analytics ID - измени ГА У season-template/index.html па поново
     покрени скрипту; она ће то пренети у све сезоне.)
  6. Ажурира seasons.js и корени index.html (почетна страна за избор сезоне)
     на основу свих пронађених сезона.

Употреба:
    python motor/update_site.py
        (покрени из корена сајта; скрипта тражи Excel фајлове у фолдеру
        "podaci/", поред "motor/")

Ниједна постојећа сезона се не брише сама од себе - скрипта само
додаје/освежава фолдере за парове фајлова које пронађе. Ако избациш неки
Excel из фолдера, стари генерисани фолдер те сезоне остаје нетакнут (само
seasons.js више неће показивати на њега, осим ако не постоји и даље на
диску - у ком случају линк и даље ради, само није излистан на почетној).
"""

import sys
import re
import json
import shutil
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Недостаје библиотека 'pandas'. Инсталирај је командом:")
    print("    pip install pandas openpyxl")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent

# Корен целог сајта - један ниво изнад фолдера "motor" у ком је ова скрипта
SITE_ROOT = SCRIPT_DIR.parent

# Где се траже Excel фајлови (фолдер "podaci", поред "motor")
EXCEL_DIR = SITE_ROOT / "podaci"

# "Мотор" сајта који се копира у сваку сезону - све измене на сајту
# (изглед, нове функције, GA4 ID) раде се ОВДЕ, на једном месту.
TEMPLATE_HTML = SCRIPT_DIR / "season-template" / "index.html"

# Шаблон за почетну страну (избор сезоне) - копира се у SITE_ROOT/index.html
ROOT_INDEX_TEMPLATE = SCRIPT_DIR / "root-index-template.html"

# Колоне које се НИКАД не уклањају, чак и ако су целу сезону 0/празне
TERMINI_PROTECTED_COLS = {"No.", "Date", "Team", "Player Name", "Points", "Minutes Played"}
GOLOVI_PROTECTED_COLS = {"No.", "Date", "Team", "Goalscorer", "Assist", "Minute", "Goalkeeper"}


# ---------------------------------------------------------------------------
# Проналажење и именовање сезона
# ---------------------------------------------------------------------------

def find_season_pairs():
    """Тражи све парове Termini <suffix>.xlsx / Golovi <suffix>.xlsx у EXCEL_DIR."""
    found = {}  # suffix -> {"termini": Path, "golovi": Path}
    for f in sorted(EXCEL_DIR.glob("*.xlsx")):
        name = f.stem
        m = re.match(r"(?i)^termini[\s_]*(.*)$", name)
        if m:
            suffix = m.group(1).strip()
            found.setdefault(suffix, {})["termini"] = f
            continue
        m = re.match(r"(?i)^golovi[\s_]*(.*)$", name)
        if m:
            suffix = m.group(1).strip()
            found.setdefault(suffix, {})["golovi"] = f

    complete = {}
    for suffix, files in found.items():
        if "termini" in files and "golovi" in files:
            complete[suffix] = files
        else:
            missing = "Golovi" if "termini" in files else "Termini"
            have = files.get("termini") or files.get("golovi")
            print(f"УПОЗОРЕЊЕ: '{have.name}' нема пар - недостаје одговарајући {missing} фајл. Прескачем ову сезону.")
    return complete


def suffix_to_season_meta(suffix: str):
    """Одређује id/label/icon сезоне на основу дела имена фајла после 'Termini'/'Golovi'."""
    s = suffix.strip()

    # летња/развојна лига: "Leto 2026", "Leto2026", "Summer 2026" ...
    m = re.match(r"(?i)^(leto|ljeto|summer)\s*[-_]?\s*(\d{4})$", s)
    if m:
        year = m.group(2)
        # конвенција: летња развојна лига увек траје 1. 6. - 31. 8. те године,
        # без обзира на то који су термини стварно одиграни
        return {
            "id": f"leto-{year}", "label": f"Летња развојна лига {year}", "icon": "☀️",
            "dateFrom": f"1. 6. {year}.", "dateTo": f"31. 8. {year}.",
        }

    # редовна сезона: "202526", "2025-26", "2025/26", "2025_26"
    m = re.match(r"^(\d{4})\s*[-/_]?\s*(\d{2})$", s)
    if m:
        y1, y2 = m.group(1), m.group(2)
        # конвенција: редовна сезона увек траје 1. 9. те године - 31. 5. следеће,
        # без обзира на то који су термини стварно одиграни
        return {
            "id": f"{y1}-{y2}", "label": f"Сезона {y1}/{y2}", "icon": "🏆",
            "dateFrom": f"1. 9. {y1}.", "dateTo": f"31. 5. {int(y1)+1}.",
        }

    # непрепознат формат - користи име фајла као ознаку, упозори корисника
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", s).strip("-").lower() or "sezona"
    print(f"УПОЗОРЕЊЕ: не препознајем формат назива сезоне '{s}'.")
    print(f"           Користим га као ознаку '{slug}' - препоручљиво је преименовати фајл")
    print(f"           на облик 'Termini 202627.xlsx' или 'Termini Leto 2027.xlsx'.")
    return {"id": slug, "label": s, "icon": "🏆"}


# ---------------------------------------------------------------------------
# Учитавање и чишћење Excel фајлова
# ---------------------------------------------------------------------------

def _format_date_series(series):
    """Датум у Excel-у је број облика 20250818 (YYYYMMDD). JS код на сајту
    очекује стринг '20250818', па га претварамо у string, попуњен нулама до
    8 карактера."""
    return series.astype("Int64").astype(str).str.zfill(8)


def _drop_helper_columns(df: pd.DataFrame) -> pd.DataFrame:
    helper_cols = [c for c in df.columns if str(c).startswith("Unnamed")]
    if helper_cols:
        df = df.drop(columns=helper_cols)
    return df


def load_termini(xlsx_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    df = _drop_helper_columns(df)
    df = df.dropna(subset=["No.", "Date"], how="any").copy()
    df["Date"] = _format_date_series(df["Date"])
    return df


def load_golovi(xlsx_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    df = _drop_helper_columns(df)
    df = df.dropna(subset=["No.", "Date"], how="any").copy()
    df["No."] = df["No."].astype("Int64")
    df["Date"] = _format_date_series(df["Date"])
    return df


def drop_allzero_columns(df: pd.DataFrame, protected_cols: set) -> pd.DataFrame:
    """Уклања бројчане колоне које су КРОЗ ЦЕЛУ сезону 0 или празне - то је
    знак да та статистика није праћена ове сезоне (нпр. летња лига нема
    шутеве/додавања/одбране). Пошто се ово ради по сезони посебно, свака
    сезона аутоматски "открива" шта јој недостаје, без ручног подешавања."""
    to_drop = []
    for col in df.columns:
        if col in protected_cols:
            continue
        if pd.api.types.is_numeric_dtype(df[col]):
            values = df[col].fillna(0)
            if (values == 0).all():
                to_drop.append(col)
    if to_drop:
        df = df.drop(columns=to_drop)
    return df, to_drop


def dataframe_to_json_records(df: pd.DataFrame):
    """Претвара DataFrame у листу речника, спремну за json.dump.
    NaN вредности постају None (тј. null у JSON-у)."""
    return json.loads(df.to_json(orient="records"))


def write_js_variable(records, var_name: str, out_path: Path):
    """Пише податке као обичан .js фајл облика: const VAR_NAME = [ ... ];
    Овакав фајл се учитава преко <script src="..."> у index.html, што ради
    и кад се сајт отвори директно из фајл-система (file://), за разлику од
    fetch()-а који браузери блокирају за локалне фајлове."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"const {var_name} = {payload};\n")


def format_srb_date(yyyymmdd: str):
    """'20250818' -> '18. 8. 2025.'"""
    if not yyyymmdd or len(str(yyyymmdd)) != 8:
        return None
    s = str(yyyymmdd)
    y, mo, d = s[0:4], s[4:6], s[6:8]
    try:
        return f"{int(d)}. {int(mo)}. {y}."
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Обрада једне сезоне
# ---------------------------------------------------------------------------

def process_season(suffix: str, files: dict):
    meta = suffix_to_season_meta(suffix)
    print(f"\n=== Сезона: {meta['label']}  (фолдер: {meta['id']}) ===")

    print(f"  Учитавам термине из: {files['termini'].name} ...")
    termini_df = load_termini(files["termini"])
    print(f"    -> {len(termini_df)} редова (играч по термину)")

    print(f"  Учитавам голове из: {files['golovi'].name} ...")
    golovi_df = load_golovi(files["golovi"])
    print(f"    -> {len(golovi_df)} голова укупно")

    termini_df, dropped_t = drop_allzero_columns(termini_df, TERMINI_PROTECTED_COLS)
    golovi_df, dropped_g = drop_allzero_columns(golovi_df, GOLOVI_PROTECTED_COLS)
    dropped = dropped_t + dropped_g
    if dropped:
        print(f"  Ова сезона нема податке за: {', '.join(dropped)} (колоне уклоњене, сајт их неће приказивати)")

    season_dir = SITE_ROOT / meta["id"]
    data_dir = season_dir / "data"

    write_js_variable(dataframe_to_json_records(termini_df), "TERMINI_DATA", data_dir / "termini.js")
    write_js_variable(dataframe_to_json_records(golovi_df), "GOLOVI_DATA", data_dir / "golovi.js")

    dates = sorted(d for d in termini_df["Date"].dropna().unique().tolist())
    computed_date_from = format_srb_date(dates[0]) if dates else None
    computed_date_to = format_srb_date(dates[-1]) if dates else None
    # користимо фиксне датуме по конвенцији (1.9-31.5 редовна, 1.6-31.8 летња) ако су
    # препознати из имена фајла; иначе (непрепознат формат) паднемо назад на стварне датуме
    date_from = meta.get("dateFrom") or computed_date_from
    date_to = meta.get("dateTo") or computed_date_to

    config = {
        "id": meta["id"],
        "label": meta["label"],
        "icon": meta["icon"],
        "dateFrom": date_from,
        "dateTo": date_to,
    }
    with open(data_dir / "config.js", "w", encoding="utf-8") as f:
        f.write("const SEASON_CONFIG = " + json.dumps(config, ensure_ascii=False) + ";\n")

    if not TEMPLATE_HTML.exists():
        print(f"  ГРЕШКА: не налазим {TEMPLATE_HTML} - прескачем копирање сајта за ову сезону!")
    else:
        shutil.copyfile(TEMPLATE_HTML, season_dir / "index.html")

    print(f"  Сачувано у: {season_dir.relative_to(SITE_ROOT)}/")

    return {
        "id": meta["id"],
        "path": f"{meta['id']}/",
        "label": meta["label"],
        "icon": meta["icon"],
        "dateFrom": date_from,
        "dateTo": date_to,
        "_sort_key": dates[0] if dates else "99999999",
    }


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    print(f"Тражим парове Termini/Golovi Excel фајлова у: {EXCEL_DIR}")
    pairs = find_season_pairs()
    if not pairs:
        print()
        print("Нема пронађених парова Excel фајлова.")
        print("Провери да ли су имена облика 'Termini 202526.xlsx' + 'Golovi 202526.xlsx'.")
        sys.exit(1)

    seasons_meta = []
    for suffix, files in pairs.items():
        seasons_meta.append(process_season(suffix, files))

    seasons_meta.sort(key=lambda s: s["_sort_key"])
    for s in seasons_meta:
        s.pop("_sort_key")

    seasons_js_path = SITE_ROOT / "seasons.js"
    with open(seasons_js_path, "w", encoding="utf-8") as f:
        f.write("const ALL_SEASONS = " + json.dumps(seasons_meta, ensure_ascii=False, indent=2) + ";\n")
    print(f"\nСачувано: {seasons_js_path.relative_to(SITE_ROOT)}")

    if ROOT_INDEX_TEMPLATE.exists():
        shutil.copyfile(ROOT_INDEX_TEMPLATE, SITE_ROOT / "index.html")
        print(f"Сачувано: {(SITE_ROOT / 'index.html').relative_to(SITE_ROOT)}")
    else:
        print(f"УПОЗОРЕЊЕ: не налазим {ROOT_INDEX_TEMPLATE} - корени index.html није освежен.")

    print()
    print(f"Готово! Пронађено сезона: {len(seasons_meta)}")
    for s in seasons_meta:
        print(f"  - {s['label']}  ->  {s['path']}")
    print()
    print("Отвори index.html (у корену) да изабереш сезону.")
    print("(Ради и кад отвориш директно дупли-кликом, без сервера.)")


if __name__ == "__main__":
    main()
