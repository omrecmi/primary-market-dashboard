from __future__ import annotations

import json
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parent.parent
SHORTCUT_PATH = ROOT / "01. Database_all - Shortcut.lnk"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
OUTPUT_HTML = OUTPUT_DIR / "secondary_market_dashboard.html"
OUTPUT_JSON = OUTPUT_DIR / "secondary_market_dashboard_data.json"
PUBLISH_DIR = Path(__file__).resolve().parent / "publish"
PUBLISH_HTML = PUBLISH_DIR / "secondary_market_dashboard.html"
PUBLISH_JSON = PUBLISH_DIR / "secondary_market_dashboard_data.json"

HANOI_FOLDER = "1. Hanoi"
ALL_FILE = "Hanoi Secondary Database all.xlsx"
HIGH_RISE_FILE = "Hanoi High-rise Secondary Database.xlsx"
LOW_RISE_FILE = "Hanoi Low-rise Secondary Database.xlsx"
LANDED_FILE = "Hanoi Landed-house Secondary Database.xlsx"

ASSET_TYPE_LABELS = {
    "cao_tang": "High-rise",
    "thap_tang": "Low-rise",
    "tho_cu": "Landed house",
}
FOCUS_GROUPS = ("VHOP", "VHSC")
LOW_RISE_FOCUS_PROJECTS = (
    "Vinhomes Ocean Park 1",
    "Vinhomes Ocean Park 2",
    "Vinhomes Ocean Park 3",
    "Vinhomes Global Gate",
)
SPECIAL_SCOPE_DISTRICTS = {
    "Gia Lâm",
    "Đông Anh",
    "Đan Phượng",
}
CUTOFF_DAY = 25


def resolve_shortcut(shortcut_path: Path) -> Path:
    escaped = str(shortcut_path).replace("'", "''")
    command = (
        "$w = New-Object -ComObject WScript.Shell; "
        f"$s = $w.CreateShortcut('{escaped}'); "
        "$s.TargetPath"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", command],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return Path(result.stdout.strip())


def month_sort_value(year_month: str) -> int:
    year_text, month_text = year_month.split("-")
    return int(year_text) * 100 + int(month_text)


def build_year_month(dataframe: pd.DataFrame, year_col: str, month_col: str) -> pd.Series:
    year = pd.to_numeric(dataframe[year_col], errors="coerce").astype("Int64")
    month = pd.to_numeric(dataframe[month_col], errors="coerce").astype("Int64")
    valid = year.notna() & month.notna() & month.between(1, 12)
    output = pd.Series(pd.NA, index=dataframe.index, dtype="object")
    output.loc[valid] = (
        year.loc[valid].astype(int).astype(str)
        + "-"
        + month.loc[valid].astype(int).astype(str).str.zfill(2)
    )
    return output


def clean_text(value: Any, fallback: str = "Unknown") -> str:
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return fallback
    return text


def normalize_project_group(value: Any, level: str) -> str:
    text = clean_text(value, "Unknown")
    lowered = text.lower()
    if "vhop" in lowered:
        return "VHOP"
    if "vhsc" in lowered:
        return "VHSC"
    if "vhm" in lowered:
        return "VHM khác"
    if "du an khac" in lowered or "dự án khác" in lowered:
        return "Dự án khác"
    if level == "group1" and text == "Unknown":
        return "Dự án khác"
    return text


def normalize_low_rise_project_name(value: Any) -> str:
    text = clean_text(value, "Unknown")
    lowered = text.lower()
    if lowered == "vinhomes ocean park":
        return "Vinhomes Ocean Park 1"
    return text


def normalize_geo_name(value: Any) -> str:
    text = clean_text(value, "Unknown")
    if text == "Unknown":
        return text
    prefixes = ("Q. ", "H. ", "P. ", "X. ", "TT. ", "TX. ")
    for prefix in prefixes:
        if text.startswith(prefix):
            return text[len(prefix):].strip()
    return text.strip()


def is_scope_district(value: Any) -> bool:
    text = clean_text(value, "")
    if not text:
        return False
    if text.startswith("Q. "):
        return True
    return normalize_geo_name(text) in SPECIAL_SCOPE_DISTRICTS


def previous_year_month(year_month: str) -> str:
    year_text, month_text = year_month.split("-")
    year = int(year_text)
    month = int(month_text)
    if month == 1:
        return f"{year - 1}-12"
    return f"{year}-{month - 1:02d}"


def load_all_secondary(base_dir: Path) -> pd.DataFrame:
    path = base_dir / HANOI_FOLDER / ALL_FILE
    dataframe = pd.read_excel(
        path,
        usecols=[
            "scd2_hash",
            "date_y",
            "date_m",
            "loai_tai_san",
            "district",
            "ward",
        ],
    )
    dataframe["scd2_hash"] = dataframe["scd2_hash"].astype(str).str.strip()
    dataframe["asset_type"] = dataframe["loai_tai_san"].map(
        lambda value: clean_text(value, "Unknown")
    )
    dataframe["year_month"] = build_year_month(dataframe, "date_y", "date_m")
    dataframe["district"] = dataframe["district"].map(lambda value: clean_text(value, "Unknown"))
    dataframe["ward"] = dataframe["ward"].map(lambda value: clean_text(value, "Unknown"))
    dataframe = dataframe[
        dataframe["scd2_hash"].ne("")
        & dataframe["year_month"].notna()
        & dataframe["asset_type"].isin(ASSET_TYPE_LABELS)
        & dataframe["district"].map(is_scope_district)
    ].copy()
    dataframe["asset_type_label"] = dataframe["asset_type"].map(ASSET_TYPE_LABELS)
    return dataframe


def load_high_rise_secondary(base_dir: Path) -> pd.DataFrame:
    path = base_dir / HANOI_FOLDER / HIGH_RISE_FILE
    dataframe = pd.read_excel(
        path,
        usecols=[
            "scd2_hash",
            "date_y",
            "date_m",
            "district",
            "ward",
            "project_name",
            "sector_name",
            "project_name_group1",
            "project_name_group2",
        ],
    )
    dataframe["scd2_hash"] = dataframe["scd2_hash"].astype(str).str.strip()
    dataframe["year_month"] = build_year_month(dataframe, "date_y", "date_m")
    for column in ("district", "ward", "project_name", "sector_name"):
        dataframe[column] = dataframe[column].map(lambda value: clean_text(value, "Unknown"))
    dataframe["project_name_group1"] = dataframe["project_name_group1"].map(
        lambda value: normalize_project_group(value, "group1")
    )
    dataframe["project_name_group2"] = dataframe["project_name_group2"].map(
        lambda value: normalize_project_group(value, "group2")
    )
    dataframe = dataframe[
        dataframe["scd2_hash"].ne("")
        & dataframe["year_month"].notna()
    ].copy()
    return dataframe


def load_low_rise_secondary(base_dir: Path) -> pd.DataFrame:
    path = base_dir / HANOI_FOLDER / LOW_RISE_FILE
    dataframe = pd.read_excel(
        path,
        usecols=[
            "scd2_hash",
            "date_y",
            "date_m",
            "district",
            "project_name",
            "sector_name",
        ],
    )
    dataframe["scd2_hash"] = dataframe["scd2_hash"].astype(str).str.strip()
    dataframe["year_month"] = build_year_month(dataframe, "date_y", "date_m")
    dataframe["district"] = dataframe["district"].map(lambda value: clean_text(value, "Unknown"))
    dataframe["project_name"] = dataframe["project_name"].map(normalize_low_rise_project_name)
    dataframe["sector_name"] = dataframe["sector_name"].map(lambda value: clean_text(value, "Unknown"))
    dataframe = dataframe[
        dataframe["scd2_hash"].ne("")
        & dataframe["year_month"].notna()
    ].copy()
    return dataframe


def load_landed_secondary(base_dir: Path) -> pd.DataFrame:
    path = base_dir / HANOI_FOLDER / LANDED_FILE
    dataframe = pd.read_excel(
        path,
        usecols=[
            "scd2_hash",
            "date_y",
            "date_m",
            "district",
            "ward",
        ],
    )
    dataframe["scd2_hash"] = dataframe["scd2_hash"].astype(str).str.strip()
    dataframe["year_month"] = build_year_month(dataframe, "date_y", "date_m")
    dataframe["district"] = dataframe["district"].map(normalize_geo_name)
    dataframe["ward"] = dataframe["ward"].map(normalize_geo_name)
    dataframe = dataframe[
        dataframe["scd2_hash"].ne("")
        & dataframe["year_month"].notna()
    ].copy()
    return dataframe


def distinct_transactions(
    dataframe: pd.DataFrame,
    dimensions: list[str],
    metric_name: str = "transactions",
) -> pd.DataFrame:
    return (
        dataframe.groupby(dimensions, dropna=False, as_index=False)
        .agg(**{metric_name: ("scd2_hash", "nunique")})
        .sort_values(dimensions)
        .reset_index(drop=True)
    )


def with_mom(rows: pd.DataFrame, value_col: str) -> pd.DataFrame:
    output = rows.sort_values("year_month").copy()
    output["mom_abs"] = output[value_col].diff()
    previous = output[value_col].shift(1)
    output["mom_pct"] = (output["mom_abs"] / previous.replace({0: pd.NA})) * 100
    return output


def dataframe_records(dataframe: pd.DataFrame) -> list[dict[str, Any]]:
    return json.loads(dataframe.to_json(orient="records", force_ascii=False))


def build_dataset() -> dict[str, Any]:
    base_dir = resolve_shortcut(SHORTCUT_PATH)
    all_df = load_all_secondary(base_dir)
    high_rise_df = load_high_rise_secondary(base_dir)
    low_rise_df = load_low_rise_secondary(base_dir)
    landed_df = load_landed_secondary(base_dir)

    monthly_asset = distinct_transactions(
        all_df,
        ["year_month", "asset_type", "asset_type_label"],
    )
    monthly_total = distinct_transactions(all_df, ["year_month"])
    monthly_total = with_mom(monthly_total, "transactions")

    asset_pivot = (
        monthly_asset.pivot_table(
            index="year_month",
            columns="asset_type",
            values="transactions",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
        .rename_axis(None, axis=1)
    )
    for asset_type in ASSET_TYPE_LABELS:
        if asset_type not in asset_pivot.columns:
            asset_pivot[asset_type] = 0

    monthly_overview = monthly_total.merge(asset_pivot, on="year_month", how="left").fillna(0)
    monthly_overview = monthly_overview.rename(
        columns={
            "cao_tang": "high_rise",
            "thap_tang": "low_rise",
            "tho_cu": "landed_house",
        }
    )
    monthly_overview["sort_value"] = monthly_overview["year_month"].map(month_sort_value)
    monthly_overview = monthly_overview.sort_values("sort_value").drop(columns=["sort_value"])

    latest_available_row = monthly_overview.iloc[-1]
    latest_available_month = str(latest_available_row["year_month"])
    now = datetime.now()
    current_month = now.strftime("%Y-%m")
    if now.day < CUTOFF_DAY:
        reporting_month_target = previous_year_month(current_month)
    else:
        reporting_month_target = current_month

    reporting_candidates = monthly_overview[
        monthly_overview["year_month"].map(month_sort_value) <= month_sort_value(reporting_month_target)
    ].copy()
    if reporting_candidates.empty:
        latest_row = latest_available_row
    else:
        latest_row = reporting_candidates.iloc[-1]
    latest_month_is_partial = str(latest_row["year_month"]) != latest_available_month
    previous_candidates = monthly_overview[
        monthly_overview["year_month"].map(month_sort_value) < month_sort_value(str(latest_row["year_month"]))
    ].copy()
    previous_row = previous_candidates.iloc[-1] if not previous_candidates.empty else None

    latest_asset_mix = {
        "high_rise": int(latest_row["high_rise"]),
        "low_rise": int(latest_row["low_rise"]),
        "landed_house": int(latest_row["landed_house"]),
    }

    high_rise_group1 = distinct_transactions(
        high_rise_df,
        ["year_month", "project_name_group1"],
    ).rename(columns={"project_name_group1": "group_name"})
    high_rise_group2 = distinct_transactions(
        high_rise_df,
        ["year_month", "project_name_group2"],
    ).rename(columns={"project_name_group2": "group_name"})

    focus_rows = high_rise_df[high_rise_df["project_name_group2"].isin(FOCUS_GROUPS)].copy()
    sector_monthly = distinct_transactions(
        focus_rows,
        ["project_name_group2", "year_month", "sector_name"],
    ).rename(columns={"project_name_group2": "focus_group"})
    sector_totals = (
        sector_monthly.groupby(["focus_group", "sector_name"], as_index=False)
        .agg(total_transactions=("transactions", "sum"))
    )
    sector_monthly = sector_monthly.merge(
        sector_totals,
        on=["focus_group", "sector_name"],
        how="left",
    )

    latest_month = str(latest_row["year_month"])
    previous_month = str(previous_row["year_month"]) if previous_row is not None else None

    sector_latest_rows: dict[str, list[dict[str, Any]]] = {}
    for focus_group in FOCUS_GROUPS:
        latest_sector = sector_monthly[
            (sector_monthly["focus_group"] == focus_group)
            & (sector_monthly["year_month"] == latest_month)
        ][["sector_name", "transactions", "total_transactions"]].copy()
        latest_sector = latest_sector.rename(columns={"transactions": "latest_count"})
        if previous_month:
            previous_sector = sector_monthly[
                (sector_monthly["focus_group"] == focus_group)
                & (sector_monthly["year_month"] == previous_month)
            ][["sector_name", "transactions"]].copy()
            previous_sector = previous_sector.rename(columns={"transactions": "prev_count"})
            latest_sector = latest_sector.merge(previous_sector, on="sector_name", how="left")
        else:
            latest_sector["prev_count"] = pd.NA
        latest_sector["prev_count"] = latest_sector["prev_count"].fillna(0)
        latest_sector["mom_abs"] = latest_sector["latest_count"] - latest_sector["prev_count"]
        latest_sector["mom_pct"] = (
            latest_sector["mom_abs"]
            / latest_sector["prev_count"].replace({0: pd.NA})
            * 100
        )
        latest_sector = latest_sector.sort_values(
            ["latest_count", "mom_abs", "total_transactions", "sector_name"],
            ascending=[False, False, False, True],
        ).reset_index(drop=True)
        sector_latest_rows[focus_group] = dataframe_records(latest_sector.head(20))

    low_rise_focus = low_rise_df[low_rise_df["project_name"].isin(LOW_RISE_FOCUS_PROJECTS)].copy()
    low_rise_project_monthly = distinct_transactions(
        low_rise_focus,
        ["project_name", "year_month"],
    )
    low_rise_sector_monthly = distinct_transactions(
        low_rise_focus,
        ["project_name", "year_month", "sector_name"],
    )
    low_rise_sector_totals = (
        low_rise_sector_monthly.groupby(["project_name", "sector_name"], as_index=False)
        .agg(total_transactions=("transactions", "sum"))
    )
    low_rise_sector_monthly = low_rise_sector_monthly.merge(
        low_rise_sector_totals,
        on=["project_name", "sector_name"],
        how="left",
    )
    low_rise_sector_latest: dict[str, list[dict[str, Any]]] = {}
    for project_name in LOW_RISE_FOCUS_PROJECTS:
        latest_sector = low_rise_sector_monthly[
            (low_rise_sector_monthly["project_name"] == project_name)
            & (low_rise_sector_monthly["year_month"] == latest_month)
        ][["sector_name", "transactions", "total_transactions"]].copy()
        latest_sector = latest_sector.rename(columns={"transactions": "latest_count"})
        if previous_month:
            previous_sector = low_rise_sector_monthly[
                (low_rise_sector_monthly["project_name"] == project_name)
                & (low_rise_sector_monthly["year_month"] == previous_month)
            ][["sector_name", "transactions"]].copy()
            previous_sector = previous_sector.rename(columns={"transactions": "prev_count"})
            latest_sector = latest_sector.merge(previous_sector, on="sector_name", how="left")
        else:
            latest_sector["prev_count"] = pd.NA
        latest_sector["prev_count"] = latest_sector["prev_count"].fillna(0)
        latest_sector["mom_abs"] = latest_sector["latest_count"] - latest_sector["prev_count"]
        latest_sector["mom_pct"] = (
            latest_sector["mom_abs"]
            / latest_sector["prev_count"].replace({0: pd.NA})
            * 100
        )
        latest_sector = latest_sector.sort_values(
            ["latest_count", "mom_abs", "total_transactions", "sector_name"],
            ascending=[False, False, False, True],
        ).reset_index(drop=True)
        low_rise_sector_latest[project_name] = dataframe_records(latest_sector.head(20))

    landed_district_monthly = distinct_transactions(
        landed_df,
        ["district", "year_month"],
    )
    landed_district_totals = (
        landed_district_monthly.groupby(["district"], as_index=False)
        .agg(total_transactions=("transactions", "sum"))
    )
    landed_district_monthly = landed_district_monthly.merge(
        landed_district_totals,
        on="district",
        how="left",
    )
    landed_ward_monthly = distinct_transactions(
        landed_df,
        ["district", "ward", "year_month"],
    )
    landed_ward_totals = (
        landed_ward_monthly.groupby(["district", "ward"], as_index=False)
        .agg(total_transactions=("transactions", "sum"))
    )
    landed_ward_monthly = landed_ward_monthly.merge(
        landed_ward_totals,
        on=["district", "ward"],
        how="left",
    )

    landed_district_latest = landed_district_monthly[
        landed_district_monthly["year_month"] == latest_month
    ][["district", "transactions", "total_transactions"]].copy()
    landed_district_latest = landed_district_latest.rename(columns={"transactions": "latest_count"})
    if previous_month:
        previous_district = landed_district_monthly[
            landed_district_monthly["year_month"] == previous_month
        ][["district", "transactions"]].copy()
        previous_district = previous_district.rename(columns={"transactions": "prev_count"})
        landed_district_latest = landed_district_latest.merge(previous_district, on="district", how="left")
    else:
        landed_district_latest["prev_count"] = pd.NA
    landed_district_latest["prev_count"] = landed_district_latest["prev_count"].fillna(0)
    landed_district_latest["mom_abs"] = landed_district_latest["latest_count"] - landed_district_latest["prev_count"]
    landed_district_latest["mom_pct"] = (
        landed_district_latest["mom_abs"]
        / landed_district_latest["prev_count"].replace({0: pd.NA})
        * 100
    )
    landed_district_latest = landed_district_latest.sort_values(
        ["latest_count", "mom_abs", "total_transactions", "district"],
        ascending=[False, False, False, True],
    ).reset_index(drop=True)

    source_stats = []
    for label, file_name in (
        ("All secondary", ALL_FILE),
        ("High-rise secondary", HIGH_RISE_FILE),
        ("Low-rise secondary", LOW_RISE_FILE),
        ("Landed-house secondary", LANDED_FILE),
    ):
        source_stats.append(
            {
                "label": label,
                "file_name": file_name,
                "exists": (base_dir / HANOI_FOLDER / file_name).exists(),
            }
        )

    return {
        "meta": {
            "title": "Hanoi Secondary Market Dashboard",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "base_dir": str(base_dir / HANOI_FOLDER),
            "latest_available_month": latest_available_month,
            "latest_month": latest_month,
            "previous_month": previous_month,
            "latest_month_is_partial": latest_month_is_partial,
            "reporting_month_target": reporting_month_target,
            "cutoff_day": CUTOFF_DAY,
            "months": [str(value) for value in monthly_overview["year_month"].tolist()],
            "focus_groups": list(FOCUS_GROUPS),
            "source_stats": source_stats,
            "low_rise_focus_projects": list(LOW_RISE_FOCUS_PROJECTS),
            "landed_districts": sorted(landed_df["district"].dropna().astype(str).unique().tolist()),
        },
        "summary": {
            "latest_month": latest_month,
            "latest_available_month": latest_available_month,
            "latest_month_is_partial": latest_month_is_partial,
            "reporting_month_target": reporting_month_target,
            "cutoff_day": CUTOFF_DAY,
            "latest_transactions": int(latest_row["transactions"]),
            "previous_month": previous_month,
            "previous_transactions": int(previous_row["transactions"]) if previous_row is not None else None,
            "mom_abs": int(latest_row["mom_abs"]) if pd.notna(latest_row["mom_abs"]) else None,
            "mom_pct": float(latest_row["mom_pct"]) if pd.notna(latest_row["mom_pct"]) else None,
            "latest_asset_mix": latest_asset_mix,
            "monthly_points": int(len(monthly_overview)),
            "distinct_transactions_all_period": int(all_df["scd2_hash"].nunique()),
            "distinct_high_rise_all_period": int(high_rise_df["scd2_hash"].nunique()),
        },
        "monthly_overview": dataframe_records(monthly_overview),
        "monthly_asset": dataframe_records(monthly_asset),
        "high_rise_group1": dataframe_records(high_rise_group1),
        "high_rise_group2": dataframe_records(high_rise_group2),
        "high_rise_sector_monthly": dataframe_records(sector_monthly),
        "high_rise_sector_latest": sector_latest_rows,
        "low_rise_project_monthly": dataframe_records(low_rise_project_monthly),
        "low_rise_sector_monthly": dataframe_records(low_rise_sector_monthly),
        "low_rise_sector_latest": low_rise_sector_latest,
        "landed_district_monthly": dataframe_records(landed_district_monthly),
        "landed_district_latest": dataframe_records(landed_district_latest.head(25)),
        "landed_ward_monthly": dataframe_records(landed_ward_monthly),
    }


def build_html(dataset: dict[str, Any]) -> str:
    data_json = json.dumps(dataset, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Hanoi Secondary Market Dashboard</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
  <style>
    :root {{
      --bg: #f4efe7;
      --panel: rgba(255, 251, 244, 0.94);
      --panel-strong: #fffdf8;
      --ink: #253238;
      --muted: #6d7a81;
      --line: rgba(90, 116, 125, 0.18);
      --navy: #173f4f;
      --teal: #2f7d79;
      --sand: #e8d9c5;
      --gold: #b88b32;
      --green: #2d8f6f;
      --red: #d95a5a;
      --blue: #4c86c6;
      --high-rise: #173f4f;
      --low-rise: #2f7d79;
      --landed: #b88b32;
      --shadow: 0 22px 48px rgba(23, 63, 79, 0.08);
      --radius: 24px;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Plus Jakarta Sans", "Segoe UI", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(184, 139, 50, 0.18), transparent 25%),
        radial-gradient(circle at top right, rgba(47, 125, 121, 0.18), transparent 28%),
        linear-gradient(180deg, #fffdf8 0%, var(--bg) 100%);
    }}
    .shell {{
      max-width: 1520px;
      margin: 0 auto;
      padding: 28px 24px 44px;
    }}
    .hero {{
      display: grid;
      grid-template-columns: minmax(0, 1.2fr) minmax(300px, 0.8fr);
      gap: 18px;
      margin-bottom: 22px;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid rgba(215, 203, 185, 0.9);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      backdrop-filter: blur(10px);
    }}
    .hero-card {{
      padding: 28px;
      position: relative;
      overflow: hidden;
    }}
    .hero-card::after {{
      content: "";
      position: absolute;
      right: -30px;
      bottom: -48px;
      width: 230px;
      height: 230px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(23, 63, 79, 0.12), transparent 72%);
    }}
    h1 {{
      margin: 0 0 12px;
      font-size: clamp(34px, 4.6vw, 56px);
      line-height: 0.94;
      letter-spacing: -0.05em;
    }}
    .subhead {{
      margin: 0;
      max-width: 840px;
      color: var(--muted);
      line-height: 1.7;
      font-size: 15px;
    }}
    .chip-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 18px;
    }}
    .chip {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 9px 14px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.78);
      border: 1px solid rgba(90, 116, 125, 0.12);
      color: var(--navy);
      font-size: 13px;
      font-weight: 700;
    }}
    .control-card {{
      padding: 22px;
      display: grid;
      gap: 16px;
      align-content: start;
    }}
    .control-block {{
      display: grid;
      gap: 10px;
    }}
    .control-label {{
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      color: var(--muted);
    }}
    .button-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }}
    .select-row {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }}
    button.toggle, select.control-select {{
      border: 1px solid var(--line);
      border-radius: 14px;
      background: #fff;
      color: var(--ink);
      font: inherit;
      font-size: 14px;
      font-weight: 700;
      padding: 10px 14px;
    }}
    button.refresh-button {{
      border: 1px solid var(--blue);
      background: var(--blue);
      color: #fff;
      border-radius: 14px;
      padding: 10px 14px;
      font: inherit;
      font-size: 14px;
      font-weight: 800;
      cursor: pointer;
      transition: transform 0.18s ease, opacity 0.18s ease;
    }}
    button.refresh-button:hover {{
      transform: translateY(-1px);
      opacity: 0.92;
    }}
    button.toggle {{
      cursor: pointer;
      transition: transform 0.18s ease, border-color 0.18s ease, background 0.18s ease;
    }}
    button.toggle:hover {{
      transform: translateY(-1px);
      border-color: var(--teal);
    }}
    button.toggle.active {{
      background: var(--navy);
      border-color: var(--navy);
      color: #fff;
    }}
    .summary-grid {{
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }}
    .summary-card {{
      padding: 18px;
      min-height: 150px;
    }}
    .summary-card h3 {{
      margin: 0 0 10px;
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      color: var(--muted);
    }}
    .summary-card strong {{
      display: block;
      margin-bottom: 8px;
      font-size: 30px;
      letter-spacing: -0.05em;
    }}
    .summary-card p {{
      margin: 0;
      font-size: 13px;
      line-height: 1.6;
      color: var(--muted);
    }}
    .summary-card .delta.up {{ color: var(--green); }}
    .summary-card .delta.down {{ color: var(--red); }}
    .summary-card .delta.flat {{ color: var(--muted); }}
    .section {{
      margin-bottom: 24px;
    }}
    .section-head {{
      display: flex;
      justify-content: space-between;
      gap: 14px;
      align-items: flex-end;
      margin-bottom: 12px;
    }}
    .section-head h2 {{
      margin: 0;
      font-size: 23px;
      letter-spacing: -0.03em;
    }}
    .section-note {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.55;
      max-width: 840px;
    }}
    .grid-2, .grid-3 {{
      display: grid;
      gap: 18px;
    }}
    .grid-2 {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    .grid-3 {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    .chart-card, .table-card {{
      padding: 20px;
    }}
    .chart-head {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: flex-start;
      margin-bottom: 12px;
    }}
    .chart-head h3 {{
      margin: 0 0 4px;
      font-size: 18px;
      letter-spacing: -0.02em;
    }}
    .chart-head p {{
      margin: 0;
      font-size: 13px;
      color: var(--muted);
    }}
    .canvas-wrap {{
      height: 320px;
    }}
    .canvas-wrap.tall {{
      height: 360px;
    }}
    .table-wrap {{
      max-height: 420px;
      overflow: auto;
      border-radius: 16px;
      border: 1px solid rgba(215, 203, 185, 0.74);
      background: var(--panel-strong);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      min-width: 680px;
    }}
    thead th {{
      position: sticky;
      top: 0;
      z-index: 1;
      padding: 12px 10px;
      text-align: left;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--muted);
      background: #f9f5ee;
      border-bottom: 1px solid var(--line);
    }}
    thead th.sortable {{
      cursor: pointer;
      user-select: none;
      transition: color 0.18s ease, background 0.18s ease;
    }}
    thead th.sortable:hover {{
      color: var(--navy);
      background: #f3ece0;
    }}
    .sort-indicator {{
      display: inline-block;
      min-width: 14px;
      margin-left: 6px;
      color: var(--blue);
      font-size: 10px;
    }}
    tbody td {{
      padding: 11px 10px;
      border-bottom: 1px solid rgba(215, 203, 185, 0.58);
      font-size: 13px;
      vertical-align: top;
    }}
    tbody td.up {{
      color: var(--green);
      background: rgba(45, 143, 111, 0.08);
      font-weight: 700;
    }}
    tbody td.down {{
      color: var(--red);
      background: rgba(217, 90, 90, 0.08);
      font-weight: 700;
    }}
    tbody td.flat {{
      color: var(--muted);
      background: rgba(109, 122, 129, 0.08);
      font-weight: 700;
    }}
    tbody tr:hover {{
      background: rgba(76, 134, 198, 0.08);
    }}
    .footnote {{
      margin-top: 16px;
      font-size: 13px;
      color: var(--muted);
      line-height: 1.65;
    }}
    @media (max-width: 1180px) {{
      .hero, .summary-grid, .grid-2, .grid-3 {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <article class="panel hero-card">
        <h1>Hanoi Secondary<br/>Market Dashboard</h1>
        <p class="subhead">
          Monthly transaction monitor for Hanoi secondary market. The core metric is
          <strong>count distinct <code>scd2_hash</code></strong> by <code>date_y</code> and <code>date_m</code>.
          High-rise drill-down focuses on Vinhomes Ocean Park and Vinhomes Smart City so sector-level month-on-month change is visible immediately.
        </p>
        <div class="chip-row">
          <span class="chip">Latest reporting month: <span id="latestMonthChip"></span></span>
          <span class="chip">Latest available month: <span id="latestAvailableMonthChip"></span></span>
          <span class="chip">Total history: <span id="historyChip"></span></span>
        </div>
      </article>
      <aside class="panel control-card">
        <div class="control-block">
          <div class="control-label">Window</div>
          <div class="select-row">
            <select id="startMonthSelect" class="control-select"></select>
            <select id="endMonthSelect" class="control-select"></select>
          </div>
          <div class="button-row">
            <button id="applyWindowButton" class="refresh-button">Refresh</button>
          </div>
        </div>
        <div class="control-block">
          <div class="control-label">Source</div>
          <div id="sourceBlock" class="section-note"></div>
        </div>
      </aside>
    </section>

    <section class="summary-grid">
      <article class="panel summary-card">
        <h3>Total Transactions</h3>
        <strong id="summaryTotal"></strong>
        <p id="summaryDelta"></p>
      </article>
      <article class="panel summary-card">
        <h3>High-rise Share</h3>
        <strong id="summaryHighRise"></strong>
        <p id="summaryHighRiseNote"></p>
      </article>
      <article class="panel summary-card">
        <h3>Low-rise Share</h3>
        <strong id="summaryLowRise"></strong>
        <p id="summaryLowRiseNote"></p>
      </article>
      <article class="panel summary-card">
        <h3>Landed-house Share</h3>
        <strong id="summaryLanded"></strong>
        <p id="summaryLandedNote"></p>
      </article>
    </section>

    <section class="section">
      <div class="section-head">
        <div>
          <h2>Market Overview</h2>
          <div class="section-note">
            Fast read of how total secondary transactions move by month, where the asset-class mix changes, and which months show the sharpest month-on-month swing.
          </div>
        </div>
      </div>
      <div class="grid-2">
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>Total Transactions by Month</h3>
              <p>Distinct <code>scd2_hash</code> across all secondary transactions.</p>
            </div>
          </div>
          <div class="canvas-wrap"><canvas id="chartTotalTrend"></canvas></div>
        </article>
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>Asset Mix by Month</h3>
              <p>High-rise, low-rise, and landed-house contribution by month.</p>
            </div>
          </div>
          <div class="canvas-wrap"><canvas id="chartAssetMix"></canvas></div>
        </article>
      </div>
      <div class="grid-2" style="margin-top: 18px;">
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>Month-on-Month Change</h3>
              <p>Absolute transaction delta versus previous month.</p>
            </div>
          </div>
          <div class="canvas-wrap"><canvas id="chartMoM"></canvas></div>
        </article>
        <article class="panel table-card">
          <div class="chart-head">
            <div>
              <h3>Monthly Snapshot Table</h3>
              <p>Latest months with total and asset-type mix.</p>
            </div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th class="sortable" data-table="monthlyOverviewTable" data-key="year_month" data-type="text">Month<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="monthlyOverviewTable" data-key="transactions" data-type="number">Total<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="monthlyOverviewTable" data-key="mom_abs" data-type="number">MoM<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="monthlyOverviewTable" data-key="high_rise" data-type="number">High-rise<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="monthlyOverviewTable" data-key="low_rise" data-type="number">Low-rise<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="monthlyOverviewTable" data-key="landed_house" data-type="number">Landed<span class="sort-indicator"></span></th>
                </tr>
              </thead>
              <tbody id="monthlyOverviewTable"></tbody>
            </table>
          </div>
        </article>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <div>
          <h2>High-rise Deep Dive</h2>
          <div class="control-block" style="margin-top: 10px; margin-bottom: 10px; max-width: 360px;">
            <div class="control-label">High-rise focus group</div>
            <select id="focusGroupSelect" class="control-select"></select>
          </div>
          <div class="section-note">
            This section uses <code>Hanoi High-rise Secondary Database.xlsx</code>. It keeps project grouping and sector detail so you can track how Vinhomes Ocean Park and Vinhomes Smart City behave by month.
          </div>
        </div>
      </div>
      <div class="grid-2">
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>Transactions by Project Group 2</h3>
              <p>Monthly distinct transactions for VHOP, VHSC, other Vinhomes, and other projects.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartGroup2"></canvas></div>
        </article>
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>Transactions by Project Group 1</h3>
              <p>Monthly view of broader high-rise grouping.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartGroup1"></canvas></div>
        </article>
      </div>
      <div class="grid-2" style="margin-top: 18px;">
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3 id="sectorTrendTitle">Sector Trend</h3>
              <p id="sectorTrendNote">Top sectors by historical transactions inside the selected focus group.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartSectorTrend"></canvas></div>
        </article>
        <article class="panel table-card">
          <div class="chart-head">
            <div>
              <h3 id="sectorTableTitle">Sector Delta Table</h3>
              <p id="sectorTableNote">Latest month versus previous month.</p>
            </div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th class="sortable" data-table="sectorDeltaTable" data-key="sector_name" data-type="text">Sector<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="sectorDeltaTable" data-key="latest_count" data-type="number">Latest<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="sectorDeltaTable" data-key="prev_count" data-type="number">Prev<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="sectorDeltaTable" data-key="mom_abs" data-type="number">MoM<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="sectorDeltaTable" data-key="mom_pct" data-type="number">MoM %<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="sectorDeltaTable" data-key="total_transactions" data-type="number">Total history<span class="sort-indicator"></span></th>
                </tr>
              </thead>
              <tbody id="sectorDeltaTable"></tbody>
            </table>
          </div>
        </article>
      </div>
      <div class="footnote">
        Assumption in this version: top-line market counts come from <code>Hanoi Secondary Database all.xlsx</code>,
        while high-rise drill-down uses the dedicated high-rise database for richer fields like <code>project_name_group1</code>,
        <code>project_name_group2</code>, and <code>sector_name</code>. If the newest calendar month is still in progress,
        headline cards and sector delta default to the latest completed month.
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <div>
          <h2>Low-rise Focus</h2>
          <div class="control-block" style="margin-top: 10px; margin-bottom: 10px; max-width: 360px;">
            <div class="control-label">Low-rise focus project</div>
            <select id="lowRiseProjectSelect" class="control-select"></select>
          </div>
          <div class="section-note">
            This section tracks low-rise secondary transactions for Vinhomes Ocean Park 1, 2, 3, and Vinhomes Global Gate. Sector-level detail uses <code>sector_name</code> where it is available in the source file.
          </div>
        </div>
      </div>
      <div class="grid-2">
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>Transactions by Focus Project</h3>
              <p>Monthly distinct transactions for the four low-rise focus projects.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartLowRiseProjects"></canvas></div>
        </article>
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3 id="lowRiseSectorTrendTitle">Low-rise Sector Trend</h3>
              <p id="lowRiseSectorTrendNote">Top sectors inside the selected low-rise project.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartLowRiseSectorTrend"></canvas></div>
        </article>
      </div>
      <div class="grid-2" style="margin-top: 18px;">
        <article class="panel table-card">
          <div class="chart-head">
            <div>
              <h3 id="lowRiseSectorTableTitle">Low-rise Sector Delta</h3>
              <p id="lowRiseSectorTableNote">Latest month versus previous month.</p>
            </div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th class="sortable" data-table="lowRiseSectorDeltaTable" data-key="sector_name" data-type="text">Sector<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="lowRiseSectorDeltaTable" data-key="latest_count" data-type="number">Latest<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="lowRiseSectorDeltaTable" data-key="prev_count" data-type="number">Prev<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="lowRiseSectorDeltaTable" data-key="mom_abs" data-type="number">MoM<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="lowRiseSectorDeltaTable" data-key="mom_pct" data-type="number">MoM %<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="lowRiseSectorDeltaTable" data-key="total_transactions" data-type="number">Total history<span class="sort-indicator"></span></th>
                </tr>
              </thead>
              <tbody id="lowRiseSectorDeltaTable"></tbody>
            </table>
          </div>
        </article>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <div>
          <h2>Landed-house Focus</h2>
          <div class="section-note">
            This section tracks landed-house secondary transactions by district, then drills into ward-level change for the selected district.
          </div>
        </div>
      </div>
      <div class="control-block" style="margin-bottom: 18px;">
        <div class="control-label">Landed district focus</div>
        <div class="select-row">
          <select id="landedDistrictSelect" class="control-select"></select>
        </div>
      </div>
      <div class="grid-2">
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3>District Trend</h3>
              <p>Top districts by historical landed-house transaction count.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartLandedDistrictTrend"></canvas></div>
        </article>
        <article class="panel chart-card">
          <div class="chart-head">
            <div>
              <h3 id="landedWardTrendTitle">Ward Trend</h3>
              <p id="landedWardTrendNote">Top wards in the selected district by historical landed-house transaction count.</p>
            </div>
          </div>
          <div class="canvas-wrap tall"><canvas id="chartLongBienWardTrend"></canvas></div>
        </article>
      </div>
      <div class="grid-2" style="margin-top: 18px;">
        <article class="panel table-card">
          <div class="chart-head">
            <div>
              <h3>District Delta Table</h3>
              <p>Latest reporting month by district.</p>
            </div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th class="sortable" data-table="landedDistrictDeltaTable" data-key="district" data-type="text">District<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="landedDistrictDeltaTable" data-key="latest_count" data-type="number">Latest<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="landedDistrictDeltaTable" data-key="prev_count" data-type="number">Prev<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="landedDistrictDeltaTable" data-key="mom_abs" data-type="number">MoM<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="landedDistrictDeltaTable" data-key="mom_pct" data-type="number">MoM %<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="landedDistrictDeltaTable" data-key="total_transactions" data-type="number">Total history<span class="sort-indicator"></span></th>
                </tr>
              </thead>
              <tbody id="landedDistrictDeltaTable"></tbody>
            </table>
          </div>
        </article>
        <article class="panel table-card">
          <div class="chart-head">
            <div>
              <h3 id="landedWardTableTitle">Ward Delta Table</h3>
              <p id="landedWardTableNote">Latest reporting month by ward.</p>
            </div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th class="sortable" data-table="longBienWardDeltaTable" data-key="ward" data-type="text">Ward<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="longBienWardDeltaTable" data-key="latest_count" data-type="number">Latest<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="longBienWardDeltaTable" data-key="prev_count" data-type="number">Prev<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="longBienWardDeltaTable" data-key="mom_abs" data-type="number">MoM<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="longBienWardDeltaTable" data-key="mom_pct" data-type="number">MoM %<span class="sort-indicator"></span></th>
                  <th class="sortable" data-table="longBienWardDeltaTable" data-key="total_transactions" data-type="number">Total history<span class="sort-indicator"></span></th>
                </tr>
              </thead>
              <tbody id="longBienWardDeltaTable"></tbody>
            </table>
          </div>
        </article>
      </div>
    </section>
  </div>

  <script>
    const dashboardData = {data_json};
    const charts = {{}};
    const allMonths = dashboardData.meta.months || [];
    const defaultEndMonth = dashboardData.summary.latest_month || allMonths[allMonths.length - 1] || "";
    const state = {{
      startMonth: allMonths[0] || "",
      endMonth: defaultEndMonth,
      pendingStartMonth: allMonths[0] || "",
      pendingEndMonth: defaultEndMonth,
      focusGroup: dashboardData.meta.focus_groups[0] || "VHOP",
      lowRiseProject: dashboardData.meta.low_rise_focus_projects[0] || "Vinhomes Ocean Park 1",
      landedDistrict: dashboardData.meta.landed_districts?.includes("Long Biên")
        ? "Long Biên"
        : (dashboardData.meta.landed_districts?.[0] || ""),
      tableSorts: {{}},
    }};

    const palette = {{
      highRise: getComputedStyle(document.documentElement).getPropertyValue("--high-rise").trim(),
      lowRise: getComputedStyle(document.documentElement).getPropertyValue("--low-rise").trim(),
      landed: getComputedStyle(document.documentElement).getPropertyValue("--landed").trim(),
      navy: getComputedStyle(document.documentElement).getPropertyValue("--navy").trim(),
      blue: getComputedStyle(document.documentElement).getPropertyValue("--blue").trim(),
      teal: getComputedStyle(document.documentElement).getPropertyValue("--teal").trim(),
      gold: getComputedStyle(document.documentElement).getPropertyValue("--gold").trim(),
      green: getComputedStyle(document.documentElement).getPropertyValue("--green").trim(),
      red: getComputedStyle(document.documentElement).getPropertyValue("--red").trim(),
      muted: getComputedStyle(document.documentElement).getPropertyValue("--muted").trim(),
    }};

    const groupColors = [
      "#173f4f",
      "#2f7d79",
      "#b88b32",
      "#4c86c6",
      "#935f9f",
      "#c06a53",
      "#5f6b6d",
      "#93a7aa",
    ];

    function formatInt(value) {{
      if (value === null || value === undefined || Number.isNaN(Number(value))) {{
        return "NA";
      }}
      return Number(value).toLocaleString("en-US", {{ maximumFractionDigits: 0 }});
    }}

    function formatPct(value) {{
      if (value === null || value === undefined || Number.isNaN(Number(value))) {{
        return "NA";
      }}
      return `${{Number(value).toLocaleString("en-US", {{ maximumFractionDigits: 1 }})}}%`;
    }}

    function deltaClass(value) {{
      if (value > 0) return "up";
      if (value < 0) return "down";
      return "flat";
    }}

    function currentWindowRows(rows) {{
      const start = state.startMonth || allMonths[0] || "";
      const end = state.endMonth || allMonths[allMonths.length - 1] || "";
      return rows.filter((row) => row.year_month >= start && row.year_month <= end);
    }}

    function destroyChart(chartId) {{
      if (charts[chartId]) {{
        charts[chartId].destroy();
      }}
    }}

    function buildChart(chartId, config) {{
      destroyChart(chartId);
      charts[chartId] = new Chart(document.getElementById(chartId), config);
    }}

    function compareSortValues(left, right, type) {{
      if (type === "number") {{
        const a = left === null || left === undefined || left === "" ? Number.NEGATIVE_INFINITY : Number(left);
        const b = right === null || right === undefined || right === "" ? Number.NEGATIVE_INFINITY : Number(right);
        return a - b;
      }}
      const a = (left ?? "").toString();
      const b = (right ?? "").toString();
      return a.localeCompare(b);
    }}

    function sortRows(tableId, rows) {{
      const sort = state.tableSorts[tableId];
      if (!sort) {{
        return rows.slice();
      }}
      return rows.slice().sort((left, right) => {{
        const base = compareSortValues(left[sort.key], right[sort.key], sort.type);
        if (base !== 0) {{
          return sort.dir === "asc" ? base : -base;
        }}
        return 0;
      }});
    }}

    function refreshSortIndicators() {{
      document.querySelectorAll("th.sortable").forEach((header) => {{
        const indicator = header.querySelector(".sort-indicator");
        if (!indicator) {{
          return;
        }}
        const sort = state.tableSorts[header.dataset.table];
        if (sort && sort.key === header.dataset.key) {{
          indicator.textContent = sort.dir === "asc" ? "▲" : "▼";
        }} else {{
          indicator.textContent = "";
        }}
      }});
    }}

    function groupSeries(rows, labelField) {{
      const months = currentWindowRows(dashboardData.monthly_overview).map((row) => row.year_month);
      const groups = [...new Set(rows.map((row) => row[labelField]))];
      return {{
        months,
        datasets: groups.map((group, index) => {{
          const monthMap = new Map(
            rows.filter((row) => row[labelField] === group).map((row) => [row.year_month, row.transactions])
          );
          return {{
            label: group,
            data: months.map((month) => monthMap.get(month) || 0),
            borderColor: groupColors[index % groupColors.length],
            backgroundColor: groupColors[index % groupColors.length],
            borderWidth: 2,
            tension: 0.28,
            fill: false,
          }};
        }}),
      }};
    }}

    function topSectorsForFocus() {{
      const focusRows = dashboardData.high_rise_sector_monthly.filter(
        (row) => row.focus_group === state.focusGroup
      );
      const totals = new Map();
      focusRows.forEach((row) => {{
        totals.set(row.sector_name, row.total_transactions || 0);
      }});
      return [...totals.entries()]
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .slice(0, 8)
        .map(([name]) => name);
    }}

    function renderSummary() {{
      const summary = dashboardData.summary;
      const latestMix = summary.latest_asset_mix;
      const total = summary.latest_transactions || 0;
      const highRiseShare = total ? (latestMix.high_rise / total) * 100 : null;
      const lowRiseShare = total ? (latestMix.low_rise / total) * 100 : null;
      const landedShare = total ? (latestMix.landed_house / total) * 100 : null;
      document.getElementById("latestMonthChip").textContent = summary.latest_month;
      document.getElementById("latestAvailableMonthChip").textContent = summary.latest_available_month;
      document.getElementById("historyChip").textContent = `${{summary.monthly_points}} months`;
      document.getElementById("summaryTotal").textContent = formatInt(summary.latest_transactions);
      const deltaText = summary.mom_abs === null
        ? "First month in series."
        : `${{summary.latest_month}} vs ${{summary.previous_month}}: ${{summary.mom_abs > 0 ? "+" : ""}}${{formatInt(summary.mom_abs)}} (${{formatPct(summary.mom_pct)}})`;
      const partialNote = summary.latest_month_is_partial
        ? ` Latest available month ${{summary.latest_available_month}} is partial, so the headline uses ${{summary.latest_month}}.`
        : "";
      document.getElementById("summaryDelta").innerHTML = `<span class="delta ${{deltaClass(summary.mom_abs || 0)}}">${{deltaText}}</span>${{partialNote}}`;
      document.getElementById("summaryHighRise").textContent = formatPct(highRiseShare);
      document.getElementById("summaryHighRiseNote").textContent = `${{formatInt(latestMix.high_rise)}} transactions in ${{summary.latest_month}}`;
      document.getElementById("summaryLowRise").textContent = formatPct(lowRiseShare);
      document.getElementById("summaryLowRiseNote").textContent = `${{formatInt(latestMix.low_rise)}} transactions in ${{summary.latest_month}}`;
      document.getElementById("summaryLanded").textContent = formatPct(landedShare);
      document.getElementById("summaryLandedNote").textContent = `${{formatInt(latestMix.landed_house)}} transactions in ${{summary.latest_month}}`;
      document.getElementById("sourceBlock").innerHTML = [
        ...dashboardData.meta.source_stats.map((row) => `${{row.label}}: <code>${{row.file_name}}</code>`),
        `Scope: all quận + H. Gia Lâm + H. Đông Anh + H. Đan Phượng`,
        `Cutoff day: ${{summary.cutoff_day}}. Reporting target for this refresh: <code>${{summary.reporting_month_target}}</code>`,
      ].join("<br/>");
    }}

    function renderTotalTrend() {{
      const rows = currentWindowRows(dashboardData.monthly_overview);
      buildChart("chartTotalTrend", {{
        type: "line",
        data: {{
          labels: rows.map((row) => row.year_month),
          datasets: [{{
            label: "Transactions",
            data: rows.map((row) => row.transactions),
            borderColor: palette.navy,
            backgroundColor: "rgba(23, 63, 79, 0.16)",
            fill: true,
            tension: 0.3,
            borderWidth: 3,
            pointRadius: 2,
          }}],
        }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function renderAssetMix() {{
      const rows = currentWindowRows(dashboardData.monthly_overview);
      buildChart("chartAssetMix", {{
        type: "bar",
        data: {{
          labels: rows.map((row) => row.year_month),
          datasets: [
            {{
              label: "High-rise",
              data: rows.map((row) => row.high_rise),
              backgroundColor: palette.highRise,
              stack: "mix",
            }},
            {{
              label: "Low-rise",
              data: rows.map((row) => row.low_rise),
              backgroundColor: palette.lowRise,
              stack: "mix",
            }},
            {{
              label: "Landed house",
              data: rows.map((row) => row.landed_house),
              backgroundColor: palette.landed,
              stack: "mix",
            }},
          ],
        }},
        options: {{
          ...baseChartOptions("Transactions"),
          scales: {{
            ...baseScales("Transactions"),
            x: {{ stacked: true, ticks: {{ color: palette.muted }} }},
            y: {{ stacked: true, beginAtZero: true, ticks: {{ color: palette.muted }} }},
          }},
        }},
      }});
    }}

    function renderMoM() {{
      const rows = currentWindowRows(dashboardData.monthly_overview).slice(1);
      buildChart("chartMoM", {{
        type: "bar",
        data: {{
          labels: rows.map((row) => row.year_month),
          datasets: [{{
            label: "MoM delta",
            data: rows.map((row) => row.mom_abs),
            backgroundColor: rows.map((row) => row.mom_abs >= 0 ? "rgba(45, 143, 111, 0.75)" : "rgba(217, 90, 90, 0.75)"),
            borderColor: rows.map((row) => row.mom_abs >= 0 ? palette.green : palette.red),
            borderWidth: 1,
          }}],
        }},
        options: baseChartOptions("Delta"),
      }});
    }}

    function renderMonthlyOverviewTable() {{
      const rows = sortRows(
        "monthlyOverviewTable",
        currentWindowRows(dashboardData.monthly_overview).slice().reverse().slice(0, 18)
      );
      document.getElementById("monthlyOverviewTable").innerHTML = rows.map((row) => `
        <tr>
          <td><strong>${{row.year_month}}</strong></td>
          <td>${{formatInt(row.transactions)}}</td>
          <td class="${{deltaClass(row.mom_abs || 0)}}">${{row.mom_abs === null ? "NA" : `${{row.mom_abs > 0 ? "+" : ""}}${{formatInt(row.mom_abs)}}`}}</td>
          <td>${{formatInt(row.high_rise)}}</td>
          <td>${{formatInt(row.low_rise)}}</td>
          <td>${{formatInt(row.landed_house)}}</td>
        </tr>
      `).join("");
    }}

    function renderGroupCharts() {{
      const group2Series = groupSeries(dashboardData.high_rise_group2, "group_name");
      buildChart("chartGroup2", {{
        type: "line",
        data: {{
          labels: group2Series.months,
          datasets: group2Series.datasets,
        }},
        options: baseChartOptions("Transactions"),
      }});

      const group1Series = groupSeries(dashboardData.high_rise_group1, "group_name");
      buildChart("chartGroup1", {{
        type: "line",
        data: {{
          labels: group1Series.months,
          datasets: group1Series.datasets,
        }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function renderSectorTrend() {{
      const sectorRows = dashboardData.high_rise_sector_monthly.filter(
        (row) => row.focus_group === state.focusGroup
      );
      const months = currentWindowRows(dashboardData.monthly_overview).map((row) => row.year_month);
      const topSectors = topSectorsForFocus();
      const datasets = topSectors.map((sector, index) => {{
        const monthMap = new Map(
          sectorRows
            .filter((row) => row.sector_name === sector)
            .map((row) => [row.year_month, row.transactions])
        );
        return {{
          label: sector,
          data: months.map((month) => monthMap.get(month) || 0),
          borderColor: groupColors[index % groupColors.length],
          backgroundColor: groupColors[index % groupColors.length],
          borderWidth: 2,
          tension: 0.28,
          fill: false,
        }};
      }});

      document.getElementById("sectorTrendTitle").textContent = `${{state.focusGroup}} sector trend`;
      document.getElementById("sectorTrendNote").textContent = `Top sectors by historical transaction count inside ${{state.focusGroup}}.`;
      buildChart("chartSectorTrend", {{
        type: "line",
        data: {{
          labels: months,
          datasets,
        }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function renderSectorTable() {{
      const rows = sortRows("sectorDeltaTable", dashboardData.high_rise_sector_latest[state.focusGroup] || []);
      document.getElementById("sectorTableTitle").textContent = `${{state.focusGroup}} sector delta table`;
      document.getElementById("sectorTableNote").textContent =
        `${{dashboardData.summary.latest_month}} vs ${{dashboardData.summary.previous_month || "previous month"}}`;
      document.getElementById("sectorDeltaTable").innerHTML = rows.map((row) => `
        <tr>
          <td><strong>${{row.sector_name}}</strong></td>
          <td>${{formatInt(row.latest_count)}}</td>
          <td>${{formatInt(row.prev_count)}}</td>
          <td class="${{deltaClass(row.mom_abs || 0)}}">${{row.mom_abs > 0 ? "+" : ""}}${{formatInt(row.mom_abs)}}</td>
          <td class="${{deltaClass(row.mom_abs || 0)}}">${{formatPct(row.mom_pct)}}</td>
          <td>${{formatInt(row.total_transactions)}}</td>
        </tr>
      `).join("");
    }}

    function renderLowRiseProjects() {{
      const months = currentWindowRows(dashboardData.monthly_overview).map((row) => row.year_month);
      const datasets = dashboardData.meta.low_rise_focus_projects.map((project, index) => {{
        const monthMap = new Map(
          dashboardData.low_rise_project_monthly
            .filter((row) => row.project_name === project)
            .map((row) => [row.year_month, row.transactions])
        );
        return {{
          label: project,
          data: months.map((month) => monthMap.get(month) || 0),
          borderColor: groupColors[index % groupColors.length],
          backgroundColor: groupColors[index % groupColors.length],
          borderWidth: 2,
          tension: 0.28,
          fill: false,
        }};
      }});
      buildChart("chartLowRiseProjects", {{
        type: "line",
        data: {{ labels: months, datasets }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function topLowRiseSectors() {{
      const rows = dashboardData.low_rise_sector_monthly.filter(
        (row) => row.project_name === state.lowRiseProject && row.sector_name !== "Unknown"
      );
      const totals = new Map();
      rows.forEach((row) => {{
        totals.set(row.sector_name, row.total_transactions || 0);
      }});
      return [...totals.entries()]
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .slice(0, 8)
        .map(([name]) => name);
    }}

    function renderLowRiseSectorTrend() {{
      const months = currentWindowRows(dashboardData.monthly_overview).map((row) => row.year_month);
      const rows = dashboardData.low_rise_sector_monthly.filter(
        (row) => row.project_name === state.lowRiseProject && row.sector_name !== "Unknown"
      );
      const topSectors = topLowRiseSectors();
      const datasets = topSectors.map((sector, index) => {{
        const monthMap = new Map(
          rows.filter((row) => row.sector_name === sector).map((row) => [row.year_month, row.transactions])
        );
        return {{
          label: sector,
          data: months.map((month) => monthMap.get(month) || 0),
          borderColor: groupColors[index % groupColors.length],
          backgroundColor: groupColors[index % groupColors.length],
          borderWidth: 2,
          tension: 0.28,
          fill: false,
        }};
      }});
      document.getElementById("lowRiseSectorTrendTitle").textContent = `${{state.lowRiseProject}} sector trend`;
      document.getElementById("lowRiseSectorTrendNote").textContent = `Top sectors with non-empty sector_name inside ${{state.lowRiseProject}}.`;
      buildChart("chartLowRiseSectorTrend", {{
        type: "line",
        data: {{ labels: months, datasets }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function renderLowRiseSectorTable() {{
      const rows = sortRows("lowRiseSectorDeltaTable", (dashboardData.low_rise_sector_latest[state.lowRiseProject] || []).filter(
        (row) => row.sector_name !== "Unknown"
      ));
      document.getElementById("lowRiseSectorTableTitle").textContent = `${{state.lowRiseProject}} sector delta`;
      document.getElementById("lowRiseSectorTableNote").textContent =
        `${{dashboardData.summary.latest_month}} vs ${{dashboardData.summary.previous_month || "previous month"}}`;
      document.getElementById("lowRiseSectorDeltaTable").innerHTML = rows.map((row) => `
        <tr>
          <td><strong>${{row.sector_name}}</strong></td>
          <td>${{formatInt(row.latest_count)}}</td>
          <td>${{formatInt(row.prev_count)}}</td>
          <td class="${{deltaClass(row.mom_abs || 0)}}">${{row.mom_abs > 0 ? "+" : ""}}${{formatInt(row.mom_abs)}}</td>
          <td class="${{deltaClass(row.mom_abs || 0)}}">${{formatPct(row.mom_pct)}}</td>
          <td>${{formatInt(row.total_transactions)}}</td>
        </tr>
      `).join("");
    }}

    function topLandedSeries(rows, nameField, topN) {{
      const totals = new Map();
      rows.forEach((row) => {{
        totals.set(row[nameField], row.total_transactions || 0);
      }});
      return [...totals.entries()]
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .slice(0, topN)
        .map(([name]) => name);
    }}

    function landedWardRowsForDistrict() {{
      return dashboardData.landed_ward_monthly.filter((row) => row.district === state.landedDistrict);
    }}

    function renderLandedDistrictTrend() {{
      const months = currentWindowRows(dashboardData.monthly_overview).map((row) => row.year_month);
      const topDistricts = topLandedSeries(dashboardData.landed_district_monthly, "district", 8);
      const datasets = topDistricts.map((district, index) => {{
        const monthMap = new Map(
          dashboardData.landed_district_monthly
            .filter((row) => row.district === district)
            .map((row) => [row.year_month, row.transactions])
        );
        return {{
          label: district,
          data: months.map((month) => monthMap.get(month) || 0),
          borderColor: groupColors[index % groupColors.length],
          backgroundColor: groupColors[index % groupColors.length],
          borderWidth: 2,
          tension: 0.28,
          fill: false,
        }};
      }});
      buildChart("chartLandedDistrictTrend", {{
        type: "line",
        data: {{ labels: months, datasets }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function renderLongBienWardTrend() {{
      const months = currentWindowRows(dashboardData.monthly_overview).map((row) => row.year_month);
      const wardRows = landedWardRowsForDistrict();
      const topWards = topLandedSeries(wardRows, "ward", 8);
      const datasets = topWards.map((ward, index) => {{
        const monthMap = new Map(
          wardRows
            .filter((row) => row.ward === ward)
            .map((row) => [row.year_month, row.transactions])
        );
        return {{
          label: ward,
          data: months.map((month) => monthMap.get(month) || 0),
          borderColor: groupColors[index % groupColors.length],
          backgroundColor: groupColors[index % groupColors.length],
          borderWidth: 2,
          tension: 0.28,
          fill: false,
        }};
      }});
      document.getElementById("landedWardTrendTitle").textContent = `${{state.landedDistrict}} ward trend`;
      document.getElementById("landedWardTrendNote").textContent =
        `Top wards in ${{state.landedDistrict}} by historical landed-house transaction count.`;
      buildChart("chartLongBienWardTrend", {{
        type: "line",
        data: {{ labels: months, datasets }},
        options: baseChartOptions("Transactions"),
      }});
    }}

    function renderLandedTables() {{
      document.getElementById("landedDistrictDeltaTable").innerHTML =
        sortRows("landedDistrictDeltaTable", dashboardData.landed_district_latest || []).map((row) => `
          <tr>
            <td><strong>${{row.district}}</strong></td>
            <td>${{formatInt(row.latest_count)}}</td>
            <td>${{formatInt(row.prev_count)}}</td>
            <td class="${{deltaClass(row.mom_abs || 0)}}">${{row.mom_abs > 0 ? "+" : ""}}${{formatInt(row.mom_abs)}}</td>
            <td class="${{deltaClass(row.mom_abs || 0)}}">${{formatPct(row.mom_pct)}}</td>
            <td>${{formatInt(row.total_transactions)}}</td>
          </tr>
        `).join("");
      const wardRows = landedWardRowsForDistrict();
      const latestWardRows = wardRows
        .filter((row) => row.year_month === dashboardData.summary.latest_month)
        .map((row) => ({{
          ward: row.ward,
          latest_count: row.transactions,
          total_transactions: row.total_transactions,
        }}));
      const previousWardMap = new Map(
        wardRows
          .filter((row) => row.year_month === (dashboardData.summary.previous_month || ""))
          .map((row) => [row.ward, row.transactions])
      );
      const wardLatestMerged = latestWardRows.map((row) => {{
        const prevCount = previousWardMap.get(row.ward) || 0;
        const momAbs = row.latest_count - prevCount;
        return {{
          ...row,
          prev_count: prevCount,
          mom_abs: momAbs,
          mom_pct: prevCount ? (momAbs / prevCount) * 100 : null,
        }};
      }});
      document.getElementById("landedWardTableTitle").textContent = `${{state.landedDistrict}} ward delta table`;
      document.getElementById("landedWardTableNote").textContent =
        `${{dashboardData.summary.latest_month}} vs ${{dashboardData.summary.previous_month || "previous month"}}`;
      document.getElementById("longBienWardDeltaTable").innerHTML =
        sortRows("longBienWardDeltaTable", wardLatestMerged).map((row) => `
          <tr>
            <td><strong>${{row.ward}}</strong></td>
            <td>${{formatInt(row.latest_count)}}</td>
            <td>${{formatInt(row.prev_count)}}</td>
            <td class="${{deltaClass(row.mom_abs || 0)}}">${{row.mom_abs > 0 ? "+" : ""}}${{formatInt(row.mom_abs)}}</td>
            <td class="${{deltaClass(row.mom_abs || 0)}}">${{formatPct(row.mom_pct)}}</td>
            <td>${{formatInt(row.total_transactions)}}</td>
          </tr>
        `).join("");
    }}

    function baseScales(yTitle) {{
      return {{
        x: {{
          ticks: {{ color: palette.muted }},
          grid: {{ color: "rgba(109, 122, 129, 0.12)" }},
        }},
        y: {{
          beginAtZero: true,
          ticks: {{ color: palette.muted }},
          title: {{
            display: true,
            text: yTitle,
            color: palette.muted,
            font: {{ weight: "700" }},
          }},
          grid: {{ color: "rgba(109, 122, 129, 0.12)" }},
        }},
      }};
    }}

    function baseChartOptions(yTitle) {{
      return {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{ mode: "index", intersect: false }},
        plugins: {{
          legend: {{
            labels: {{
              usePointStyle: true,
              boxWidth: 10,
              color: palette.ink,
              font: {{ weight: "700" }},
            }},
          }},
          tooltip: {{
            callbacks: {{
              label: (context) => `${{context.dataset.label}}: ${{formatInt(context.parsed.y)}}`,
            }},
          }},
        }},
        scales: baseScales(yTitle),
      }};
    }}

    function renderAll() {{
      renderSummary();
      renderTotalTrend();
      renderAssetMix();
      renderMoM();
      renderMonthlyOverviewTable();
      renderGroupCharts();
      renderSectorTrend();
      renderSectorTable();
      renderLowRiseProjects();
      renderLowRiseSectorTrend();
      renderLowRiseSectorTable();
      renderLandedDistrictTrend();
      renderLongBienWardTrend();
      renderLandedTables();
      refreshSortIndicators();
    }}

    function bindSortableHeaders() {{
      document.querySelectorAll("th.sortable").forEach((header) => {{
        header.addEventListener("click", () => {{
          const tableId = header.dataset.table;
          const key = header.dataset.key;
          const type = header.dataset.type || "text";
          const existing = state.tableSorts[tableId];
          if (existing && existing.key === key) {{
            state.tableSorts[tableId] = {{
              key,
              type,
              dir: existing.dir === "asc" ? "desc" : "asc",
            }};
          }} else {{
            state.tableSorts[tableId] = {{
              key,
              type,
              dir: type === "text" ? "asc" : "desc",
            }};
          }}
          renderAll();
        }});
      }});
    }}

    function bindControls() {{
      const startMonthSelect = document.getElementById("startMonthSelect");
      const endMonthSelect = document.getElementById("endMonthSelect");
      const applyWindowButton = document.getElementById("applyWindowButton");
      const monthOptions = allMonths.map((month) => `<option value="${{month}}">${{month}}</option>`).join("");
      startMonthSelect.innerHTML = monthOptions;
      endMonthSelect.innerHTML = monthOptions;
      startMonthSelect.value = state.pendingStartMonth;
      endMonthSelect.value = state.pendingEndMonth;
      startMonthSelect.addEventListener("change", (event) => {{
        state.pendingStartMonth = event.target.value;
        if (state.pendingEndMonth < state.pendingStartMonth) {{
          state.pendingEndMonth = state.pendingStartMonth;
          endMonthSelect.value = state.pendingEndMonth;
        }}
      }});
      endMonthSelect.addEventListener("change", (event) => {{
        state.pendingEndMonth = event.target.value;
        if (state.pendingStartMonth > state.pendingEndMonth) {{
          state.pendingStartMonth = state.pendingEndMonth;
          startMonthSelect.value = state.pendingStartMonth;
        }}
      }});
      applyWindowButton.addEventListener("click", () => {{
        state.startMonth = state.pendingStartMonth;
        state.endMonth = state.pendingEndMonth;
        renderAll();
      }});

      const focusSelect = document.getElementById("focusGroupSelect");
      focusSelect.innerHTML = dashboardData.meta.focus_groups
        .map((group) => `<option value="${{group}}">${{group}}</option>`)
        .join("");
      focusSelect.value = state.focusGroup;
      focusSelect.addEventListener("change", (event) => {{
        state.focusGroup = event.target.value;
        renderAll();
      }});

      const lowRiseSelect = document.getElementById("lowRiseProjectSelect");
      lowRiseSelect.innerHTML = dashboardData.meta.low_rise_focus_projects
        .map((project) => `<option value="${{project}}">${{project}}</option>`)
        .join("");
      lowRiseSelect.value = state.lowRiseProject;
      lowRiseSelect.addEventListener("change", (event) => {{
        state.lowRiseProject = event.target.value;
        renderAll();
      }});

      const landedDistrictSelect = document.getElementById("landedDistrictSelect");
      landedDistrictSelect.innerHTML = (dashboardData.meta.landed_districts || [])
        .map((district) => `<option value="${{district}}">${{district}}</option>`)
        .join("");
      landedDistrictSelect.value = state.landedDistrict;
      landedDistrictSelect.addEventListener("change", (event) => {{
        state.landedDistrict = event.target.value;
        renderAll();
      }});
    }}

    bindSortableHeaders();
    bindControls();
    renderAll();
  </script>
</body>
</html>
"""


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    dataset = build_dataset()
    OUTPUT_JSON.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    OUTPUT_HTML.write_text(build_html(dataset), encoding="utf-8")

    if PUBLISH_DIR.exists():
      PUBLISH_JSON.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
      PUBLISH_HTML.write_text(build_html(dataset), encoding="utf-8")

    print(f"Wrote dashboard data to {OUTPUT_JSON}")
    print(f"Wrote dashboard HTML to {OUTPUT_HTML}")
    if PUBLISH_DIR.exists():
        print(f"Wrote publish dashboard data to {PUBLISH_JSON}")
        print(f"Wrote publish dashboard HTML to {PUBLISH_HTML}")


if __name__ == "__main__":
    main()
