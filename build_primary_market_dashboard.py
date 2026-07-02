from __future__ import annotations

import json
import re
import subprocess
import unicodedata
import urllib.request
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
SHORTCUT_PATH = ROOT / "01. Database_all - Shortcut.lnk"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
OUTPUT_HTML = OUTPUT_DIR / "primary_market_dashboard.html"
OUTPUT_JSON = OUTPUT_DIR / "primary_market_dashboard_data.json"
BOUNDARY_CACHE = Path(__file__).resolve().parent / "gadm41_vnm_1.json"
GADM_VNM_LEVEL1_URL = "https://geodata.ucdavis.edu/gadm/gadm4.1/json/gadm41_VNM_1.json"

PROJECT_COORDINATE_OVERRIDES = {
    # Workbook currently places Iconia Lakeside far north of the Tố Hữu corridor.
    # Use an explicit dashboard override so the project appears in the Hà Đông /
    # Lương Thế Vinh area until the source workbook is corrected.
    ("Hanoi", "882"): {
        "project_latitude": 20.985659,
        "project_longitude": 105.777720,
    },
}


WORKBOOK_CONFIGS = (
    {
        "city": "Hanoi",
        "folder": "1. Hanoi",
        "file": "Hanoi High-rise Primary Database.xlsx",
        "sheet": "3.Time Series",
        "header_row": 2,
        "project_sheet": "2.Project identification",
        "project_sheet_header_row": 3,
        "project_id_col": "Project ID",
        "project_name_col": "Project Name",
        "quarter_col": "Quarter",
        "segment_col": "District",
        "grading_col": "Grading_utd",
        "region_col": "Region",
        "new_launched_col": "New launched",
        "new_sold_col": "New Sold",
        "price_col": "Current VND Selling Price",
        "current_supply_col": "Current supply",
        "available_begin_col": "Available for sale at the beginning",
        "available_end_col": "Available for sale at the end",
        "current_sold_out_col": "Current sold-out",
        "status_col": "SaleStatus",
        "new_project_col": "new_project",
        "segment_map": {
            "van giang": "Van Giang HY",
        },
        "default_segment": "Central Hanoi",
    },
    {
        "city": "HCMC",
        "folder": "2. HCM",
        "file": "HCMC High-rise Primary Database.xlsx",
        "sheet": "Time Series",
        "header_row": 1,
        "project_sheet": "Project Identification",
        "project_sheet_header_row": 2,
        "project_id_col": "Project ID",
        "project_name_col": "Project Name",
        "quarter_col": "Quarter",
        "segment_col": "Province",
        "grading_col": "Grading",
        "region_col": "Region",
        "new_launched_col": "New launched",
        "new_sold_col": "New sold",
        "price_col": "Current VND Selling Price",
        "current_supply_col": "Current supply",
        "available_begin_col": "Available for sale at the beginning",
        "available_end_col": "Available for sale at the end",
        "current_sold_out_col": "Current sold-out",
        "status_col": "Sale Status",
        "new_project_col": None,
        "segment_map": {
            "hcmc": "Central HCMC",
            "binh duong": "Binh Duong",
            "br-vt": "BR-VT",
        },
        "default_segment": "Other",
    },
)


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


def quarter_key(value: str) -> int:
    match = re.fullmatch(r"(\d{4})Q([1-4])", str(value).strip())
    if not match:
        return -1
    year, quarter = match.groups()
    return int(year) * 10 + int(quarter)


def normalize_text(value: Any) -> str:
    text = str(value).strip().lower()
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(char for char in normalized if not unicodedata.combining(char))


def normalize_project_id(value: Any) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    return text


def normalize_segment(raw_value: Any, config: dict[str, Any]) -> str:
    text = str(raw_value).strip()
    lowered = normalize_text(text)
    return config["segment_map"].get(lowered, config["default_segment"])


def normalize_dimension(raw_value: Any, dimension_type: str) -> str:
    text = str(raw_value).strip()
    if not text or text.lower() == "nan" or text == "-":
        return "Unknown"
    if dimension_type == "grading":
        lowered = normalize_text(text)
        grading_map = {
            "mid-end": "Mid-end",
            "mid end": "Mid-end",
            "premium": "Premium",
            "luxury": "Luxury",
            "ultra luxury": "Ultra Luxury",
            "affordable": "Affordable",
        }
        return grading_map.get(lowered, text)
    if dimension_type == "developer":
        return text
    return text


def load_project_identification(base_dir: Path, config: dict[str, Any]) -> pd.DataFrame:
    workbook_path = base_dir / config["folder"] / config["file"]
    dataframe = pd.read_excel(
        workbook_path,
        sheet_name=config["project_sheet"],
        header=config["project_sheet_header_row"] - 1,
    )
    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    dataframe = dataframe[
        [
            "Project ID",
            "Project Name",
            "Latitude",
            "Longitude",
            "Current Developer (Main Entity)",
        ]
    ].copy()
    dataframe["project_id"] = dataframe["Project ID"].map(normalize_project_id)
    dataframe = dataframe[dataframe["project_id"].notna()].copy()
    dataframe["project_latitude"] = pd.to_numeric(dataframe["Latitude"], errors="coerce")
    dataframe["project_longitude"] = pd.to_numeric(dataframe["Longitude"], errors="coerce")
    dataframe["project_name_lookup"] = dataframe["Project Name"].astype(str).str.strip()
    dataframe["project_developer"] = dataframe["Current Developer (Main Entity)"].map(
        lambda value: normalize_dimension(value, "developer")
    )
    return dataframe[
        [
            "project_id",
            "project_name_lookup",
            "project_latitude",
            "project_longitude",
            "project_developer",
        ]
    ].drop_duplicates(subset=["project_id"])


def load_city_rows(base_dir: Path, config: dict[str, Any]) -> pd.DataFrame:
    workbook_path = base_dir / config["folder"] / config["file"]
    project_lookup = load_project_identification(base_dir, config)
    dataframe = pd.read_excel(
        workbook_path,
        sheet_name=config["sheet"],
        header=config["header_row"] - 1,
    )
    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    dataframe = dataframe[
        [
            config["project_id_col"],
            config["project_name_col"],
            config["quarter_col"],
            config["segment_col"],
            config["grading_col"],
            config["region_col"],
            config["new_launched_col"],
            config["new_sold_col"],
            config["price_col"],
            config["current_supply_col"],
            config["available_begin_col"],
            config["available_end_col"],
            config["current_sold_out_col"],
            config["status_col"],
            *([config["new_project_col"]] if config["new_project_col"] else []),
        ]
    ].copy()

    dataframe["city"] = config["city"]
    dataframe["project_name"] = dataframe[config["project_name_col"]].astype(str).str.strip()
    dataframe["project_id"] = dataframe[config["project_id_col"]].map(normalize_project_id)
    dataframe["quarter"] = dataframe[config["quarter_col"]].astype(str).str.strip()
    dataframe["quarter_sort"] = dataframe["quarter"].map(quarter_key)
    dataframe = dataframe[dataframe["quarter_sort"] > 0].copy()
    dataframe["segment"] = dataframe[config["segment_col"]].map(
        lambda value: normalize_segment(value, config)
    )
    dataframe["grading"] = dataframe[config["grading_col"]].map(
        lambda value: normalize_dimension(value, "grading")
    )
    dataframe["region"] = dataframe[config["region_col"]].map(
        lambda value: normalize_dimension(value, "region")
    )
    dataframe["developer"] = "Unknown"
    dataframe["sale_status"] = (
        dataframe[config["status_col"]].astype(str).str.strip().str.lower()
    )
    dataframe = dataframe[dataframe["sale_status"] == "current"].copy()

    for source_col, target_col in (
        (config["new_launched_col"], "new_launched"),
        (config["new_sold_col"], "new_sold"),
        (config["price_col"], "price"),
        (config["current_supply_col"], "current_supply"),
        (config["available_begin_col"], "available_begin"),
        (config["available_end_col"], "available_end"),
        (config["current_sold_out_col"], "current_sold_out"),
    ):
        dataframe[target_col] = pd.to_numeric(dataframe[source_col], errors="coerce")

    if config["new_project_col"]:
        dataframe["new_project_marker"] = dataframe[config["new_project_col"]]
    else:
        dataframe["new_project_marker"] = pd.NA

    dataframe["new_launched"] = dataframe["new_launched"].fillna(0)
    dataframe["new_sold"] = dataframe["new_sold"].fillna(0)

    dataframe["is_actual_project"] = dataframe["project_id"].notna() & ~dataframe[
        "project_name"
    ].str.lower().str.startswith("total")

    dataframe["project_key"] = dataframe.apply(
        lambda row: f"{row['city']}::{row['project_id']}"
        if row["is_actual_project"]
        else None,
        axis=1,
    )
    dataframe = dataframe.merge(
        project_lookup,
        on="project_id",
        how="left",
    )
    dataframe["developer"] = dataframe["project_developer"].combine_first(dataframe["developer"])
    override_mask = dataframe["project_id"].notna()
    if override_mask.any():
        for (city, project_id), coordinates in PROJECT_COORDINATE_OVERRIDES.items():
            project_mask = (dataframe["city"] == city) & (dataframe["project_id"] == project_id)
            if project_mask.any():
                dataframe.loc[project_mask, "project_latitude"] = coordinates["project_latitude"]
                dataframe.loc[project_mask, "project_longitude"] = coordinates["project_longitude"]
    return dataframe


def attach_first_launch_flags(dataframe: pd.DataFrame) -> pd.DataFrame:
    actual = dataframe[
        dataframe["is_actual_project"] & (dataframe["new_launched"] > 0)
    ][["project_key", "quarter_sort", "quarter"]].copy()
    if actual.empty:
        dataframe["is_new_project"] = False
        return dataframe

    first_launch = (
        actual.sort_values(["project_key", "quarter_sort"])
        .groupby("project_key", as_index=False)
        .first()
        .rename(columns={"quarter": "first_launch_quarter"})
    )
    dataframe = dataframe.merge(first_launch[["project_key", "first_launch_quarter"]], on="project_key", how="left")
    dataframe["is_new_project"] = dataframe["quarter"] == dataframe["first_launch_quarter"]
    return dataframe


def load_city_boundaries() -> dict[str, Any]:
    if BOUNDARY_CACHE.exists():
        data = json.loads(BOUNDARY_CACHE.read_text(encoding="utf-8"))
    else:
        with urllib.request.urlopen(GADM_VNM_LEVEL1_URL, timeout=60) as response:
            data = json.load(response)
        BOUNDARY_CACHE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    boundary_map: dict[str, Any] = {}
    for feature in data.get("features", []):
        name = normalize_text(feature.get("properties", {}).get("NAME_1", ""))
        if name == "hanoi":
            boundary_map["Hanoi"] = feature.get("geometry")
        elif name == "hochiminh":
            boundary_map["HCMC"] = feature.get("geometry")
    return boundary_map


def summarize_scope(dataframe: pd.DataFrame, city: str, segment: str) -> list[dict[str, Any]]:
    scope_rows = dataframe.copy()

    base_metrics = (
        scope_rows.groupby("quarter", as_index=False)
        .agg(
            new_launched=("new_launched", "sum"),
            new_sold=("new_sold", "sum"),
        )
    )

    price_current = scope_rows[
        scope_rows["price"].notna()
        & (scope_rows["price"] > 0)
        & scope_rows["current_supply"].notna()
        & (scope_rows["current_supply"] > 0)
    ].copy()
    if not price_current.empty:
        price_current["weighted_value"] = price_current["price"] * price_current["current_supply"]
        current_price = price_current.groupby("quarter", as_index=False).agg(
            current_supply_weight=("current_supply", "sum"),
            current_supply_weighted_price=("weighted_value", "sum"),
        )
        current_price["weighted_price_current_supply"] = (
            current_price["current_supply_weighted_price"] / current_price["current_supply_weight"]
        )
        current_price = current_price.drop(columns=["current_supply_weighted_price"])
    else:
        current_price = pd.DataFrame(columns=["quarter", "current_supply_weight", "weighted_price_current_supply"])

    price_available = scope_rows[
        scope_rows["price"].notna()
        & (scope_rows["price"] > 0)
        & scope_rows["available_begin"].notna()
        & (scope_rows["available_begin"] > 0)
    ].copy()
    if not price_available.empty:
        price_available["weighted_value"] = price_available["price"] * price_available["available_begin"]
        available_price = price_available.groupby("quarter", as_index=False).agg(
            available_begin=("available_begin", "sum"),
            available_weighted_price=("weighted_value", "sum"),
        )
        available_price["weighted_price_available_supply"] = (
            available_price["available_weighted_price"] / available_price["available_begin"]
        )
        available_price = available_price.drop(columns=["available_weighted_price"])
    else:
        available_price = pd.DataFrame(columns=["quarter", "available_begin", "weighted_price_available_supply"])

    actual_projects = scope_rows[scope_rows["is_actual_project"]].copy()
    open_projects = (
        actual_projects.groupby("quarter", as_index=False)
        .agg(active_projects=("project_id", "nunique"))
        if not actual_projects.empty
        else pd.DataFrame(columns=["quarter", "active_projects"])
    )

    new_projects = actual_projects[actual_projects["is_new_project"]].copy()
    new_projects_count = (
        new_projects.groupby("quarter", as_index=False)
        .agg(new_projects=("project_key", "nunique"))
        if not new_projects.empty
        else pd.DataFrame(columns=["quarter", "new_projects"])
    )

    new_project_prices = new_projects[
        new_projects["price"].notna() & (new_projects["price"] > 0)
    ].copy()
    if not new_project_prices.empty:
        new_project_price = (
            new_project_prices.groupby(["quarter", "project_key"], as_index=False)
            .agg(project_price=("price", "mean"))
            .groupby("quarter", as_index=False)
            .agg(new_project_avg_price=("project_price", "mean"))
        )
    else:
        new_project_price = pd.DataFrame(columns=["quarter", "new_project_avg_price"])

    summary = base_metrics
    for extra in (
        current_price,
        available_price,
        new_project_price,
        open_projects,
        new_projects_count,
    ):
        summary = summary.merge(extra, on="quarter", how="left")

    for column in (
        "current_supply_weight",
        "available_begin",
        "weighted_price_current_supply",
        "weighted_price_available_supply",
        "new_project_avg_price",
        "active_projects",
        "new_projects",
        "current_sold_out",
        "current_supply",
        "available_end",
    ):
        if column not in summary.columns:
            summary[column] = pd.NA

    supply_rollup = scope_rows.groupby("quarter", as_index=False).agg(
        current_supply=("current_supply", "sum"),
        current_sold_out=("current_sold_out", "sum"),
        available_begin=("available_begin", "sum"),
        available_end=("available_end", "sum"),
    )
    summary = summary.merge(supply_rollup, on="quarter", how="left", suffixes=("", "_rollup"))
    for column in ("current_supply", "current_sold_out", "available_begin", "available_end"):
        rolled = f"{column}_rollup"
        if rolled in summary.columns:
            summary[column] = summary[rolled].combine_first(summary[column])
            summary = summary.drop(columns=[rolled])

    summary["sold_rate_quarterly"] = pd.NA
    quarterly_denominator = summary["available_begin"]
    quarterly_mask = quarterly_denominator.notna() & (quarterly_denominator > 0)
    summary.loc[quarterly_mask, "sold_rate_quarterly"] = (
        summary.loc[quarterly_mask, "new_sold"] / quarterly_denominator.loc[quarterly_mask] * 100
    )

    summary["prev_available_end"] = summary["available_end"].shift(1)
    summary["sold_rate_cumulative"] = pd.NA
    cumulative_denominator = summary["current_supply"] + summary["prev_available_end"]
    cumulative_mask = cumulative_denominator.notna() & (cumulative_denominator > 0)
    summary.loc[cumulative_mask, "sold_rate_cumulative"] = (
        summary.loc[cumulative_mask, "current_sold_out"] / cumulative_denominator.loc[cumulative_mask] * 100
    )

    masterise_rows = scope_rows[
        scope_rows["developer"].astype(str).str.strip().str.lower() == "masterise homes"
    ].copy()
    if not masterise_rows.empty:
        masterise_sold = masterise_rows.groupby("quarter", as_index=False).agg(
            masterise_new_sold=("new_sold", "sum")
        )
        summary = summary.merge(masterise_sold, on="quarter", how="left")
    else:
        summary["masterise_new_sold"] = pd.NA
    summary["masterise_market_share"] = pd.NA
    masterise_mask = summary["new_sold"].notna() & (summary["new_sold"] > 0)
    summary.loc[masterise_mask, "masterise_market_share"] = (
        summary.loc[masterise_mask, "masterise_new_sold"].fillna(0)
        / summary.loc[masterise_mask, "new_sold"]
        * 100
    )

    summary["city"] = city
    summary["segment"] = segment
    summary["quarter_sort"] = summary["quarter"].map(quarter_key)
    summary["year"] = summary["quarter"].str.extract(r"(\d{4})").astype(int)
    summary["quarter_no"] = summary["quarter"].str.extract(r"Q([1-4])").astype(int)
    summary = summary.sort_values("quarter_sort").reset_index(drop=True)
    return json.loads(summary.to_json(orient="records", force_ascii=False))


def summarize_dimension_breakdown(
    dataframe: pd.DataFrame,
    city: str,
    segment: str,
    dimension_type: str,
) -> list[dict[str, Any]]:
    actual = dataframe[dataframe["is_actual_project"]].copy()
    if actual.empty:
        return []

    dimension_rows = (
        actual.groupby(["quarter", dimension_type], as_index=False)
        .agg(
            new_launched=("new_launched", "sum"),
            new_sold=("new_sold", "sum"),
        )
        .rename(columns={dimension_type: "dimension_value"})
    )
    if dimension_rows.empty:
        return []

    dimension_rows["city"] = city
    dimension_rows["segment"] = segment
    dimension_rows["dimension_type"] = dimension_type
    dimension_rows["quarter_sort"] = dimension_rows["quarter"].map(quarter_key)
    dimension_rows = dimension_rows.sort_values(["quarter_sort", "dimension_value"]).reset_index(drop=True)
    return json.loads(dimension_rows.to_json(orient="records", force_ascii=False))


def build_dataset() -> dict[str, Any]:
    base_dir = resolve_shortcut(SHORTCUT_PATH)
    city_frames = []
    for config in WORKBOOK_CONFIGS:
        city_frames.append(attach_first_launch_flags(load_city_rows(base_dir, config)))

    dataset: list[dict[str, Any]] = []
    dimension_breakdowns: list[dict[str, Any]] = []
    segment_map: dict[str, list[str]] = {}
    map_points: list[dict[str, Any]] = []
    for city_frame in city_frames:
        city_name = city_frame["city"].iloc[0]
        segments = sorted(city_frame["segment"].dropna().unique().tolist())
        segment_map[city_name] = ["Total", *segments]

        dataset.extend(summarize_scope(city_frame, city_name, "Total"))
        dimension_breakdowns.extend(
            summarize_dimension_breakdown(city_frame, city_name, "Total", "grading")
        )
        dimension_breakdowns.extend(
            summarize_dimension_breakdown(city_frame, city_name, "Total", "region")
        )
        dimension_breakdowns.extend(
            summarize_dimension_breakdown(city_frame, city_name, "Total", "developer")
        )
        for segment in segments:
            scoped_frame = city_frame[city_frame["segment"] == segment].copy()
            dataset.extend(
                summarize_scope(
                    scoped_frame,
                    city_name,
                    segment,
                )
            )
            dimension_breakdowns.extend(
                summarize_dimension_breakdown(scoped_frame, city_name, segment, "grading")
            )
            dimension_breakdowns.extend(
                summarize_dimension_breakdown(scoped_frame, city_name, segment, "region")
            )
            dimension_breakdowns.extend(
                summarize_dimension_breakdown(scoped_frame, city_name, segment, "developer")
            )

        point_rows = city_frame[
            city_frame["is_actual_project"]
            & city_frame["project_latitude"].notna()
            & city_frame["project_longitude"].notna()
        ][
            [
                "city",
                "segment",
                "quarter",
                "quarter_sort",
                "project_id",
                "project_name",
                "developer",
                "project_latitude",
                "project_longitude",
                "price",
                "current_supply",
                "available_begin",
                "available_end",
                "current_sold_out",
                "new_launched",
                "new_sold",
                "sale_status",
            ]
        ].copy()
        point_rows = point_rows.rename(
            columns={
                "project_latitude": "latitude",
                "project_longitude": "longitude",
            }
        )
        map_points.extend(
            json.loads(point_rows.to_json(orient="records", force_ascii=False))
        )

    quarter_list = sorted(
        {record["quarter"] for record in dataset},
        key=quarter_key,
    )

    return {
        "records": dataset,
        "dimension_breakdowns": dimension_breakdowns,
        "map_points": map_points,
        "city_boundaries": load_city_boundaries(),
        "segments": segment_map,
        "markets": [config["city"] for config in WORKBOOK_CONFIGS],
        "quarters": quarter_list,
        "generated_from": str(SHORTCUT_PATH),
    }


def build_html(dataset: dict[str, Any]) -> str:
    data_json = json.dumps(dataset, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Primary Market Comparison Dashboard</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
  <link
    rel="stylesheet"
    href="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css"
  />
  <script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    :root {{
      --bg: #f6f8fb;
      --panel: rgba(255, 255, 255, 0.96);
      --panel-2: rgba(245, 249, 255, 0.98);
      --ink: #293537;
      --muted: #5f6b6d;
      --line: rgba(161, 198, 250, 0.42);
      --blue: #3281f5;
      --blue-soft: #a1c6fa;
      --teal: #3599b8;
      --green: #68b794;
      --red: #dc707a;
      --navy: #1b4d5c;
      --gold: #b59525;
      --hanoi: #1b4d5c;
      --hcmc: #3281f5;
      --shadow: 0 20px 44px rgba(27, 77, 92, 0.08);
      --radius: 24px;
    }}
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: "Manrope", "Segoe UI", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(161, 198, 250, 0.24), transparent 26%),
        radial-gradient(circle at top right, rgba(104, 183, 148, 0.18), transparent 22%),
        linear-gradient(180deg, #ffffff 0%, var(--bg) 48%, #eef4fb 100%);
    }}
    .shell {{
      max-width: 1540px;
      margin: 0 auto;
      padding: 28px 24px 44px;
    }}
    .hero,
    .compare-grid,
    .price-grid,
    .deep-grid {{
      display: grid;
      gap: 18px;
    }}
    .hero {{
      grid-template-columns: minmax(0, 1.25fr) minmax(340px, 0.75fr);
      margin-bottom: 22px;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid rgba(215, 203, 185, 0.9);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      backdrop-filter: blur(8px);
    }}
    .hero-card {{
      padding: 28px;
      position: relative;
      overflow: hidden;
    }}
    .hero-card::after {{
      content: "";
      position: absolute;
      right: -36px;
      bottom: -48px;
      width: 240px;
      height: 240px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(50, 129, 245, 0.14), transparent 70%);
    }}
    h1 {{
      margin: 0 0 12px;
      font-size: clamp(32px, 4vw, 52px);
      line-height: 0.94;
      letter-spacing: -0.05em;
    }}
    .subhead {{
      max-width: 820px;
      margin: 0;
      color: var(--muted);
      font-size: 15px;
      line-height: 1.65;
    }}
    .hero-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 18px;
    }}
    .chip {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      border-radius: 999px;
      padding: 9px 14px;
      background: var(--panel-2);
      color: var(--navy);
      font-size: 13px;
      font-weight: 600;
    }}
    .control-card {{
      padding: 22px;
      display: flex;
      flex-direction: column;
      gap: 18px;
    }}
    .control-group {{
      display: flex;
      flex-direction: column;
      gap: 10px;
    }}
    .control-label {{
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--muted);
      font-weight: 700;
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
    select.control-select {{
      width: 100%;
      border: 1px solid var(--line);
      background: #fff;
      color: var(--ink);
      border-radius: 14px;
      padding: 10px 12px;
      font-size: 14px;
      font-weight: 600;
    }}
    button.toggle {{
      border: 1px solid var(--line);
      background: #fff;
      color: var(--ink);
      border-radius: 999px;
      padding: 10px 14px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: transform 0.18s ease, background 0.18s ease, border-color 0.18s ease;
    }}
    button.toggle:hover {{
      transform: translateY(-1px);
      border-color: var(--teal);
    }}
    button.toggle.active {{
      background: var(--navy);
      border-color: var(--navy);
      color: #fff;
      box-shadow: 0 12px 24px rgba(27, 77, 92, 0.18);
    }}
    button.refresh-button {{
      border: 1px solid var(--blue);
      background: var(--blue);
      color: #fff;
      border-radius: 14px;
      padding: 12px 14px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      transition: transform 0.18s ease, opacity 0.18s ease;
    }}
    button.refresh-button:hover {{
      transform: translateY(-1px);
      opacity: 0.92;
    }}
    .summary-bar {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }}
    .summary-card {{
      padding: 18px 18px 16px;
      min-height: 144px;
    }}
    .summary-card h3 {{
      margin: 0 0 10px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--muted);
    }}
    .summary-values {{
      display: grid;
      gap: 8px;
    }}
    .summary-line {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: baseline;
      font-size: 14px;
    }}
    .delta-line {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      color: var(--muted);
      font-size: 12px;
    }}
    .delta-value {{
      font-weight: 700;
    }}
    .delta-value.up {{
      color: var(--green);
    }}
    .delta-value.down {{
      color: var(--red);
    }}
    .delta-value.flat {{
      color: var(--muted);
    }}
    .summary-line strong {{
      font-size: 24px;
      letter-spacing: -0.03em;
    }}
    .summary-foot {{
      margin-top: 10px;
      font-size: 13px;
      color: var(--muted);
    }}
    .section-block {{
      margin-bottom: 24px;
    }}
    .section-head {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: end;
      margin-bottom: 12px;
    }}
    .section-head h2 {{
      margin: 0;
      font-size: 22px;
      letter-spacing: -0.03em;
    }}
    .section-note {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }}
    .compare-grid {{
      grid-template-columns: minmax(0, 1.5fr) minmax(320px, 0.5fr);
    }}
    .price-grid,
    .deep-grid {{
      grid-template-columns: repeat(3, minmax(0, 1fr));
    }}
    .chart-card,
    .table-card {{
      padding: 20px;
    }}
    .chart-head {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: baseline;
      margin-bottom: 12px;
    }}
    .panel[data-export-key] {{
      position: relative;
    }}
    .panel[data-export-key] .chart-head {{
      padding-right: 116px;
    }}
    .export-button {{
      position: absolute;
      top: 18px;
      right: 18px;
      border: 1px solid rgba(76, 135, 204, 0.24);
      background: var(--panel-2);
      color: var(--ink);
      border-radius: 999px;
      padding: 7px 12px;
      font-size: 0.75rem;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      cursor: pointer;
      transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
      z-index: 2;
    }}
    .export-button:hover {{
      transform: translateY(-1px);
      border-color: rgba(58, 124, 235, 0.36);
      box-shadow: 0 10px 20px rgba(58, 124, 235, 0.12);
    }}
    .chart-title {{
      margin: 0;
      font-size: 18px;
      letter-spacing: -0.02em;
    }}
    .chart-note {{
      color: var(--muted);
      font-size: 13px;
    }}
    .canvas-wrap {{
      height: 320px;
    }}
    .canvas-wrap.small {{
      height: 260px;
    }}
    .legend-list {{
      display: grid;
      gap: 12px;
      margin-top: 8px;
    }}
    .legend-item {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 12px 14px;
      border-radius: 16px;
      background: var(--panel-2);
      font-size: 14px;
    }}
    .legend-swatch {{
      display: inline-block;
      width: 12px;
      height: 12px;
      border-radius: 50%;
      margin-right: 8px;
    }}
    .deep-grid .panel {{
      min-height: 360px;
    }}
    .map-shell {{
      display: grid;
      grid-template-columns: minmax(0, 1.1fr) minmax(340px, 0.9fr);
      gap: 18px;
    }}
    .map-analytics {{
      display: grid;
      gap: 18px;
      margin-bottom: 18px;
    }}
    .map-metric-grid,
    .map-chart-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    .map-kpi-card {{
      padding: 18px 18px 16px;
      min-height: 146px;
    }}
    .map-kpi-card h3 {{
      margin: 0 0 10px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--muted);
    }}
    .map-kpi-card strong {{
      display: block;
      font-size: 28px;
      letter-spacing: -0.04em;
      margin-bottom: 8px;
    }}
    .map-kpi-card p {{
      margin: 0;
      font-size: 13px;
      line-height: 1.55;
      color: var(--muted);
    }}
    #projectMap {{
      width: 100%;
      height: 460px;
      border-radius: 18px;
      overflow: hidden;
      border: 1px solid rgba(215, 203, 185, 0.7);
    }}
    .project-filter-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 12px;
    }}
    .project-table-wrap {{
      max-height: 410px;
      overflow: auto;
      border-radius: 16px;
      border: 1px solid rgba(215, 203, 185, 0.7);
      background: var(--panel-2);
    }}
    .project-table {{
      width: 100%;
      border-collapse: collapse;
      min-width: 760px;
    }}
    .project-table th,
    .project-table td {{
      padding: 10px 12px;
      border-bottom: 1px solid rgba(215, 203, 185, 0.55);
      text-align: left;
      font-size: 13px;
      vertical-align: top;
    }}
    .project-table th {{
      position: sticky;
      top: 0;
      z-index: 1;
      background: #f5f9ff;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-size: 11px;
    }}
    .project-table th.sortable {{
      cursor: pointer;
      user-select: none;
    }}
    .sort-indicator {{
      margin-left: 6px;
      color: var(--blue);
      font-size: 10px;
    }}
    .project-table td strong {{
      font-size: 14px;
      letter-spacing: -0.01em;
    }}
    .project-item {{
      padding: 14px;
      border-radius: 16px;
      background: var(--panel-2);
      border: 1px solid rgba(215, 203, 185, 0.7);
    }}
    .project-item h4 {{
      margin: 0 0 8px;
      font-size: 15px;
      letter-spacing: -0.01em;
    }}
    .project-meta {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 6px 12px;
      color: var(--muted);
      font-size: 13px;
    }}
    .table-wrap {{
      max-height: 420px;
      overflow: auto;
      border-radius: 16px;
      border: 1px solid rgba(215, 203, 185, 0.7);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }}
    thead th {{
      text-align: left;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--muted);
      padding: 12px 10px;
      border-bottom: 1px solid var(--line);
      position: sticky;
      top: 0;
      background: rgba(255, 251, 244, 0.98);
    }}
    tbody td {{
      padding: 11px 10px;
      border-bottom: 1px solid rgba(215, 203, 185, 0.6);
    }}
    tbody tr:hover {{
      background: rgba(161, 198, 250, 0.14);
    }}
    .footnote {{
      margin-top: 18px;
      font-size: 13px;
      color: var(--muted);
      line-height: 1.6;
    }}
    @media (max-width: 1180px) {{
      .hero,
      .compare-grid,
      .price-grid,
      .deep-grid,
      .map-chart-grid,
      .map-metric-grid,
      .map-shell,
      .summary-bar {{
        grid-template-columns: 1fr;
      }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <div class="panel hero-card">
        <h1>Hanoi vs HCMC Primary Market</h1>
        <p class="subhead">
          Compare quarterly primary high-rise momentum across Hanoi and HCMC first, then drill into province-level breakdown and segment-level behavior. The dashboard keeps the same business rules you approved for launches, sales, and weighted prices.
        </p>
        <div class="hero-meta">
          <span class="chip">Sale status filter: current</span>
          <span class="chip">Weighted price uses only price &gt; 0</span>
          <span class="chip">Province drill-down: Central / satellite markets</span>
        </div>
      </div>
      <div class="panel control-card">
        <div class="control-group">
          <span class="control-label">Flow Metric</span>
          <div class="button-row" id="flowMetricButtons"></div>
        </div>
        <div class="control-group">
          <span class="control-label">Segment Deep Dive Market</span>
          <div class="button-row" id="deepCityButtons"></div>
        </div>
        <div class="control-group">
          <span class="control-label">Segment Deep Dive Scope</span>
          <div class="button-row" id="deepSegmentButtons"></div>
        </div>
        <div class="control-group">
          <span class="control-label">Timeframe</span>
          <div class="select-row">
            <select class="control-select" id="startQuarterSelect"></select>
            <select class="control-select" id="endQuarterSelect"></select>
          </div>
        </div>
        <div class="control-group">
          <span class="control-label">Current Focus</span>
          <div class="chip" id="focusLabel"></div>
        </div>
        <div class="control-group">
          <button class="refresh-button" id="refreshButton" type="button">Refresh Data</button>
        </div>
      </div>
    </section>

    <section class="summary-bar" id="summaryBar"></section>

    <section class="section-block">
      <div class="section-head">
        <div>
          <h2>Top-Line Comparison</h2>
          <div class="section-note">One chart, two lines, direct comparison between total Hanoi and total HCMC by quarter.</div>
        </div>
      </div>
      <div class="compare-grid">
        <article class="panel chart-card" data-export-key="flow-compare">
          <div class="chart-head">
            <h3 class="chart-title" id="flowCompareTitle">Flow Comparison</h3>
            <span class="chart-note">Total market view</span>
          </div>
          <div class="canvas-wrap"><canvas id="chartFlowCompare"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="latest-snapshot">
          <div class="chart-head">
            <h3 class="chart-title">Latest Quarter Snapshot</h3>
            <span class="chart-note">Fast read on current quarter totals</span>
          </div>
          <div class="legend-list" id="latestCompareList"></div>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-head">
        <div>
          <h2>Province / Submarket Breakdown</h2>
          <div class="section-note">Deeper view of launched and sold volume inside each market. Hanoi splits into Central Hanoi and Van Giang HY. HCMC splits into Central HCMC, Binh Duong, and BR-VT.</div>
        </div>
      </div>
      <div class="price-grid">
        <article class="panel chart-card" data-export-key="hanoi-breakdown">
          <div class="chart-head">
            <h3 class="chart-title">Hanoi Breakdown</h3>
            <span class="chart-note" id="hanoiBreakdownNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartHanoiBreakdown"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="hcmc-breakdown">
          <div class="chart-head">
            <h3 class="chart-title">HCMC Breakdown</h3>
            <span class="chart-note" id="hcmcBreakdownNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartHcmcBreakdown"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="segment-mix-table">
          <div class="chart-head">
            <h3 class="chart-title">Segment Mix Table</h3>
            <span class="chart-note">Latest quarter by selected flow metric</span>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Market</th>
                  <th>Segment</th>
                  <th>Latest quarter</th>
                  <th id="segmentMixMetricHead">Value</th>
                </tr>
              </thead>
              <tbody id="segmentMixBody"></tbody>
            </table>
          </div>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-head">
        <div>
          <h2>Total Market Price Comparison</h2>
          <div class="section-note">Direct comparisons between Hanoi total and HCMC total across price definitions plus quarterly and cumulative sold rate.</div>
        </div>
      </div>
      <div class="price-grid">
        <article class="panel chart-card" data-export-key="price-current-compare">
          <div class="chart-head">
            <h3 class="chart-title">Weighted Price by Current Supply</h3>
            <span class="chart-note">Total Hanoi vs total HCMC</span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartPriceCurrentCompare"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="price-available-compare">
          <div class="chart-head">
            <h3 class="chart-title">Weighted Price by Available Supply</h3>
            <span class="chart-note">Total Hanoi vs total HCMC</span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartPriceAvailableCompare"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="new-project-compare">
          <div class="chart-head">
            <h3 class="chart-title">Average Price of New Projects</h3>
            <span class="chart-note">First-launch project cohorts</span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartNewProjectCompare"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="sold-rate-quarterly-compare">
          <div class="chart-head">
            <h3 class="chart-title">Quarterly Sold Rate</h3>
            <span class="chart-note">New sold / current for sale at the beginning</span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartSoldRateQuarterlyCompare"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="sold-rate-cumulative-compare">
          <div class="chart-head">
            <h3 class="chart-title">Cumulative Sold Rate</h3>
            <span class="chart-note">Current sold-out / (current supply + previous quarter available end)</span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartSoldRateCumulativeCompare"></canvas></div>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-head">
        <div>
          <h2>Deep Dive by Segment</h2>
          <div class="section-note">Use the controls above to inspect a single segment over time. This is where you can read deeper into price, supply-weighted price, and new project formation.</div>
        </div>
      </div>
      <div class="deep-grid">
        <article class="panel chart-card" data-export-key="deep-current-price">
          <div class="chart-head">
            <h3 class="chart-title">Segment Weighted Price by Current Supply</h3>
            <span class="chart-note" id="deepCurrentNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepCurrentPrice"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-available-price">
          <div class="chart-head">
            <h3 class="chart-title">Segment Weighted Price by Available Supply</h3>
            <span class="chart-note" id="deepAvailableNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepAvailablePrice"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-new-project-price">
          <div class="chart-head">
            <h3 class="chart-title">Segment New Project Avg Price</h3>
            <span class="chart-note" id="deepNewProjectNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepNewProjectPrice"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-new-projects">
          <div class="chart-head">
            <h3 class="chart-title">Segment New Projects</h3>
            <span class="chart-note">First positive launch quarter count</span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepNewProjects"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-sold-rate-quarterly">
          <div class="chart-head">
            <h3 class="chart-title">Segment Sold Rate</h3>
            <span class="chart-note" id="deepSoldRateQuarterlyNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepSoldRateQuarterly"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-sold-rate-cumulative">
          <div class="chart-head">
            <h3 class="chart-title">Segment Cumulative Sold Rate</h3>
            <span class="chart-note" id="deepSoldRateCumulativeNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepSoldRateCumulative"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-grading-mix">
          <div class="chart-head">
            <h3 class="chart-title">Flow Mix by Grading</h3>
            <span class="chart-note" id="deepGradingMixNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepGradingMix"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-region-mix">
          <div class="chart-head">
            <h3 class="chart-title">Flow Mix by Region</h3>
            <span class="chart-note" id="deepRegionMixNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepRegionMix"></canvas></div>
        </article>
        <article class="panel chart-card" data-export-key="deep-developer-mix">
          <div class="chart-head">
            <h3 class="chart-title">Flow Mix by Developer</h3>
            <span class="chart-note" id="deepDeveloperMixNote"></span>
          </div>
          <div class="canvas-wrap small"><canvas id="chartDeepDeveloperMix"></canvas></div>
        </article>
        <article class="panel table-card" data-export-key="deep-quarterly-table">
          <div class="chart-head">
            <h3 class="chart-title">Deep Dive Quarterly Table</h3>
            <span class="chart-note">Selected market and segment</span>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Quarter</th>
                  <th>New launched</th>
                  <th>New sold</th>
                  <th>Price current</th>
                  <th>Price available</th>
                  <th>New project avg price</th>
                  <th>Quarterly sold rate</th>
                  <th>Cumulative sold rate</th>
                  <th>Active projects</th>
                  <th>New projects</th>
                </tr>
              </thead>
              <tbody id="deepTableBody"></tbody>
            </table>
          </div>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-head">
        <div>
          <h2>Map of Active Projects</h2>
          <div class="section-note">Projects with sale status = current, joined from Time Series to Project Identification for latitude and longitude. This map is intentionally global and does not follow the dashboard filters.</div>
        </div>
      </div>
      <div class="map-analytics">
        <div class="map-metric-grid" id="mapMetricGrid">
          <article class="panel map-kpi-card">
            <h3>Quarterly Sold Rate</h3>
            <strong id="mapQuarterlySoldRate">-</strong>
            <p id="mapQuarterlySoldRateNote">Latest quarter aggregate for mapped active projects.</p>
          </article>
          <article class="panel map-kpi-card">
            <h3>Accumulated Sold Rate</h3>
            <strong id="mapAccumulatedSoldRate">-</strong>
            <p id="mapAccumulatedSoldRateNote">Current sold-out divided by current supply plus previous quarter available end.</p>
          </article>
        </div>
        <div class="map-chart-grid">
          <article class="panel chart-card" data-export-key="map-launch-segment">
            <div class="chart-head">
              <h3 class="chart-title">New Launched by Segment</h3>
              <span class="chart-note" id="mapLaunchChartNote"></span>
            </div>
            <div class="canvas-wrap small"><canvas id="chartMapLaunchSegment"></canvas></div>
          </article>
          <article class="panel chart-card" data-export-key="map-sold-segment">
            <div class="chart-head">
              <h3 class="chart-title">New Sold by Segment</h3>
              <span class="chart-note" id="mapSoldChartNote"></span>
            </div>
            <div class="canvas-wrap small"><canvas id="chartMapSoldSegment"></canvas></div>
          </article>
        </div>
      </div>
      <div class="map-shell">
        <article class="panel chart-card" data-export-key="project-map">
          <div class="chart-head">
            <h3 class="chart-title">Project Map</h3>
            <span class="chart-note" id="mapNote"></span>
          </div>
          <div id="projectMap"></div>
        </article>
        <article class="panel chart-card" data-export-key="active-project-list">
          <div class="chart-head">
            <h3 class="chart-title">Active Project List</h3>
            <span class="chart-note">Mapped projects in the latest quarter available in the source</span>
          </div>
          <div class="project-filter-row" id="projectCityFilterButtons"></div>
          <div class="project-table-wrap">
            <table class="project-table">
              <thead>
                <tr id="projectTableHeaderRow"></tr>
              </thead>
              <tbody id="projectListBody"></tbody>
            </table>
          </div>
        </article>
      </div>
    </section>

    <section class="panel table-card" data-export-key="quarterly-compare-table">
      <div class="chart-head">
        <h3 class="chart-title">Quarterly Compare Table</h3>
        <span class="chart-note">Total Hanoi and total HCMC side by side</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Quarter</th>
              <th>Hanoi launched</th>
              <th>HCMC launched</th>
              <th>Hanoi sold</th>
              <th>HCMC sold</th>
              <th>Hanoi quarterly sold rate</th>
              <th>HCMC quarterly sold rate</th>
              <th>Hanoi cumulative sold rate</th>
              <th>HCMC cumulative sold rate</th>
              <th>Hanoi active projects</th>
              <th>HCMC active projects</th>
              <th>Hanoi price current</th>
              <th>HCMC price current</th>
            </tr>
          </thead>
          <tbody id="compareTableBody"></tbody>
        </table>
      </div>
      <div class="footnote">
        Rules in use: only rows with sale status = current are included. Launch and sold rollups still keep aggregate total rows where the source workbook only provides total values. Weighted price excludes rows where price &lt;= 0 or the relevant weight is &lt;= 0. Quarterly sold rate = new sold / current for sale at the beginning. Cumulative sold rate = current sold-out / (current supply + previous quarter available end).
      </div>
    </section>
  </div>

  <script>
    const dashboardData = {data_json};

    const state = {{
      flowMetric: "new_launched",
      deepCity: "Hanoi",
      deepSegment: "Total",
      mapCityFilter: "All",
      startQuarter: dashboardData.quarters.find((quarter) => quarterSort(quarter) >= quarterSort("2022Q1")) || dashboardData.quarters[0],
      endQuarter: dashboardData.quarters[dashboardData.quarters.length - 1],
      projectSortKey: "current_supply",
      projectSortDirection: "desc",
    }};

    const charts = {{}};
    let projectMap = null;
    let projectMarkers = [];
    let cityBoundaryLayers = [];
    const cityColors = {{
      Hanoi: "#1B4D5C",
      HCMC: "#3281F5",
    }};
    const segmentColors = {{
      "Central Hanoi": "#1B4D5C",
      "Van Giang HY": "#A1C6FA",
      "Central HCMC": "#3281F5",
      "Binh Duong": "#68B794",
      "BR-VT": "#B59525",
      "Total": "#293537",
    }};
    const gradingColors = {{
      "Affordable": "#9FD356",
      "Mid-end": "#3281F5",
      "Premium": "#68B794",
      "Luxury": "#B59525",
      "Ultra Luxury": "#DC707A",
      "Unknown": "#A0A7B4",
    }};
    const regionColors = {{
      "CBD": "#1B4D5C",
      "Core CBD": "#1B4D5C",
      "East": "#3281F5",
      "North": "#8A6FE8",
      "South": "#68B794",
      "West": "#E08E45",
      "Khu Dong": "#3281F5",
      "Khu Tay": "#E08E45",
      "Khu Nam": "#68B794",
      "Khu Bac": "#8A6FE8",
      "Khu TT": "#4F6D7A",
      "Hung Yen": "#B59525",
      "Unknown": "#A0A7B4",
    }};
    const developerPalette = [
      "#1B4D5C",
      "#3281F5",
      "#68B794",
      "#B59525",
      "#DC707A",
      "#8A6FE8",
      "#E08E45",
      "#4F6D7A",
      "#9FD356",
      "#5C80BC",
      "#D17C9F",
      "#7F8C8D",
    ];
    const metricLabels = {{
      new_launched: "New launched",
      new_sold: "New sold",
      weighted_price_current_supply: "Weighted price by current supply",
      weighted_price_available_supply: "Weighted price by available supply",
      new_project_avg_price: "Average price of new projects",
      sold_rate_quarterly: "Quarterly sold rate",
      sold_rate_cumulative: "Cumulative sold rate",
      active_projects: "Active projects",
      new_projects: "New projects",
      masterise_market_share: "Masterise Homes market share",
    }};
    const projectQuarterLookup = new Map();
    dashboardData.map_points.forEach((point) => {{
      projectQuarterLookup.set(`${{point.city}}::${{point.project_id}}::${{point.quarter}}`, point);
    }});

    function formatNumber(value, kind = "number") {{
      if (value === null || value === undefined || Number.isNaN(value)) {{
        return "-";
      }}
      if (kind === "currency") {{
        return new Intl.NumberFormat("en-US", {{ maximumFractionDigits: 0 }}).format(value);
      }}
      if (kind === "percent") {{
        return `${{new Intl.NumberFormat("en-US", {{ minimumFractionDigits: 1, maximumFractionDigits: 1 }}).format(value)}}%`;
      }}
      if (kind === "decimal") {{
        return new Intl.NumberFormat("en-US", {{ maximumFractionDigits: 1 }}).format(value);
      }}
      return new Intl.NumberFormat("en-US", {{ maximumFractionDigits: 0 }}).format(value);
    }}

    function metricKind(metricKey) {{
      if (metricKey === "sold_rate_quarterly" || metricKey === "sold_rate_cumulative" || metricKey === "masterise_market_share") {{
        return "percent";
      }}
      return metricKey.includes("price") ? "currency" : "number";
    }}

    function formatPercent(value) {{
      if (value === null || value === undefined || Number.isNaN(value)) {{
        return "-";
      }}
      return formatNumber(value, "percent");
    }}

    function deltaClass(value) {{
      if (value === null || value === undefined || Number.isNaN(value)) {{
        return "flat";
      }}
      if (value > 0) {{
        return "up";
      }}
      if (value < 0) {{
        return "down";
      }}
      return "flat";
    }}

    function recordsFor(city, segment) {{
      return dashboardData.records
        .filter((row) => row.city === city && row.segment === segment)
        .sort((a, b) => a.quarter_sort - b.quarter_sort);
    }}

    function quarterSort(quarter) {{
      const row = dashboardData.records.find((item) => item.quarter === quarter);
      return row ? row.quarter_sort : -1;
    }}

    function isQuarterInRange(quarter) {{
      const q = quarterSort(quarter);
      return q >= quarterSort(state.startQuarter) && q <= quarterSort(state.endQuarter);
    }}

    function previousQuarter(quarter) {{
      const match = /^(\\d{{4}})Q([1-4])$/.exec(quarter || "");
      if (!match) {{
        return null;
      }}
      const year = Number(match[1]);
      const quarterNo = Number(match[2]);
      if (quarterNo === 1) {{
        return `${{year - 1}}Q4`;
      }}
      return `${{year}}Q${{quarterNo - 1}}`;
    }}

    function latestMapQuarter() {{
      return dashboardData.quarters[dashboardData.quarters.length - 1];
    }}

    function latestMapPoints() {{
      return dashboardData.map_points
        .filter((row) => row.quarter === latestMapQuarter())
        .sort((a, b) => (b.current_supply ?? 0) - (a.current_supply ?? 0));
    }}

    function previousPoint(point) {{
      const prevQuarter = previousQuarter(point.quarter);
      if (!prevQuarter) {{
        return null;
      }}
      return projectQuarterLookup.get(`${{point.city}}::${{point.project_id}}::${{prevQuarter}}`) || null;
    }}

    function quarterlySoldRateForPoint(point) {{
      const denominator = Number(point.available_begin ?? 0);
      if (!Number.isFinite(denominator) || denominator <= 0) {{
        return null;
      }}
      return (Number(point.new_sold ?? 0) / denominator) * 100;
    }}

    function accumulatedSoldRateForPoint(point) {{
      const prev = previousPoint(point);
      const denominator = Number(point.current_supply ?? 0) + Number(prev?.available_end ?? 0);
      if (!Number.isFinite(denominator) || denominator <= 0) {{
        return null;
      }}
      return (Number(point.current_sold_out ?? 0) / denominator) * 100;
    }}

    function aggregateMapMetrics(points) {{
      let totalNewSold = 0;
      let totalAvailableBegin = 0;
      let totalCurrentSoldOut = 0;
      let totalAccumulatedDenominator = 0;

      points.forEach((point) => {{
        totalNewSold += Number(point.new_sold ?? 0);
        totalAvailableBegin += Number(point.available_begin ?? 0);
        totalCurrentSoldOut += Number(point.current_sold_out ?? 0);
        totalAccumulatedDenominator += Number(point.current_supply ?? 0) + Number(previousPoint(point)?.available_end ?? 0);
      }});

      return {{
        quarterlySoldRate: totalAvailableBegin > 0 ? (totalNewSold / totalAvailableBegin) * 100 : null,
        accumulatedSoldRate: totalAccumulatedDenominator > 0 ? (totalCurrentSoldOut / totalAccumulatedDenominator) * 100 : null,
      }};
    }}

    function groupedSegmentChartData(points, metricKey) {{
      const segmentOrder = ["Central Hanoi", "Van Giang HY", "Central HCMC", "Binh Duong", "BR-VT"];
      const segmentSet = new Set(points.map((point) => point.segment).filter(Boolean));
      const labels = segmentOrder.filter((segment) => segmentSet.has(segment));
      const cities = ["Hanoi", "HCMC"];
      const datasets = cities.map((city) => {{
        return {{
          label: city,
          data: labels.map((segment) => points
            .filter((point) => point.city === city && point.segment === segment)
            .reduce((sum, point) => sum + Number(point[metricKey] ?? 0), 0)),
          backgroundColor: cityColors[city],
          borderColor: cityColors[city],
          borderWidth: 1,
          borderRadius: 8,
          maxBarThickness: 38,
        }};
      }});
      return {{ labels, datasets }};
    }}

    function filteredRecordsFor(city, segment) {{
      return recordsFor(city, segment).filter((row) => isQuarterInRange(row.quarter));
    }}

    function dimensionRecordsFor(city, segment, dimensionType) {{
      return dashboardData.dimension_breakdowns
        .filter((row) => row.city === city && row.segment === segment && row.dimension_type === dimensionType)
        .sort((a, b) => a.quarter_sort - b.quarter_sort);
    }}

    function filteredDimensionRecordsFor(city, segment, dimensionType) {{
      return dimensionRecordsFor(city, segment, dimensionType)
        .filter((row) => isQuarterInRange(row.quarter));
    }}

    function latestRecord(city, segment = "Total") {{
      const records = filteredRecordsFor(city, segment);
      return records.at(-1) || null;
    }}

    function fullLatestRecord(city, segment = "Total") {{
      const records = recordsFor(city, segment);
      return records.at(-1) || null;
    }}

    function unionQuarterLabels(scopes) {{
      const map = new Map();
      scopes.forEach(([city, segment]) => {{
        filteredRecordsFor(city, segment).forEach((row) => {{
          map.set(row.quarter, row.quarter_sort);
        }});
      }});
      return [...map.entries()].sort((a, b) => a[1] - b[1]).map(([quarter]) => quarter);
    }}

    function dimensionQuarterLabels(city, segment, dimensionType) {{
      const map = new Map();
      filteredDimensionRecordsFor(city, segment, dimensionType).forEach((row) => {{
        map.set(row.quarter, row.quarter_sort);
      }});
      return [...map.entries()].sort((a, b) => a[1] - b[1]).map(([quarter]) => quarter);
    }}

    function toSeriesByQuarter(city, segment, labels, metricKey) {{
      const byQuarter = new Map(filteredRecordsFor(city, segment).map((row) => [row.quarter, row]));
      return labels.map((quarter) => byQuarter.get(quarter)?.[metricKey] ?? null);
    }}

    function calcDelta(city, segment, metricKey, stepsBack) {{
      const fullSeries = recordsFor(city, segment);
      const latest = latestRecord(city, segment);
      if (!latest || !fullSeries.length) {{
        return null;
      }}
      const latestIndex = fullSeries.findIndex((row) => row.quarter === latest.quarter);
      if (latestIndex < 0 || latestIndex - stepsBack < 0) {{
        return null;
      }}
      const previousValue = fullSeries[latestIndex - stepsBack]?.[metricKey];
      const currentValue = latest[metricKey];
      if (
        currentValue === null || currentValue === undefined || Number.isNaN(currentValue) ||
        previousValue === null || previousValue === undefined || Number.isNaN(previousValue) ||
        previousValue === 0
      ) {{
        return null;
      }}
      return ((currentValue - previousValue) / previousValue) * 100;
    }}

    function sanitizeFilename(value) {{
      return String(value ?? "export")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "") || "export";
    }}

    function csvCell(value) {{
      if (value === null || value === undefined) {{
        return "";
      }}
      const text = String(value);
      if (/[",\\n]/.test(text)) {{
        return `"${{text.replace(/"/g, '""')}}"`;
      }}
      return text;
    }}

    function downloadRowsAsCsv(filename, rows) {{
      const safeRows = rows.length ? rows : [{{ note: "No data available" }}];
      const headers = [...new Set(safeRows.flatMap((row) => Object.keys(row)))];
      const lines = [
        headers.map(csvCell).join(","),
        ...safeRows.map((row) => headers.map((header) => csvCell(row[header])).join(",")),
      ];
      const blob = new Blob([lines.join("\\n")], {{ type: "text/csv;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }}

    function timeframeSlug() {{
      return `${{state.startQuarter}}_to_${{state.endQuarter}}`;
    }}

    function latestSnapshotRows() {{
      return ["Hanoi", "HCMC"].map((city) => {{
        const latest = latestRecord(city, "Total");
        return {{
          market: city,
          quarter: latest?.quarter ?? "",
          flow_metric: metricLabels[state.flowMetric],
          value: latest?.[state.flowMetric] ?? null,
        }};
      }});
    }}

    function compareMetricRows(metricKey) {{
      const labels = unionQuarterLabels([["Hanoi", "Total"], ["HCMC", "Total"]]);
      return labels.flatMap((quarter) => {{
        return ["Hanoi", "HCMC"].map((city) => {{
          const row = filteredRecordsFor(city, "Total").find((item) => item.quarter === quarter);
          return {{
            quarter,
            market: city,
            metric: metricLabels[metricKey],
            value: row?.[metricKey] ?? null,
          }};
        }});
      }});
    }}

    function breakdownRows(city, segments) {{
      const labels = unionQuarterLabels(segments.map((segment) => [city, segment]));
      return labels.flatMap((quarter) => segments.map((segment) => {{
        const row = filteredRecordsFor(city, segment).find((item) => item.quarter === quarter);
        return {{
          quarter,
          market: city,
          segment,
          flow_metric: metricLabels[state.flowMetric],
          value: row?.[state.flowMetric] ?? null,
        }};
      }}));
    }}

    function segmentMixRows() {{
      return [
        ["Hanoi", ["Central Hanoi", "Van Giang HY"]],
        ["HCMC", ["Central HCMC", "Binh Duong", "BR-VT"]],
      ].flatMap(([city, segments]) => segments.map((segment) => {{
        const latest = latestRecord(city, segment);
        return {{
          market: city,
          segment,
          latest_quarter: latest?.quarter ?? "",
          flow_metric: metricLabels[state.flowMetric],
          value: latest?.[state.flowMetric] ?? null,
        }};
      }}));
    }}

    function deepMetricRows(metricKey) {{
      return filteredRecordsFor(state.deepCity, state.deepSegment).map((row) => {{
        return {{
          quarter: row.quarter,
          market: state.deepCity,
          segment: state.deepSegment,
          metric: metricLabels[metricKey],
          value: row?.[metricKey] ?? null,
        }};
      }});
    }}

    function deepDimensionRows(dimensionType) {{
      const metricKey = state.flowMetric;
      const rows = filteredDimensionRecordsFor(state.deepCity, state.deepSegment, dimensionType);
      if (dimensionType === "developer") {{
        const labels = dimensionQuarterLabels(state.deepCity, state.deepSegment, dimensionType);
        const latestQuarter = labels[labels.length - 1];
        const latestRows = rows.filter((row) => row.quarter === latestQuarter);
        const totals = new Map();
        latestRows.forEach((row) => {{
          totals.set(row.dimension_value, (totals.get(row.dimension_value) || 0) + Number(row[metricKey] ?? 0));
        }});
        const ordered = [...totals.entries()].sort((a, b) => b[1] - a[1]);
        const topFive = ordered.slice(0, 5).map(([dimension_value, value]) => ({{
          quarter: latestQuarter,
          market: state.deepCity,
          segment: state.deepSegment,
          dimension_type: "developer",
          dimension_value,
          flow_metric: metricLabels[metricKey],
          value,
          grouping: "Top 5",
        }}));
        const otherValue = ordered.slice(5).reduce((sum, [, value]) => sum + value, 0);
        if (otherValue > 0) {{
          topFive.push({{
            quarter: latestQuarter,
            market: state.deepCity,
            segment: state.deepSegment,
            dimension_type: "developer",
            dimension_value: "Other",
            flow_metric: metricLabels[metricKey],
            value: otherValue,
            grouping: "Other",
          }});
        }}
        return topFive;
      }}
      return rows.map((row) => {{
        return {{
          quarter: row.quarter,
          market: row.city,
          segment: row.segment,
          dimension_type: row.dimension_type,
          dimension_value: row.dimension_value,
          flow_metric: metricLabels[metricKey],
          value: row[metricKey] ?? null,
        }};
      }});
    }}

    function deepTableRows() {{
      return [...recordsFor(state.deepCity, state.deepSegment)]
        .filter((row) => isQuarterInRange(row.quarter))
        .map((row) => {{
          return {{
            quarter: row.quarter,
            market: state.deepCity,
            segment: state.deepSegment,
            new_launched: row.new_launched,
            new_sold: row.new_sold,
            price_current: row.weighted_price_current_supply,
            price_available: row.weighted_price_available_supply,
            new_project_avg_price: row.new_project_avg_price,
            quarterly_sold_rate: row.sold_rate_quarterly,
            cumulative_sold_rate: row.sold_rate_cumulative,
            active_projects: row.active_projects,
            new_projects: row.new_projects,
          }};
        }});
    }}

    function quarterlyCompareRows() {{
      const labels = unionQuarterLabels([["Hanoi", "Total"], ["HCMC", "Total"]]).reverse();
      const hanoi = new Map(filteredRecordsFor("Hanoi", "Total").map((row) => [row.quarter, row]));
      const hcmc = new Map(filteredRecordsFor("HCMC", "Total").map((row) => [row.quarter, row]));
      return labels.map((quarter) => {{
        const hn = hanoi.get(quarter);
        const hc = hcmc.get(quarter);
        return {{
          quarter,
          hanoi_launched: hn?.new_launched ?? null,
          hcmc_launched: hc?.new_launched ?? null,
          hanoi_sold: hn?.new_sold ?? null,
          hcmc_sold: hc?.new_sold ?? null,
          hanoi_quarterly_sold_rate: hn?.sold_rate_quarterly ?? null,
          hcmc_quarterly_sold_rate: hc?.sold_rate_quarterly ?? null,
          hanoi_cumulative_sold_rate: hn?.sold_rate_cumulative ?? null,
          hcmc_cumulative_sold_rate: hc?.sold_rate_cumulative ?? null,
          hanoi_active_projects: hn?.active_projects ?? null,
          hcmc_active_projects: hc?.active_projects ?? null,
          hanoi_price_current: hn?.weighted_price_current_supply ?? null,
          hcmc_price_current: hc?.weighted_price_current_supply ?? null,
        }};
      }});
    }}

    function mapSegmentRows(metricKey) {{
      const latestQuarter = latestMapQuarter();
      return latestMapPoints().map((point) => point.segment)
        .filter((segment, index, all) => segment && all.indexOf(segment) === index)
        .flatMap((segment) => ["Hanoi", "HCMC"].map((city) => {{
          const points = latestMapPoints().filter((point) => point.city === city && point.segment === segment);
          return {{
            quarter: latestQuarter,
            market: city,
            segment,
            metric: metricLabels[metricKey],
            value: points.reduce((sum, point) => sum + Number(point[metricKey] ?? 0), 0),
          }};
        }}));
    }}

    function mappedProjectRows(points) {{
      return points.map((point) => {{
        const quarterlySoldRate = quarterlySoldRateForPoint(point);
        const accumulatedSoldRate = accumulatedSoldRateForPoint(point);
        return {{
          project_name: point.project_name,
          developer: point.developer || "",
          market: point.city,
          segment: point.segment,
          quarter: point.quarter,
          latitude: point.latitude,
          longitude: point.longitude,
          new_launched: point.new_launched,
          new_sold: point.new_sold,
          current_supply: point.current_supply,
          available_begin: point.available_begin,
          quarterly_sold_rate: quarterlySoldRate,
          accumulated_sold_rate: accumulatedSoldRate,
          price: point.price,
        }};
      }});
    }}

    function exportPayloadFor(key) {{
      const payloads = {{
        "flow-compare": () => ({{
          filename: `flow-compare-${{sanitizeFilename(metricLabels[state.flowMetric])}}-${{timeframeSlug()}}.csv`,
          rows: compareMetricRows(state.flowMetric),
        }}),
        "latest-snapshot": () => ({{
          filename: `latest-quarter-snapshot-${{sanitizeFilename(metricLabels[state.flowMetric])}}.csv`,
          rows: latestSnapshotRows(),
        }}),
        "hanoi-breakdown": () => ({{
          filename: `hanoi-breakdown-${{sanitizeFilename(metricLabels[state.flowMetric])}}-${{timeframeSlug()}}.csv`,
          rows: breakdownRows("Hanoi", ["Central Hanoi", "Van Giang HY"]),
        }}),
        "hcmc-breakdown": () => ({{
          filename: `hcmc-breakdown-${{sanitizeFilename(metricLabels[state.flowMetric])}}-${{timeframeSlug()}}.csv`,
          rows: breakdownRows("HCMC", ["Central HCMC", "Binh Duong", "BR-VT"]),
        }}),
        "segment-mix-table": () => ({{
          filename: `segment-mix-${{sanitizeFilename(metricLabels[state.flowMetric])}}.csv`,
          rows: segmentMixRows(),
        }}),
        "price-current-compare": () => ({{
          filename: `weighted-price-current-supply-${{timeframeSlug()}}.csv`,
          rows: compareMetricRows("weighted_price_current_supply"),
        }}),
        "price-available-compare": () => ({{
          filename: `weighted-price-available-supply-${{timeframeSlug()}}.csv`,
          rows: compareMetricRows("weighted_price_available_supply"),
        }}),
        "new-project-compare": () => ({{
          filename: `average-price-new-projects-${{timeframeSlug()}}.csv`,
          rows: compareMetricRows("new_project_avg_price"),
        }}),
        "sold-rate-quarterly-compare": () => ({{
          filename: `quarterly-sold-rate-${{timeframeSlug()}}.csv`,
          rows: compareMetricRows("sold_rate_quarterly"),
        }}),
        "sold-rate-cumulative-compare": () => ({{
          filename: `cumulative-sold-rate-${{timeframeSlug()}}.csv`,
          rows: compareMetricRows("sold_rate_cumulative"),
        }}),
        "deep-current-price": () => ({{
          filename: `deep-current-price-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepMetricRows("weighted_price_current_supply"),
        }}),
        "deep-available-price": () => ({{
          filename: `deep-available-price-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepMetricRows("weighted_price_available_supply"),
        }}),
        "deep-new-project-price": () => ({{
          filename: `deep-new-project-price-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepMetricRows("new_project_avg_price"),
        }}),
        "deep-new-projects": () => ({{
          filename: `deep-new-projects-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepMetricRows("new_projects"),
        }}),
        "deep-sold-rate-quarterly": () => ({{
          filename: `deep-quarterly-sold-rate-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepMetricRows("sold_rate_quarterly"),
        }}),
        "deep-sold-rate-cumulative": () => ({{
          filename: `deep-cumulative-sold-rate-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepMetricRows("sold_rate_cumulative"),
        }}),
        "deep-grading-mix": () => ({{
          filename: `deep-grading-mix-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{sanitizeFilename(metricLabels[state.flowMetric])}}-${{timeframeSlug()}}.csv`,
          rows: deepDimensionRows("grading"),
        }}),
        "deep-region-mix": () => ({{
          filename: `deep-region-mix-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{sanitizeFilename(metricLabels[state.flowMetric])}}-${{timeframeSlug()}}.csv`,
          rows: deepDimensionRows("region"),
        }}),
        "deep-developer-mix": () => ({{
          filename: `deep-developer-mix-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{sanitizeFilename(metricLabels[state.flowMetric])}}.csv`,
          rows: deepDimensionRows("developer"),
        }}),
        "deep-quarterly-table": () => ({{
          filename: `deep-quarterly-table-${{sanitizeFilename(state.deepCity)}}-${{sanitizeFilename(state.deepSegment)}}-${{timeframeSlug()}}.csv`,
          rows: deepTableRows(),
        }}),
        "map-launch-segment": () => ({{
          filename: `map-new-launched-by-segment-${{latestMapQuarter()}}.csv`,
          rows: mapSegmentRows("new_launched"),
        }}),
        "map-sold-segment": () => ({{
          filename: `map-new-sold-by-segment-${{latestMapQuarter()}}.csv`,
          rows: mapSegmentRows("new_sold"),
        }}),
        "project-map": () => ({{
          filename: `project-map-points-${{sanitizeFilename(state.mapCityFilter)}}-${{latestMapQuarter()}}.csv`,
          rows: mappedProjectRows(filteredProjectPoints(latestMapPoints())),
        }}),
        "active-project-list": () => ({{
          filename: `active-project-list-${{sanitizeFilename(state.mapCityFilter)}}-${{latestMapQuarter()}}.csv`,
          rows: mappedProjectRows(filteredProjectPoints(latestMapPoints())),
        }}),
        "quarterly-compare-table": () => ({{
          filename: `quarterly-compare-table-${{timeframeSlug()}}.csv`,
          rows: quarterlyCompareRows(),
        }}),
      }};
      return payloads[key] ? payloads[key]() : null;
    }}

    function initializeExportButtons() {{
      document.querySelectorAll("[data-export-key]").forEach((panel) => {{
        if (panel.querySelector(".export-button")) {{
          return;
        }}
        const key = panel.dataset.exportKey;
        const button = document.createElement("button");
        button.type = "button";
        button.className = "export-button";
        button.textContent = "Download CSV";
        button.addEventListener("click", () => {{
          const payload = exportPayloadFor(key);
          if (!payload) {{
            return;
          }}
          downloadRowsAsCsv(payload.filename, payload.rows);
        }});
        panel.appendChild(button);
      }});
    }}

    function renderToggleRow(hostId, options, activeValue, onClick) {{
      const host = document.getElementById(hostId);
      host.innerHTML = "";
      options.forEach(([value, label]) => {{
        const button = document.createElement("button");
        button.className = `toggle${{value === activeValue ? " active" : ""}}`;
        button.textContent = label;
        button.addEventListener("click", () => onClick(value));
        host.appendChild(button);
      }});
    }}

    function renderProjectCityFilterButtons() {{
      renderToggleRow(
        "projectCityFilterButtons",
        [
          ["All", "All"],
          ["Hanoi", "Hanoi"],
          ["HCMC", "HCMC"],
        ],
        state.mapCityFilter,
        (value) => {{
          state.mapCityFilter = value;
          renderProjectCityFilterButtons();
          renderMap();
        }}
      );
    }}

    function renderProjectTableHeaders() {{
      const headers = [
        ["project_name", "Project"],
        ["developer", "Developer"],
        ["city", "Market"],
        ["segment", "Segment"],
        ["quarter", "Quarter"],
        ["new_launched", "New launched"],
        ["new_sold", "New sold"],
        ["current_supply", "Current supply"],
        ["available_begin", "Available begin"],
        ["quarterly_sold_rate", "Quarterly sold rate"],
        ["accumulated_sold_rate", "Accumulated sold rate"],
        ["price", "Price"],
      ];
      const row = document.getElementById("projectTableHeaderRow");
      row.innerHTML = "";
      headers.forEach(([key, label]) => {{
        const th = document.createElement("th");
        th.className = "sortable";
        const isActive = state.projectSortKey === key;
        const arrow = isActive ? (state.projectSortDirection === "asc" ? "▲" : "▼") : "";
        th.innerHTML = `${{label}}<span class="sort-indicator">${{arrow}}</span>`;
        th.addEventListener("click", () => {{
          if (state.projectSortKey === key) {{
            state.projectSortDirection = state.projectSortDirection === "asc" ? "desc" : "asc";
          }} else {{
            state.projectSortKey = key;
            state.projectSortDirection = ["project_name", "developer", "city", "segment", "quarter"].includes(key) ? "asc" : "desc";
          }}
          renderProjectTableHeaders();
          renderMap();
        }});
        row.appendChild(th);
      }});
    }}

    function filteredProjectPoints(points) {{
      if (state.mapCityFilter === "All") {{
        return points;
      }}
      return points.filter((point) => point.city === state.mapCityFilter);
    }}

    function renderSummary() {{
      const host = document.getElementById("summaryBar");
      host.innerHTML = "";
      const metrics = [
        "new_launched",
        "new_sold",
        "sold_rate_quarterly",
        "weighted_price_current_supply",
        "new_projects",
        "masterise_market_share",
      ];
      const latestHanoi = latestRecord("Hanoi", "Total");
      const latestHcmc = latestRecord("HCMC", "Total");
      const latestQuarter = latestHcmc?.quarter || latestHanoi?.quarter || "No data";
      document.getElementById("focusLabel").textContent = `${{state.deepCity}} / ${{state.deepSegment}} / Flow: ${{metricLabels[state.flowMetric]}}`;

      metrics.forEach((metricKey) => {{
        const card = document.createElement("article");
        card.className = "panel summary-card";
        const hanoiQoQ = calcDelta("Hanoi", "Total", metricKey, 1);
        const hanoiYoY = calcDelta("Hanoi", "Total", metricKey, 4);
        const hcmcQoQ = calcDelta("HCMC", "Total", metricKey, 1);
        const hcmcYoY = calcDelta("HCMC", "Total", metricKey, 4);
        card.innerHTML = `
          <h3>${{metricLabels[metricKey]}}</h3>
          <div class="summary-values">
            <div class="summary-line"><span>Hanoi</span><strong>${{formatNumber(latestHanoi?.[metricKey], metricKind(metricKey))}}</strong></div>
            <div class="delta-line"><span>QoQ</span><span class="delta-value ${{deltaClass(hanoiQoQ)}}">${{formatPercent(hanoiQoQ)}}</span></div>
            <div class="delta-line"><span>YoY</span><span class="delta-value ${{deltaClass(hanoiYoY)}}">${{formatPercent(hanoiYoY)}}</span></div>
            <div class="summary-line"><span>HCMC</span><strong>${{formatNumber(latestHcmc?.[metricKey], metricKind(metricKey))}}</strong></div>
            <div class="delta-line"><span>QoQ</span><span class="delta-value ${{deltaClass(hcmcQoQ)}}">${{formatPercent(hcmcQoQ)}}</span></div>
            <div class="delta-line"><span>YoY</span><span class="delta-value ${{deltaClass(hcmcYoY)}}">${{formatPercent(hcmcYoY)}}</span></div>
          </div>
          <div class="summary-foot">Latest reference: ${{latestQuarter}}</div>
        `;
        host.appendChild(card);
      }});
    }}

    function buildOrReplaceChart(chartId, config) {{
      if (charts[chartId]) {{
        charts[chartId].destroy();
      }}
      charts[chartId] = new Chart(document.getElementById(chartId), config);
    }}

    function baseLineOptions(metricKey, yLabel, stacked = false) {{
      return {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{
          intersect: false,
          mode: "index",
        }},
        scales: {{
          x: {{
            stacked,
            grid: {{ display: false }},
            ticks: {{ color: "#6d6458" }},
          }},
          y: {{
            stacked,
            beginAtZero: metricKey.includes("price") ? false : true,
            grid: {{ color: "rgba(161, 198, 250, 0.20)" }},
            ticks: {{
              color: "#5F6B6D",
              callback(value) {{
                return formatNumber(value, metricKind(metricKey));
              }},
            }},
            title: {{
              display: true,
              text: yLabel,
              color: "#5F6B6D",
            }},
          }},
        }},
        plugins: {{
          legend: {{
            display: true,
            labels: {{
              usePointStyle: true,
              boxWidth: 8,
              color: "#293537",
            }},
          }},
          tooltip: {{
            backgroundColor: "rgba(41, 53, 55, 0.94)",
            padding: 12,
            callbacks: {{
              label(context) {{
                return `${{context.dataset.label}}: ${{formatNumber(context.parsed.y, metricKind(metricKey))}}`;
              }},
            }},
          }},
        }},
      }};
    }}

    function dimensionColor(dimensionType, label) {{
      if (dimensionType === "grading") {{
        return gradingColors[label] || "#A0A7B4";
      }}
      if (dimensionType === "developer") {{
        const index = [...label].reduce((sum, char) => sum + char.charCodeAt(0), 0) % developerPalette.length;
        return developerPalette[index];
      }}
      return regionColors[label] || "#A0A7B4";
    }}

    function dimensionOrder(dimensionType, labels, records = []) {{
      const gradingOrder = ["Affordable", "Mid-end", "Premium", "Luxury", "Ultra Luxury", "Unknown"];
      const regionOrder = ["CBD", "Core CBD", "East", "North", "South", "West", "Khu Dong", "Khu Tay", "Khu Nam", "Khu Bac", "Khu TT", "Hung Yen", "Unknown"];
      if (dimensionType === "developer") {{
        const totals = new Map();
        records.forEach((row) => {{
          const key = row.dimension_value;
          totals.set(key, (totals.get(key) || 0) + Number(row[state.flowMetric] ?? 0));
        }});
        return [...labels].sort((a, b) => (totals.get(b) || 0) - (totals.get(a) || 0));
      }}
      const preferred = dimensionType === "grading" ? gradingOrder : regionOrder;
      const labelSet = new Set(labels);
      const ordered = preferred.filter((label) => labelSet.has(label));
      const extras = labels.filter((label) => !preferred.includes(label)).sort();
      return [...ordered, ...extras];
    }}

    function renderFlowCompare() {{
      const labels = unionQuarterLabels([["Hanoi", "Total"], ["HCMC", "Total"]]);
      const metricKey = state.flowMetric;
      document.getElementById("flowCompareTitle").textContent = `${{metricLabels[metricKey]}}: Hanoi vs HCMC`;
      buildOrReplaceChart("chartFlowCompare", {{
        type: "line",
        data: {{
          labels,
          datasets: [
            {{
              label: "Hanoi",
              data: toSeriesByQuarter("Hanoi", "Total", labels, metricKey),
              borderColor: cityColors.Hanoi,
              backgroundColor: cityColors.Hanoi,
              pointRadius: 3,
              pointHoverRadius: 5,
              borderWidth: 3,
              tension: 0.25,
            }},
            {{
              label: "HCMC",
              data: toSeriesByQuarter("HCMC", "Total", labels, metricKey),
              borderColor: cityColors.HCMC,
              backgroundColor: cityColors.HCMC,
              pointRadius: 3,
              pointHoverRadius: 5,
              borderWidth: 3,
              tension: 0.25,
            }},
          ],
        }},
        options: baseLineOptions(
          metricKey,
          metricKind(metricKey) === "currency" ? "VND" : metricKind(metricKey) === "percent" ? "%" : "Units"
        ),
      }});

      const latestList = document.getElementById("latestCompareList");
      latestList.innerHTML = "";
      ["Hanoi", "HCMC"].forEach((city) => {{
        const latest = latestRecord(city, "Total");
        const item = document.createElement("div");
        item.className = "legend-item";
        item.innerHTML = `
          <div><span class="legend-swatch" style="background:${{cityColors[city]}}"></span>${{city}}</div>
          <strong>${{formatNumber(latest?.[metricKey], metricKind(metricKey))}}</strong>
        `;
        latestList.appendChild(item);
      }});
    }}

    function renderBreakdownChart(chartId, city, segments) {{
      const labels = unionQuarterLabels(segments.map((segment) => [city, segment]));
      const metricKey = state.flowMetric;
      const datasets = segments.map((segment) => {{
        return {{
          label: segment,
          data: toSeriesByQuarter(city, segment, labels, metricKey),
          backgroundColor: segmentColors[segment] || "#999999",
          borderColor: segmentColors[segment] || "#999999",
          borderWidth: 1,
        }};
      }});
      buildOrReplaceChart(chartId, {{
        type: "bar",
        data: {{ labels, datasets }},
        options: baseLineOptions(
          metricKey,
          metricKind(metricKey) === "percent" ? "%" : "Units",
          true
        ),
      }});
    }}

    function renderBreakdownSection() {{
      document.getElementById("hanoiBreakdownNote").textContent = metricLabels[state.flowMetric];
      document.getElementById("hcmcBreakdownNote").textContent = metricLabels[state.flowMetric];
      renderBreakdownChart("chartHanoiBreakdown", "Hanoi", ["Central Hanoi", "Van Giang HY"]);
      renderBreakdownChart("chartHcmcBreakdown", "HCMC", ["Central HCMC", "Binh Duong", "BR-VT"]);

      const body = document.getElementById("segmentMixBody");
      body.innerHTML = "";
      document.getElementById("segmentMixMetricHead").textContent = metricLabels[state.flowMetric];
      [["Hanoi", ["Central Hanoi", "Van Giang HY"]], ["HCMC", ["Central HCMC", "Binh Duong", "BR-VT"]]].forEach(([city, segments]) => {{
        segments.forEach((segment) => {{
          const latest = latestRecord(city, segment);
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td>${{city}}</td>
            <td>${{segment}}</td>
            <td>${{latest?.quarter ?? "-"}}</td>
            <td>${{formatNumber(latest?.[state.flowMetric], metricKind(state.flowMetric))}}</td>
          `;
          body.appendChild(tr);
        }});
      }});
    }}

    function renderPriceCompareChart(chartId, metricKey) {{
      const yLabel = metricKind(metricKey) === "currency" ? "VND" : metricKind(metricKey) === "percent" ? "%" : "Units";
      const labels = unionQuarterLabels([["Hanoi", "Total"], ["HCMC", "Total"]]);
      buildOrReplaceChart(chartId, {{
        type: "line",
        data: {{
          labels,
          datasets: [
            {{
              label: "Hanoi",
              data: toSeriesByQuarter("Hanoi", "Total", labels, metricKey),
              borderColor: cityColors.Hanoi,
              backgroundColor: cityColors.Hanoi,
              pointRadius: 2.5,
              borderWidth: 3,
              tension: 0.22,
            }},
            {{
              label: "HCMC",
              data: toSeriesByQuarter("HCMC", "Total", labels, metricKey),
              borderColor: cityColors.HCMC,
              backgroundColor: cityColors.HCMC,
              pointRadius: 2.5,
              borderWidth: 3,
              tension: 0.22,
            }},
          ],
        }},
        options: baseLineOptions(metricKey, yLabel),
      }});
    }}

    function renderPriceCompare() {{
      renderPriceCompareChart("chartPriceCurrentCompare", "weighted_price_current_supply");
      renderPriceCompareChart("chartPriceAvailableCompare", "weighted_price_available_supply");
      renderPriceCompareChart("chartNewProjectCompare", "new_project_avg_price");
      renderPriceCompareChart("chartSoldRateQuarterlyCompare", "sold_rate_quarterly");
      renderPriceCompareChart("chartSoldRateCumulativeCompare", "sold_rate_cumulative");
    }}

    function renderDeepCityButtons() {{
      renderToggleRow(
        "deepCityButtons",
        dashboardData.markets.map((city) => [city, city]),
        state.deepCity,
        (value) => {{
          state.deepCity = value;
          state.deepSegment = "Total";
          renderDeepSegmentButtons();
          renderDeepDive();
        }}
      );
    }}

    function renderDeepSegmentButtons() {{
      renderToggleRow(
        "deepSegmentButtons",
        (dashboardData.segments[state.deepCity] || ["Total"]).map((segment) => [segment, segment]),
        state.deepSegment,
        (value) => {{
          state.deepSegment = value;
          renderDeepDive();
        }}
      );
    }}

    function renderDeepLineChart(chartId, metricKey, yLabel) {{
      const labels = unionQuarterLabels([[state.deepCity, state.deepSegment]]);
      buildOrReplaceChart(chartId, {{
        type: "line",
        data: {{
          labels,
          datasets: [
            {{
              label: `${{state.deepCity}} / ${{state.deepSegment}}`,
              data: toSeriesByQuarter(state.deepCity, state.deepSegment, labels, metricKey),
              borderColor: segmentColors[state.deepSegment] || cityColors[state.deepCity],
              backgroundColor: segmentColors[state.deepSegment] || cityColors[state.deepCity],
              pointRadius: 2.5,
              borderWidth: 3,
              tension: 0.22,
            }},
          ],
        }},
        options: baseLineOptions(metricKey, yLabel),
      }});
    }}

    function renderDeepBarChart(chartId, metricKey, yLabel) {{
      const labels = unionQuarterLabels([[state.deepCity, state.deepSegment]]);
      buildOrReplaceChart(chartId, {{
        type: "bar",
        data: {{
          labels,
          datasets: [
            {{
              label: `${{state.deepCity}} / ${{state.deepSegment}}`,
              data: toSeriesByQuarter(state.deepCity, state.deepSegment, labels, metricKey),
              backgroundColor: segmentColors[state.deepSegment] || cityColors[state.deepCity],
              borderColor: segmentColors[state.deepSegment] || cityColors[state.deepCity],
              borderWidth: 1,
            }},
          ],
        }},
        options: baseLineOptions(metricKey, yLabel),
      }});
    }}

    function renderDeepDimensionChart(chartId, dimensionType) {{
      const metricKey = state.flowMetric;
      let records = filteredDimensionRecordsFor(state.deepCity, state.deepSegment, dimensionType);
      if (dimensionType === "developer") {{
        const labels = dimensionQuarterLabels(state.deepCity, state.deepSegment, dimensionType);
        const latestQuarter = labels[labels.length - 1];
        const latestRecords = records.filter((row) => row.quarter === latestQuarter);
        const totals = new Map();
        latestRecords.forEach((row) => {{
          totals.set(row.dimension_value, (totals.get(row.dimension_value) || 0) + Number(row[metricKey] ?? 0));
        }});
        const ordered = [...totals.entries()].sort((a, b) => b[1] - a[1]);
        const topFive = ordered.slice(0, 5);
        const otherTotal = ordered.slice(5).reduce((sum, [, value]) => sum + value, 0);
        const pieLabels = topFive.map(([label]) => label);
        const pieData = topFive.map(([, value]) => value);
        if (otherTotal > 0) {{
          pieLabels.push("Other");
          pieData.push(otherTotal);
        }}
        buildOrReplaceChart(chartId, {{
          type: "pie",
          data: {{
            labels: pieLabels,
            datasets: [{{
              data: pieData,
              backgroundColor: pieLabels.map((label) => dimensionColor("developer", label)),
              borderColor: "#ffffff",
              borderWidth: 2,
            }}],
          }},
          options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
              legend: {{
                position: "top",
                labels: {{
                  usePointStyle: true,
                  boxWidth: 8,
                  color: "#293537",
                }},
              }},
              tooltip: {{
                backgroundColor: "rgba(41, 53, 55, 0.94)",
                titleColor: "#fff",
                bodyColor: "#fff",
                padding: 12,
                callbacks: {{
                  label(context) {{
                    return `${{context.label}}: ${{formatNumber(context.parsed, metricKind(metricKey))}}`;
                  }},
                }},
              }},
            }},
          }},
        }});
        return;
      }}
      const labels = dimensionQuarterLabels(state.deepCity, state.deepSegment, dimensionType);
      const dimensionLabels = dimensionOrder(
        dimensionType,
        [...new Set(records.map((row) => row.dimension_value))],
        records
      );
      const datasets = dimensionLabels.map((label) => {{
        const recordMap = new Map(
          records
            .filter((row) => row.dimension_value === label)
            .map((row) => [row.quarter, row])
        );
        return {{
          label,
          data: labels.map((quarter) => recordMap.get(quarter)?.[metricKey] ?? 0),
          backgroundColor: dimensionColor(dimensionType, label),
          borderColor: dimensionColor(dimensionType, label),
          borderWidth: 1,
          borderRadius: 8,
          maxBarThickness: 34,
        }};
      }});
      buildOrReplaceChart(chartId, {{
        type: "bar",
        data: {{ labels, datasets }},
        options: baseLineOptions(metricKey, "Units", true),
      }});
    }}

    function renderDeepDive() {{
      document.getElementById("deepCurrentNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}}`;
      document.getElementById("deepAvailableNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}}`;
      document.getElementById("deepNewProjectNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}}`;
      document.getElementById("deepSoldRateQuarterlyNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}}`;
      document.getElementById("deepSoldRateCumulativeNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}}`;
      document.getElementById("deepGradingMixNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}} / ${{metricLabels[state.flowMetric]}}`;
      document.getElementById("deepRegionMixNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}} / ${{metricLabels[state.flowMetric]}}`;
      document.getElementById("deepDeveloperMixNote").textContent = `${{state.deepCity}} / ${{state.deepSegment}} / ${{metricLabels[state.flowMetric]}} / latest quarter / top 5`;

      renderDeepLineChart("chartDeepCurrentPrice", "weighted_price_current_supply", "VND");
      renderDeepLineChart("chartDeepAvailablePrice", "weighted_price_available_supply", "VND");
      renderDeepLineChart("chartDeepNewProjectPrice", "new_project_avg_price", "VND");
      renderDeepLineChart("chartDeepSoldRateQuarterly", "sold_rate_quarterly", "%");
      renderDeepLineChart("chartDeepSoldRateCumulative", "sold_rate_cumulative", "%");
      renderDeepBarChart("chartDeepNewProjects", "new_projects", "Projects");
      renderDeepDimensionChart("chartDeepGradingMix", "grading");
      renderDeepDimensionChart("chartDeepRegionMix", "region");
      renderDeepDimensionChart("chartDeepDeveloperMix", "developer");

      const body = document.getElementById("deepTableBody");
      body.innerHTML = "";
      const rows = [...recordsFor(state.deepCity, state.deepSegment)].reverse();
      const filteredRows = rows.filter((row) => isQuarterInRange(row.quarter));
      filteredRows.forEach((row) => {{
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${{row.quarter}}</td>
          <td>${{formatNumber(row.new_launched)}}</td>
          <td>${{formatNumber(row.new_sold, "decimal")}}</td>
          <td>${{formatNumber(row.weighted_price_current_supply, "currency")}}</td>
          <td>${{formatNumber(row.weighted_price_available_supply, "currency")}}</td>
          <td>${{formatNumber(row.new_project_avg_price, "currency")}}</td>
          <td>${{formatNumber(row.sold_rate_quarterly, "percent")}}</td>
          <td>${{formatNumber(row.sold_rate_cumulative, "percent")}}</td>
          <td>${{formatNumber(row.active_projects)}}</td>
          <td>${{formatNumber(row.new_projects)}}</td>
        `;
        body.appendChild(tr);
      }});
    }}

    function renderCompareTable() {{
      const body = document.getElementById("compareTableBody");
      body.innerHTML = "";
      const labels = unionQuarterLabels([["Hanoi", "Total"], ["HCMC", "Total"]]).reverse();
      const hanoi = new Map(filteredRecordsFor("Hanoi", "Total").map((row) => [row.quarter, row]));
      const hcmc = new Map(filteredRecordsFor("HCMC", "Total").map((row) => [row.quarter, row]));
      labels.forEach((quarter) => {{
        const hn = hanoi.get(quarter);
        const hc = hcmc.get(quarter);
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${{quarter}}</td>
          <td>${{formatNumber(hn?.new_launched)}}</td>
          <td>${{formatNumber(hc?.new_launched)}}</td>
          <td>${{formatNumber(hn?.new_sold, "decimal")}}</td>
          <td>${{formatNumber(hc?.new_sold, "decimal")}}</td>
          <td>${{formatNumber(hn?.sold_rate_quarterly, "percent")}}</td>
          <td>${{formatNumber(hc?.sold_rate_quarterly, "percent")}}</td>
          <td>${{formatNumber(hn?.sold_rate_cumulative, "percent")}}</td>
          <td>${{formatNumber(hc?.sold_rate_cumulative, "percent")}}</td>
          <td>${{formatNumber(hn?.active_projects)}}</td>
          <td>${{formatNumber(hc?.active_projects)}}</td>
          <td>${{formatNumber(hn?.weighted_price_current_supply, "currency")}}</td>
          <td>${{formatNumber(hc?.weighted_price_current_supply, "currency")}}</td>
        `;
        body.appendChild(tr);
      }});
    }}

    function renderMapAnalytics(points, latestQuarter) {{
      const metrics = aggregateMapMetrics(points);
      document.getElementById("mapQuarterlySoldRate").textContent = formatNumber(metrics.quarterlySoldRate, "percent");
      document.getElementById("mapAccumulatedSoldRate").textContent = formatNumber(metrics.accumulatedSoldRate, "percent");
      document.getElementById("mapQuarterlySoldRateNote").textContent = `Latest quarter / ${{latestQuarter}} / formula: new sold / available at beginning`;
      document.getElementById("mapAccumulatedSoldRateNote").textContent = `Latest quarter / ${{latestQuarter}} / formula: current sold-out / (current supply + previous quarter available end)`;
      document.getElementById("mapLaunchChartNote").textContent = `Latest quarter / ${{latestQuarter}} / stacked by market`;
      document.getElementById("mapSoldChartNote").textContent = `Latest quarter / ${{latestQuarter}} / stacked by market`;

      buildOrReplaceChart("chartMapLaunchSegment", {{
        type: "bar",
        data: groupedSegmentChartData(points, "new_launched"),
        options: baseLineOptions("new_launched", "Units", true),
      }});
      buildOrReplaceChart("chartMapSoldSegment", {{
        type: "bar",
        data: groupedSegmentChartData(points, "new_sold"),
        options: baseLineOptions("new_sold", "Units", true),
      }});
    }}

    function renderProjectTable(points) {{
      const body = document.getElementById("projectListBody");
      body.innerHTML = "";
      if (!points.length) {{
        const tr = document.createElement("tr");
        tr.innerHTML = '<td colspan="11">No mapped projects for the selected filter.</td>';
        body.appendChild(tr);
        return;
      }}

      const sortedPoints = [...points].sort((a, b) => {{
        const key = state.projectSortKey;
        const aValue = key === "quarterly_sold_rate"
          ? quarterlySoldRateForPoint(a)
          : key === "accumulated_sold_rate"
            ? accumulatedSoldRateForPoint(a)
            : a[key];
        const bValue = key === "quarterly_sold_rate"
          ? quarterlySoldRateForPoint(b)
          : key === "accumulated_sold_rate"
            ? accumulatedSoldRateForPoint(b)
            : b[key];
        let comparison = 0;
        if (typeof aValue === "string" || typeof bValue === "string") {{
          comparison = String(aValue ?? "").localeCompare(String(bValue ?? ""));
        }} else {{
          comparison = Number(aValue ?? 0) - Number(bValue ?? 0);
        }}
        return state.projectSortDirection === "asc" ? comparison : -comparison;
      }});

      sortedPoints.forEach((point) => {{
        const quarterlySoldRate = quarterlySoldRateForPoint(point);
        const accumulatedSoldRate = accumulatedSoldRateForPoint(point);
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><strong>${{point.project_name}}</strong></td>
          <td>${{point.developer || "-"}}</td>
          <td>${{point.city}}</td>
          <td>${{point.segment}}</td>
          <td>${{point.quarter}}</td>
          <td>${{formatNumber(point.new_launched)}}</td>
          <td>${{formatNumber(point.new_sold, "decimal")}}</td>
          <td>${{formatNumber(point.current_supply)}}</td>
          <td>${{formatNumber(point.available_begin)}}</td>
          <td>${{formatNumber(quarterlySoldRate, "percent")}}</td>
          <td>${{formatNumber(accumulatedSoldRate, "percent")}}</td>
          <td>${{formatNumber(point.price, "currency")}}</td>
        `;
        body.appendChild(tr);
      }});
    }}

    function renderCityBoundaries() {{
      cityBoundaryLayers.forEach((layer) => projectMap.removeLayer(layer));
      cityBoundaryLayers = [];

      ["Hanoi", "HCMC"].forEach((city) => {{
        const geometry = dashboardData.city_boundaries?.[city];
        if (!geometry) {{
          return;
        }}
        const layer = L.geoJSON(geometry, {{
          style: {{
          color: cityColors[city],
          weight: 2.5,
          opacity: 0.95,
          fillColor: cityColors[city],
          fillOpacity: 0.04,
          dashArray: "8 6",
          }},
        }})
          .bindTooltip(`${{city}} market boundary`, {{ sticky: true }})
          .addTo(projectMap);
        cityBoundaryLayers.push(layer);
      }});
    }}

    function ensureMap() {{
      if (!projectMap) {{
        projectMap = L.map("projectMap", {{ zoomControl: true }}).setView([16.1, 106.2], 6);
        L.tileLayer("https://{{s}}.basemaps.cartocdn.com/light_nolabels/{{z}}/{{x}}/{{y}}{{r}}.png", {{
          maxZoom: 18,
          attribution: "&copy; OpenStreetMap contributors &copy; CARTO",
        }}).addTo(projectMap);
      }}
    }}

    function renderMap() {{
      ensureMap();
      projectMarkers.forEach((marker) => projectMap.removeLayer(marker));
      projectMarkers = [];
      const latestQuarter = latestMapQuarter();
      document.getElementById("mapNote").textContent = `${{state.mapCityFilter === "All" ? "All markets" : state.mapCityFilter}} / latest quarter / ${{latestQuarter}}`;

      const allPoints = latestMapPoints();
      renderMapAnalytics(allPoints, latestQuarter);
      renderCityBoundaries();
      const points = filteredProjectPoints(allPoints);
      renderProjectTable(points);

      if (!points.length) {{
        projectMap.setView([16.1, 106.2], 6);
        return;
      }}

      const bounds = [];
      points.forEach((point) => {{
        const lat = Number(point.latitude);
        const lng = Number(point.longitude);
        if (!Number.isFinite(lat) || !Number.isFinite(lng)) {{
          return;
        }}
        bounds.push([lat, lng]);
        const marker = L.marker([lat, lng]).addTo(projectMap);
        const quarterlySoldRate = quarterlySoldRateForPoint(point);
        const accumulatedSoldRate = accumulatedSoldRateForPoint(point);
        marker.bindPopup(`
          <strong>${{point.project_name}}</strong><br />
          Developer: ${{point.developer || "-"}}<br />
          Quarter: ${{point.quarter}}<br />
          New launched: ${{formatNumber(point.new_launched)}}<br />
          New sold: ${{formatNumber(point.new_sold, "decimal")}}<br />
          Current supply: ${{formatNumber(point.current_supply)}}<br />
          Quarterly sold rate: ${{formatNumber(quarterlySoldRate, "percent")}}<br />
          Accumulated sold rate: ${{formatNumber(accumulatedSoldRate, "percent")}}<br />
          Price: ${{formatNumber(point.price, "currency")}}
        `);
        projectMarkers.push(marker);
      }});

      if (bounds.length === 1) {{
        projectMap.setView(bounds[0], 12);
      }} else if (bounds.length > 1) {{
        projectMap.fitBounds(bounds, {{ padding: [24, 24] }});
      }}
    }}

    function renderFlowMetricButtons() {{
      renderToggleRow(
        "flowMetricButtons",
        [
          ["new_launched", "New launched"],
          ["new_sold", "New sold"],
        ],
        state.flowMetric,
        (value) => {{
          state.flowMetric = value;
          renderComparisonSections();
        }}
      );
    }}

    function renderComparisonSections() {{
      renderSummary();
      renderFlowCompare();
      renderBreakdownSection();
      renderCompareTable();
      document.getElementById("focusLabel").textContent = `${{state.deepCity}} / ${{state.deepSegment}} / Flow: ${{metricLabels[state.flowMetric]}}`;
    }}

    function renderTimeframeControls() {{
      const startSelect = document.getElementById("startQuarterSelect");
      const endSelect = document.getElementById("endQuarterSelect");
      const eligibleQuarters = dashboardData.quarters.filter(
        (quarter) => quarterSort(quarter) >= quarterSort("2022Q1")
      );
      const options = eligibleQuarters
        .map((quarter) => `<option value="${{quarter}}">${{quarter}}</option>`)
        .join("");
      startSelect.innerHTML = options;
      endSelect.innerHTML = options;
      if (!eligibleQuarters.includes(state.startQuarter)) {{
        state.startQuarter = eligibleQuarters[0];
      }}
      if (!eligibleQuarters.includes(state.endQuarter)) {{
        state.endQuarter = eligibleQuarters[eligibleQuarters.length - 1];
      }}
      startSelect.value = state.startQuarter;
      endSelect.value = state.endQuarter;

      startSelect.onchange = () => {{
        if (quarterSort(startSelect.value) > quarterSort(state.endQuarter)) {{
          state.endQuarter = startSelect.value;
          endSelect.value = state.endQuarter;
        }}
        state.startQuarter = startSelect.value;
      }};
      endSelect.onchange = () => {{
        if (quarterSort(endSelect.value) < quarterSort(state.startQuarter)) {{
          state.startQuarter = endSelect.value;
          startSelect.value = state.startQuarter;
        }}
        state.endQuarter = endSelect.value;
      }};
    }}

    function bindRefreshButton() {{
      const button = document.getElementById("refreshButton");
      button.onclick = () => {{
        renderDashboard();
      }};
    }}

    function renderDashboard() {{
      initializeExportButtons();
      renderFlowMetricButtons();
      renderDeepCityButtons();
      renderDeepSegmentButtons();
      renderProjectCityFilterButtons();
      renderProjectTableHeaders();
      renderTimeframeControls();
      bindRefreshButton();
      renderComparisonSections();
      renderPriceCompare();
      renderDeepDive();
      renderMap();
    }}

    renderDashboard();
  </script>
</body>
</html>
"""


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    dataset = build_dataset()
    OUTPUT_JSON.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    OUTPUT_HTML.write_text(build_html(dataset), encoding="utf-8")
    print(f"Wrote dashboard data to {OUTPUT_JSON}")
    print(f"Wrote dashboard HTML to {OUTPUT_HTML}")


if __name__ == "__main__":
    main()
