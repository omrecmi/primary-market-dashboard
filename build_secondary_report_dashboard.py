from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parent.parent
SHORTCUT_PATH = ROOT / "01. Database_all - Shortcut.lnk"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
OUTPUT_HTML = OUTPUT_DIR / "secondary_report_dashboard.html"
OUTPUT_JSON = OUTPUT_DIR / "secondary_report_dashboard_data.json"

HANOI_FOLDER = "1. Hanoi"
ALL_FILE = "Hanoi Secondary Database all.xlsx"

DEFAULT_START_MONTH = "2025-01"
DEFAULT_END_MONTH = "2026-05"

SPECIAL_SCOPE_DISTRICTS = {
    "Gia Lâm",
    "Đông Anh",
    "Đan Phượng",
}

ASSET_CATEGORY_ORDER = [
    "Cao tầng",
    "Thổ cư",
    "Thấp tầng",
    "Loại hình khác",
]

ASSET_CATEGORY_MAP = {
    "cao_tang": "Cao tầng",
    "tho_cu": "Thổ cư",
    "thap_tang": "Thấp tầng",
}

REGION_ORDER = [
    "Khu Đông",
    "Khu Nam",
    "Khu Bắc",
    "Khu Trung tâm",
    "Nội thành",
    "Khu Tây",
]


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


def clean_text(value: Any, fallback: str = "Unknown") -> str:
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return fallback
    return text


def normalize_geo_name(value: Any) -> str:
    text = clean_text(value, "")
    prefixes = ("Q. ", "H. ", "P. ", "X. ", "TT. ", "TX. ")
    for prefix in prefixes:
        if text.startswith(prefix):
            return text[len(prefix) :].strip()
    return text


def is_scope_district(value: Any) -> bool:
    text = clean_text(value, "")
    if text.startswith("Q. "):
        return True
    return normalize_geo_name(text) in SPECIAL_SCOPE_DISTRICTS


def build_year_month(dataframe: pd.DataFrame) -> pd.Series:
    year = pd.to_numeric(dataframe["date_y"], errors="coerce").astype("Int64")
    month = pd.to_numeric(dataframe["date_m"], errors="coerce").astype("Int64")
    valid = year.notna() & month.notna() & month.between(1, 12)
    output = pd.Series(pd.NA, index=dataframe.index, dtype="object")
    output.loc[valid] = (
        year.loc[valid].astype(int).astype(str)
        + "-"
        + month.loc[valid].astype(int).astype(str).str.zfill(2)
    )
    return output


def month_label(year_month: str) -> str:
    year, month = year_month.split("-")
    return f"T{int(month)}/{year}"


def month_sort_key(year_month: str) -> int:
    year, month = year_month.split("-")
    return int(year) * 100 + int(month)


def previous_month(year_month: str) -> str:
    year, month = [int(part) for part in year_month.split("-")]
    if month == 1:
        return f"{year - 1}-12"
    return f"{year}-{month - 1:02d}"


def previous_year(year_month: str) -> str:
    year, month = [int(part) for part in year_month.split("-")]
    return f"{year - 1}-{month:02d}"


def normalize_asset(value: Any) -> str:
    text = clean_text(value, "other")
    return ASSET_CATEGORY_MAP.get(text, "Loại hình khác")


def normalize_region(value: Any) -> str:
    text = clean_text(value, "Khu vực khác")
    aliases = {
        "Khu trung tâm": "Khu Trung tâm",
        "Trung tâm": "Khu Trung tâm",
        "Khu Nội thành": "Nội thành",
    }
    return aliases.get(text, text)


def unique_count_frame(
    dataframe: pd.DataFrame,
    group_columns: list[str],
    value_column: str = "scd2_hash",
) -> pd.DataFrame:
    return (
        dataframe.groupby(group_columns, dropna=False)[value_column]
        .nunique()
        .reset_index(name="transactions")
    )


def make_monthly_lookup(dataframe: pd.DataFrame, category_col: str) -> dict[tuple[str, str], int]:
    grouped = unique_count_frame(dataframe, ["year_month", category_col])
    return {
        (str(row["year_month"]), str(row[category_col])): int(row["transactions"])
        for _, row in grouped.iterrows()
    }


def load_secondary_all() -> pd.DataFrame:
    base_dir = resolve_shortcut(SHORTCUT_PATH)
    path = base_dir / HANOI_FOLDER / ALL_FILE
    dataframe = pd.read_excel(
        path,
        usecols=[
            "date_y",
            "date_m",
            "scd2_hash",
            "loai_tai_san",
            "district",
            "region_new1",
        ],
    )
    dataframe["scd2_hash"] = dataframe["scd2_hash"].map(lambda value: clean_text(value, ""))
    dataframe["year_month"] = build_year_month(dataframe)
    dataframe["asset_category"] = dataframe["loai_tai_san"].map(normalize_asset)
    dataframe["district"] = dataframe["district"].map(lambda value: clean_text(value, "Unknown"))
    dataframe["region"] = dataframe["region_new1"].map(normalize_region)
    dataframe = dataframe[
        dataframe["scd2_hash"].ne("")
        & dataframe["year_month"].notna()
        & dataframe["district"].map(is_scope_district)
    ].copy()
    dataframe["year_month"] = dataframe["year_month"].astype(str)
    return dataframe


def build_dashboard_data(dataframe: pd.DataFrame) -> dict[str, Any]:
    months = sorted(dataframe["year_month"].unique().tolist(), key=month_sort_key)

    asset_lookup = make_monthly_lookup(dataframe, "asset_category")
    monthly_totals = (
        dataframe.groupby("year_month")["scd2_hash"].nunique().reindex(months).fillna(0).astype(int)
    )

    monthly_records: list[dict[str, Any]] = []
    for month in months:
        counts = {asset: asset_lookup.get((month, asset), 0) for asset in ASSET_CATEGORY_ORDER}
        total = int(monthly_totals.loc[month])
        prev_total = int(monthly_totals.get(previous_month(month), 0))
        mom = ((total / prev_total) - 1) if prev_total else None
        monthly_records.append(
            {
                "month": month,
                "label": month_label(month),
                "total": total,
                "mom": mom,
                "asset_counts": counts,
            }
        )

    line_records: list[dict[str, Any]] = []
    for month in months:
        row: dict[str, Any] = {"month": month, "label": month_label(month)}
        for asset in ["Cao tầng", "Thổ cư", "Thấp tầng"]:
            current = asset_lookup.get((month, asset), 0)
            previous = asset_lookup.get((previous_month(month), asset), 0)
            year_ago = asset_lookup.get((previous_year(month), asset), 0)
            row[asset] = {
                "transactions": current,
                "mom": ((current / previous) - 1) if previous else None,
                "yoy": ((current / year_ago) - 1) if year_ago else None,
            }
        line_records.append(row)

    region_data: dict[str, list[dict[str, Any]]] = {}
    for asset_key, output_key in [("Cao tầng", "high_rise"), ("Thổ cư", "landed_house")]:
        scoped = dataframe[
            dataframe["asset_category"].eq(asset_key) & dataframe["region"].isin(REGION_ORDER)
        ].copy()
        region_lookup = make_monthly_lookup(scoped, "region")
        records: list[dict[str, Any]] = []
        for month in months:
            counts = {region: region_lookup.get((month, region), 0) for region in REGION_ORDER}
            total = sum(counts.values())
            shares = {
                region: (counts[region] / total if total else 0)
                for region in REGION_ORDER
            }
            records.append(
                {
                    "month": month,
                    "label": month_label(month),
                    "total": total,
                    "counts": counts,
                    "shares": shares,
                }
            )
        region_data[output_key] = records

    visible_months = [
        month for month in months if DEFAULT_START_MONTH <= month <= DEFAULT_END_MONTH
    ]
    default_start = DEFAULT_START_MONTH if DEFAULT_START_MONTH in months else months[0]
    default_end = DEFAULT_END_MONTH if DEFAULT_END_MONTH in months else months[-1]
    if not visible_months:
        default_start, default_end = months[0], months[-1]

    return {
        "meta": {
            "source_file": ALL_FILE,
            "scope": "Only Q. districts plus H. Gia Lâm, H. Đông Anh, H. Đan Phượng",
            "months": months,
            "default_start": default_start,
            "default_end": default_end,
            "asset_order": ASSET_CATEGORY_ORDER,
            "line_asset_order": ["Cao tầng", "Thổ cư", "Thấp tầng"],
            "region_order": REGION_ORDER,
        },
        "monthly_asset_mix": monthly_records,
        "asset_line_trend": line_records,
        "region_share": region_data,
    }


HTML_TEMPLATE = r"""<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Secondary Report Chart Dashboard</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800;900&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
  <style>
    :root {
      --paper: #f7f1e7;
      --ink: #1d2528;
      --muted: #677174;
      --line: rgba(29, 37, 40, 0.12);
      --card: rgba(255, 253, 248, 0.92);
      --blue-soft: #75b7e5;
      --blue-dark: #145b8f;
      --green: #5aa469;
      --pink: #e26c7c;
      --gray: #a8adb2;
      --shadow: 0 24px 70px rgba(66, 48, 20, 0.12);
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% 8%, rgba(117, 183, 229, 0.23), transparent 34rem),
        radial-gradient(circle at 88% 4%, rgba(226, 108, 124, 0.16), transparent 28rem),
        linear-gradient(135deg, #fbf8ef 0%, var(--paper) 45%, #eef5f0 100%);
      font-family: "Manrope", sans-serif;
    }

    main {
      width: min(1520px, calc(100vw - 44px));
      margin: 0 auto;
      padding: 34px 0 60px;
    }

    .hero {
      display: grid;
      grid-template-columns: 1.25fr 0.75fr;
      gap: 24px;
      align-items: stretch;
      margin-bottom: 24px;
    }

    .panel {
      background: var(--card);
      border: 1px solid rgba(29, 37, 40, 0.1);
      border-radius: 28px;
      box-shadow: var(--shadow);
      backdrop-filter: blur(14px);
    }

    .intro {
      padding: 34px 38px;
      position: relative;
      overflow: hidden;
    }

    .intro::after {
      content: "";
      position: absolute;
      right: -80px;
      bottom: -90px;
      width: 280px;
      height: 280px;
      border-radius: 50%;
      background: conic-gradient(from 130deg, rgba(20, 91, 143, 0.35), rgba(90, 164, 105, 0.24), rgba(226, 108, 124, 0.22), rgba(20, 91, 143, 0.35));
      opacity: 0.48;
    }

    .eyebrow {
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
      letter-spacing: 0.18em;
      text-transform: uppercase;
    }

    h1 {
      margin: 10px 0 16px;
      font-size: clamp(42px, 6vw, 86px);
      line-height: 0.9;
      letter-spacing: -0.06em;
    }

    .intro p {
      max-width: 760px;
      margin: 0;
      color: #445054;
      font-size: 17px;
      line-height: 1.65;
      position: relative;
      z-index: 1;
    }

    .controls {
      padding: 24px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      gap: 18px;
    }

    .control-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 14px;
    }

    .style-controls {
      display: grid;
      grid-template-columns: 1fr;
      gap: 12px;
      padding-top: 4px;
    }

    .range-row {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 10px;
      align-items: center;
    }

    input[type="range"] {
      width: 100%;
      accent-color: #1d2528;
    }

    .range-value {
      min-width: 58px;
      color: var(--ink);
      font-size: 12px;
      font-weight: 900;
      text-align: right;
    }

    .button-row {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }

    .button-row button {
      flex: 1 1 86px;
      padding: 10px 12px;
      border: 1px solid rgba(29, 37, 40, 0.16);
      background: #fffdf7;
      color: #1d2528;
    }

    label {
      display: block;
      margin-bottom: 7px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.1em;
      text-transform: uppercase;
    }

    select, button {
      width: 100%;
      border: 1px solid rgba(29, 37, 40, 0.16);
      border-radius: 16px;
      background: #fffdf7;
      color: var(--ink);
      font: inherit;
      font-weight: 750;
      padding: 13px 14px;
    }

    button {
      border: none;
      background: #1d2528;
      color: #fffaf0;
      cursor: pointer;
      transition: transform 0.18s ease, box-shadow 0.18s ease;
    }

    button:hover {
      transform: translateY(-1px);
      box-shadow: 0 12px 28px rgba(29, 37, 40, 0.18);
    }

    .note {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }

    .stats {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 14px;
      margin-bottom: 24px;
    }

    .stat {
      padding: 18px 20px;
    }

    .stat span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }

    .stat strong {
      display: block;
      margin-top: 6px;
      font-size: 31px;
      letter-spacing: -0.04em;
    }

    .chart-grid {
      display: grid;
      grid-template-columns: 1fr;
      gap: 24px;
    }

    .chart-card {
      padding: 24px;
    }

    .chart-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 18px;
      margin-bottom: 12px;
    }

    h2 {
      margin: 0 0 8px;
      font-size: 24px;
      letter-spacing: -0.03em;
    }

    .chart-head p {
      max-width: 900px;
      margin: 0;
      color: var(--muted);
      line-height: 1.55;
    }

    .chart-wrap {
      width: 100%;
      height: 460px;
      margin-top: 12px;
    }

    .short-chart {
      height: 410px;
    }

    .mom-strip {
      display: grid;
      grid-template-columns: 72px 1fr;
      align-items: stretch;
      gap: 8px;
      margin-top: -4px;
    }

    .mom-label {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      color: #000;
      font-size: 13px;
      font-style: italic;
      font-weight: 900;
    }

    .mom-wrap {
      height: 96px;
      margin-top: 0;
    }

    .legend-pill {
      flex: 0 0 auto;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 9px 13px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
      background: rgba(255,255,255,0.58);
    }

    @media (max-width: 960px) {
      main { width: min(100vw - 24px, 760px); }
      .hero { grid-template-columns: 1fr; }
      .stats { grid-template-columns: repeat(2, 1fr); }
      .chart-head { flex-direction: column; }
      .chart-wrap { height: 390px; }
    }

    @media (max-width: 640px) {
      .control-grid, .stats { grid-template-columns: 1fr; }
      .intro, .controls, .chart-card { padding: 20px; }
      h1 { font-size: 44px; }
    }
  </style>
</head>
<body>
  <main>
    <section class="hero">
      <div class="panel intro">
        <div class="eyebrow">Hanoi Secondary Market</div>
        <h1>Report Chart<br/>Dashboard</h1>
        <p>
          Dashboard này tạo nhanh 4 chart cho phần Tổng quan thị trường thứ cấp:
          tổng giao dịch theo loại hình, trend 3 loại hình chính, cơ cấu cao tầng theo khu vực,
          và cơ cấu thổ cư theo khu vực. Metric chính là count distinct <code>scd2_hash</code>.
        </p>
      </div>
      <aside class="panel controls">
        <div>
          <div class="control-grid">
            <div>
              <label for="startMonth">Tháng bắt đầu</label>
              <select id="startMonth"></select>
            </div>
            <div>
              <label for="endMonth">Tháng kết thúc</label>
              <select id="endMonth"></select>
            </div>
          </div>
        </div>
        <button id="refreshBtn">Refresh charts</button>
        <div class="style-controls">
          <div>
            <label for="chartControlTarget">Chart cần chỉnh</label>
            <select id="chartControlTarget">
              <option value="assetStack">Chart 1 - Tổng GD theo loại hình</option>
              <option value="assetLines">Chart 2 - Line 3 loại hình</option>
              <option value="highRiseRegion">Chart 3 - Cao tầng theo khu vực</option>
              <option value="landedRegion">Chart 4 - Thổ cư theo khu vực</option>
            </select>
          </div>
          <div>
            <label>Cỡ chữ chart</label>
            <div class="button-row">
              <button id="decreaseChartFont" type="button">A-</button>
              <button id="increaseChartFont" type="button">A+</button>
              <span class="range-value" id="fontSizeValue">12px</span>
            </div>
          </div>
          <div>
            <label>Chart width</label>
            <div class="button-row">
              <button id="decreaseChartWidth" type="button">W-</button>
              <button id="increaseChartWidth" type="button">W+</button>
              <span class="range-value" id="chartWidthValue">100%</span>
            </div>
          </div>
          <div>
            <label>Chart height</label>
            <div class="button-row">
              <button id="decreaseChartHeight" type="button">H-</button>
              <button id="increaseChartHeight" type="button">H+</button>
              <span class="range-value" id="chartHeightValue">460px</span>
            </div>
          </div>
          <div>
            <label>Total label</label>
            <div class="button-row">
              <button id="toggleChartTotals" type="button">Ẩn total</button>
            </div>
          </div>
        </div>
        <div class="note" id="scopeNote"></div>
      </aside>
    </section>

    <section class="stats">
      <div class="panel stat"><span>Tổng GD trong kỳ</span><strong id="statTotal">-</strong></div>
      <div class="panel stat"><span>TB / tháng</span><strong id="statAvg">-</strong></div>
      <div class="panel stat"><span>Tháng cao nhất</span><strong id="statPeak">-</strong></div>
      <div class="panel stat"><span>Latest month</span><strong id="statLatest">-</strong></div>
    </section>

    <section class="chart-grid">
      <article class="panel chart-card">
        <div class="chart-head">
          <div>
            <h2>1. Tổng lượng giao dịch thứ cấp theo loại hình, theo tháng</h2>
            <p>Stacked column theo 4 loại hình, nhãn tổng ở đỉnh cột, nhãn tỷ trọng bên trong, đường TB và line MoM ở trục dưới.</p>
          </div>
          <div class="legend-pill">Stacked + Average + MoM</div>
        </div>
        <div class="chart-wrap" id="wrapAssetStack"><canvas id="chartAssetStack"></canvas></div>
        <div class="mom-strip">
          <div class="mom-label">MoM, %</div>
          <div class="chart-wrap mom-wrap" id="wrapAssetMom"><canvas id="chartAssetMom"></canvas></div>
        </div>
      </article>

      <article class="panel chart-card">
        <div class="chart-head">
          <div>
            <h2>2. Tổng lượng giao dịch thứ cấp cao tầng & thổ cư theo tháng</h2>
            <p>Line chart cho Cao tầng, Thổ cư, Thấp tầng; điểm cuối có nhãn số lượng kèm MoM và YoY.</p>
          </div>
          <div class="legend-pill">Line trend</div>
        </div>
        <div class="chart-wrap short-chart" id="wrapAssetLines"><canvas id="chartAssetLines"></canvas></div>
      </article>

      <article class="panel chart-card">
        <div class="chart-head">
          <div>
            <h2>3. Tỷ lệ giao dịch thứ cấp cao tầng theo khu vực</h2>
            <p>100% stacked column theo 6 khu vực, dùng để theo dõi thị phần giao dịch cao tầng giữa các tháng.</p>
          </div>
          <div class="legend-pill">100% stacked</div>
        </div>
        <div class="chart-wrap short-chart" id="wrapHighRiseRegion"><canvas id="chartHighRiseRegion"></canvas></div>
      </article>

      <article class="panel chart-card">
        <div class="chart-head">
          <div>
            <h2>4. Tỷ lệ giao dịch thổ cư theo khu vực</h2>
            <p>100% stacked column theo cùng 6 khu vực, cho thấy phân bổ giao dịch thổ cư đồng đều hay lệch vùng.</p>
          </div>
          <div class="legend-pill">100% stacked</div>
        </div>
        <div class="chart-wrap short-chart" id="wrapLandedRegion"><canvas id="chartLandedRegion"></canvas></div>
      </article>
    </section>
  </main>

  <script>
    const dashboardData = __DASHBOARD_DATA__;
    Chart.register(ChartDataLabels);
    Chart.defaults.font.family = "Manrope";
    Chart.defaults.color = "#000";

    const assetColors = {
      "Cao tầng": "#75b7e5",
      "Thổ cư": "#5aa469",
      "Thấp tầng": "#e26c7c",
      "Loại hình khác": "#a8adb2"
    };

    const regionColors = {
      "Khu Đông": "#145b8f",
      "Khu Nam": "#3a86c8",
      "Khu Bắc": "#89c4e8",
      "Khu Trung tâm": "#7a9a9a",
      "Nội thành": "#5aa469",
      "Khu Tây": "#e26c7c"
    };

    let charts = {};
    const chartSettings = {
      assetStack: { fontSize: 12, width: 100, height: 460, showTotals: true },
      assetLines: { fontSize: 12, width: 100, height: 410, showTotals: true },
      highRiseRegion: { fontSize: 12, width: 100, height: 410, showTotals: true },
      landedRegion: { fontSize: 12, width: 100, height: 410, showTotals: true }
    };

    const chartWraps = {
      assetStack: ["wrapAssetStack", "wrapAssetMom"],
      assetLines: ["wrapAssetLines"],
      highRiseRegion: ["wrapHighRiseRegion"],
      landedRegion: ["wrapLandedRegion"]
    };

    function formatK(value) {
      if (value === null || value === undefined || Number.isNaN(value)) return "-";
      const abs = Math.abs(value);
      if (abs >= 1000) return `${(value / 1000).toFixed(abs >= 10000 ? 0 : 1)}K`;
      return `${Math.round(value)}`;
    }

    function formatPct(value, digits = 0) {
      if (value === null || value === undefined || Number.isNaN(value)) return "-";
      const sign = value > 0 ? "+" : "";
      return `${sign}${(value * 100).toFixed(digits)}%`;
    }

    function formatDelta(value) {
      if (value === null || value === undefined || Number.isNaN(value)) return "n/a";
      const arrow = value > 0 ? "▲" : value < 0 ? "▼" : "■";
      return `${arrow}${Math.abs(value * 100).toFixed(0)}%`;
    }

    function monthInRange(row, start, end) {
      return row.month >= start && row.month <= end;
    }

    function destroyChart(id) {
      if (charts[id]) {
        charts[id].destroy();
        delete charts[id];
      }
    }

    function buildChart(id, config) {
      destroyChart(id);
      charts[id] = new Chart(document.getElementById(id), config);
    }

    function selectedChartKey() {
      return document.getElementById("chartControlTarget").value;
    }

    function syncStyleControls() {
      const settings = chartSettings[selectedChartKey()];
      document.getElementById("fontSizeValue").textContent = `${settings.fontSize}px`;
      document.getElementById("chartWidthValue").textContent = `${settings.width}%`;
      document.getElementById("chartHeightValue").textContent = `${settings.height}px`;
      document.getElementById("toggleChartTotals").textContent = settings.showTotals ? "Ẩn total" : "Hiện total";
    }

    function applyChartSizing() {
      for (const [key, wrapIds] of Object.entries(chartWraps)) {
        const settings = chartSettings[key];
        for (const wrapId of wrapIds) {
          const element = document.getElementById(wrapId);
          element.style.width = `${settings.width}%`;
          element.style.marginLeft = "auto";
          element.style.marginRight = "auto";
          element.style.height = wrapId === "wrapAssetMom"
            ? `${Math.max(74, Math.round(settings.height * 0.22))}px`
            : `${settings.height}px`;
        }
      }
    }

    function populateMonthSelects() {
      const months = dashboardData.meta.months;
      for (const id of ["startMonth", "endMonth"]) {
        const select = document.getElementById(id);
        select.innerHTML = months.map(month => `<option value="${month}">${labelForMonth(month)}</option>`).join("");
      }
      document.getElementById("startMonth").value = dashboardData.meta.default_start;
      document.getElementById("endMonth").value = dashboardData.meta.default_end;
      document.getElementById("scopeNote").innerHTML =
        `Nguồn: <strong>${dashboardData.meta.source_file}</strong><br/>Scope: ${dashboardData.meta.scope}`;
      syncStyleControls();
      applyChartSizing();
    }

    function labelForMonth(month) {
      const [year, monthNumber] = month.split("-");
      return `T${Number(monthNumber)}/${year}`;
    }

    function selectedRange() {
      let start = document.getElementById("startMonth").value;
      let end = document.getElementById("endMonth").value;
      if (start > end) {
        [start, end] = [end, start];
        document.getElementById("startMonth").value = start;
        document.getElementById("endMonth").value = end;
      }
      return { start, end };
    }

    function wireStyleControls() {
      document.getElementById("chartControlTarget").addEventListener("change", syncStyleControls);
      document.getElementById("decreaseChartFont").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].fontSize = Math.max(9, chartSettings[key].fontSize - 1);
        syncStyleControls();
        renderAll();
      });
      document.getElementById("increaseChartFont").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].fontSize = Math.min(22, chartSettings[key].fontSize + 1);
        syncStyleControls();
        renderAll();
      });
      document.getElementById("decreaseChartWidth").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].width = Math.max(60, chartSettings[key].width - 5);
        syncStyleControls();
        renderAll();
      });
      document.getElementById("increaseChartWidth").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].width = Math.min(120, chartSettings[key].width + 5);
        syncStyleControls();
        renderAll();
      });
      document.getElementById("decreaseChartHeight").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].height = Math.max(260, chartSettings[key].height - 40);
        syncStyleControls();
        renderAll();
      });
      document.getElementById("increaseChartHeight").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].height = Math.min(900, chartSettings[key].height + 40);
        syncStyleControls();
        renderAll();
      });
      document.getElementById("toggleChartTotals").addEventListener("click", () => {
        const key = selectedChartKey();
        chartSettings[key].showTotals = !chartSettings[key].showTotals;
        syncStyleControls();
        renderAll();
      });
    }

    const totalLabelPlugin = {
      id: "totalLabelPlugin",
      afterDatasetsDraw(chart, args, options) {
        if (!options || !options.display || !Array.isArray(options.totals)) return;
        const { ctx, chartArea, scales } = chart;
        const xScale = scales.x;
        const yScale = scales.y;
        if (!xScale || !yScale) return;
        ctx.save();
        ctx.fillStyle = "#000";
        ctx.textAlign = "center";
        ctx.textBaseline = "bottom";
        ctx.font = `900 ${Math.max(10, options.fontSize || 12)}px Manrope`;
        options.totals.forEach((value, index) => {
          if (!value && value !== 0) return;
          const x = xScale.getPixelForValue(index);
          const y = options.fixedTop
            ? chartArea.top + 18
            : Math.max(chartArea.top + 16, yScale.getPixelForValue(options.percentChart ? 100 : value) - 8);
          ctx.fillText(formatK(value), x, y);
        });
        ctx.restore();
      }
    };
    Chart.register(totalLabelPlugin);

    function chartFont(chartKey, extra = {}) {
      return {
        family: "Manrope",
        size: chartSettings[chartKey].fontSize,
        weight: 900,
        ...extra
      };
    }

    function baseOptions(yTitle, stacked = false, chartKey = "assetStack") {
      const fontSize = chartSettings[chartKey].fontSize;
      return {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        layout: { padding: { top: 16, right: 18, bottom: 6, left: 6 } },
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              color: "#000",
              usePointStyle: true,
              boxWidth: 8,
              boxHeight: 8,
              padding: 18,
              font: chartFont(chartKey, { size: fontSize, weight: 900 })
            }
          },
          tooltip: {
            titleFont: chartFont(chartKey),
            bodyFont: chartFont(chartKey, { weight: 800 }),
            callbacks: {
              label(context) {
                const label = context.dataset.label || "";
                const value = context.parsed.y;
                if (context.dataset.yAxisID === "mom") return `${label}: ${formatPct(value, 0)}`;
                return `${label}: ${formatK(value)}`;
              }
            }
          }
        },
        scales: {
          x: {
            stacked,
            grid: { display: false },
            ticks: {
              color: "#000",
              maxRotation: 0,
              autoSkip: true,
              font: chartFont(chartKey)
            }
          },
          y: {
            stacked,
            beginAtZero: true,
            title: {
              display: true,
              text: yTitle,
              color: "#000",
              font: chartFont(chartKey, { style: "italic" })
            },
            ticks: {
              color: "#000",
              callback: value => formatK(value),
              font: chartFont(chartKey)
            },
            grid: { color: "rgba(29,37,40,0.08)" }
          }
        }
      };
    }

    function updateStats(rows) {
      const total = rows.reduce((sum, row) => sum + row.total, 0);
      const avg = rows.length ? total / rows.length : 0;
      const peak = rows.reduce((best, row) => !best || row.total > best.total ? row : best, null);
      const latest = rows[rows.length - 1];
      document.getElementById("statTotal").textContent = formatK(total);
      document.getElementById("statAvg").textContent = formatK(avg);
      document.getElementById("statPeak").textContent = peak ? `${peak.label} (${formatK(peak.total)})` : "-";
      document.getElementById("statLatest").textContent = latest ? `${latest.label} (${formatK(latest.total)})` : "-";
    }

    function renderAssetStack(rows) {
      const labels = rows.map(row => row.label);
      const average = rows.length ? rows.reduce((sum, row) => sum + row.total, 0) / rows.length : 0;
      const totals = rows.map(row => row.total);
      const datasets = dashboardData.meta.asset_order.map(asset => ({
        type: "bar",
        label: asset,
        data: rows.map(row => row.asset_counts[asset] || 0),
        backgroundColor: assetColors[asset],
        borderColor: "rgba(255,255,255,0.92)",
        borderWidth: 1,
        stack: "asset"
      }));
      datasets.push({
        type: "line",
        label: `TB: ${formatK(average)} căn/tháng`,
        data: rows.map(() => average),
        borderColor: "#145b8f",
        borderDash: [7, 6],
        borderWidth: 2,
        pointRadius: 0,
        yAxisID: "y"
      });
      const options = baseOptions("Số giao dịch (K)", true, "assetStack");
      options.plugins.datalabels = {
        color(context) {
          return "#000";
        },
        anchor(context) {
          if (context.dataset.label.startsWith("TB")) return "end";
          const datasetIndex = context.datasetIndex;
          const isTopBar = datasetIndex === dashboardData.meta.asset_order.length - 1;
          return isTopBar ? "end" : "center";
        },
        align(context) {
          if (context.dataset.label.startsWith("TB")) return "top";
          const datasetIndex = context.datasetIndex;
          const isTopBar = datasetIndex === dashboardData.meta.asset_order.length - 1;
          return isTopBar ? "top" : "center";
        },
        offset: 2,
        clamp: true,
        display(context) {
          if (context.dataset.label.startsWith("TB")) return false;
          const value = context.dataset.data[context.dataIndex] || 0;
          const total = totals[context.dataIndex] || 0;
          if (!total) return false;
          const share = value / total;
          return share >= 0.08;
        },
        formatter(value, context) {
          const total = totals[context.dataIndex] || 0;
          return total ? `${Math.round((value / total) * 100)}%` : "";
        },
        font: chartFont("assetStack", { size: Math.max(9, chartSettings.assetStack.fontSize - 1) })
      };
      options.plugins.totalLabelPlugin = {
        display: chartSettings.assetStack.showTotals,
        totals,
        fontSize: chartSettings.assetStack.fontSize
      };
      buildChart("chartAssetStack", { data: { labels, datasets }, options });
    }

    function renderMomLine(rows) {
      const labels = rows.map(row => row.label);
      const fontSize = chartSettings.assetStack.fontSize;
      const values = rows.map(row => row.mom);
      const options = {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        layout: { padding: { top: 18, right: 18, bottom: 0, left: 6 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            titleFont: chartFont("assetStack"),
            bodyFont: chartFont("assetStack"),
            callbacks: { label: context => `MoM: ${formatPct(context.parsed.y, 0)}` }
          },
          datalabels: {
            display: context => context.raw !== null && context.raw !== undefined,
            align: context => context.parsed.y >= 0 ? "top" : "bottom",
            anchor: context => context.parsed.y >= 0 ? "end" : "start",
            offset: 6,
            clip: false,
            clamp: true,
            color: context => context.parsed.y >= 0 ? "#5aa469" : "#e26c7c",
            formatter: value => formatPct(value, 0),
            font: chartFont("assetStack", { size: Math.max(9, fontSize - 1) })
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { display: false, color: "#000", font: chartFont("assetStack") },
            border: { display: false }
          },
          y: {
            display: false,
            suggestedMin: Math.min(-0.6, ...values.filter(value => value !== null)),
            suggestedMax: Math.max(0.6, ...values.filter(value => value !== null)),
            grid: { display: false },
            border: { display: false }
          }
        }
      };
      buildChart("chartAssetMom", {
        type: "line",
        data: {
          labels,
          datasets: [{
            label: "MoM (%)",
            data: values,
            borderColor: "#75b7e5",
            backgroundColor: "#5aa469",
            pointBackgroundColor: rows.map(row => (row.mom || 0) >= 0 ? "#5aa469" : "#e26c7c"),
            pointBorderColor: "#ffffff",
            pointBorderWidth: 1,
            pointRadius: 4,
            pointHoverRadius: 6,
            borderWidth: 2,
            tension: 0.35
          }]
        },
        options
      });
    }

    function renderAssetLines(rows) {
      const labels = rows.map(row => row.label);
      const assets = dashboardData.meta.line_asset_order;
      const datasets = assets.map(asset => ({
        label: asset,
        data: rows.map(row => row[asset].transactions),
        borderColor: assetColors[asset],
        backgroundColor: assetColors[asset],
        borderWidth: 3,
        pointRadius: 3,
        pointHoverRadius: 6,
        tension: 0.35
      }));
      const options = baseOptions("Số giao dịch (K)", false, "assetLines");
      options.plugins.datalabels = {
        display(context) {
          return context.dataIndex === rows.length - 1;
        },
        align: "right",
        anchor: "end",
        offset: 6,
        color: "#1d2528",
        backgroundColor: "rgba(255,253,248,0.9)",
        borderColor: "rgba(29,37,40,0.14)",
        borderWidth: 1,
        borderRadius: 8,
        padding: 6,
        formatter(value, context) {
          const asset = context.dataset.label;
          const last = rows[rows.length - 1][asset];
          return `${asset}: ${formatK(value)} (${formatDelta(last.mom)} MoM, ${formatDelta(last.yoy)} YoY)`;
        },
        font: chartFont("assetLines", { size: Math.max(9, chartSettings.assetLines.fontSize - 1) })
      };
      options.plugins.totalLabelPlugin = {
        display: chartSettings.assetLines.showTotals,
        totals: rows.map(row => dashboardData.meta.line_asset_order.reduce((sum, asset) => sum + (row[asset]?.transactions || 0), 0)),
        fontSize: chartSettings.assetLines.fontSize,
        fixedTop: true
      };
      buildChart("chartAssetLines", { type: "line", data: { labels, datasets }, options });
    }

    function renderRegionShare(chartId, rows) {
      const chartKey = chartId === "chartHighRiseRegion" ? "highRiseRegion" : "landedRegion";
      const labels = rows.map(row => row.label);
      const datasets = dashboardData.meta.region_order.map(region => ({
        label: region,
        data: rows.map(row => (row.shares[region] || 0) * 100),
        backgroundColor: regionColors[region],
        borderColor: "rgba(255,255,255,0.95)",
        borderWidth: 1,
        stack: "region"
      }));
      const options = baseOptions("Tỷ trọng", true, chartKey);
      options.scales.y.max = 100;
      options.scales.y.ticks = {
        color: "#000",
        callback: value => `${value}%`,
        font: chartFont(chartKey)
      };
      options.plugins.tooltip.callbacks.label = context => `${context.dataset.label}: ${context.parsed.y.toFixed(1)}%`;
      options.plugins.datalabels = {
        color: "#000",
        display(context) {
          return context.dataset.data[context.dataIndex] >= 6;
        },
        formatter(value) {
          return `${Math.round(value)}%`;
        },
        font: chartFont(chartKey, { size: Math.max(9, chartSettings[chartKey].fontSize - 1) })
      };
      options.plugins.totalLabelPlugin = {
        display: chartSettings[chartKey].showTotals,
        totals: rows.map(row => row.total),
        fontSize: chartSettings[chartKey].fontSize,
        percentChart: true
      };
      buildChart(chartId, { type: "bar", data: { labels, datasets }, options });
    }

    function renderAll() {
      const { start, end } = selectedRange();
      const monthlyRows = dashboardData.monthly_asset_mix.filter(row => monthInRange(row, start, end));
      const lineRows = dashboardData.asset_line_trend.filter(row => monthInRange(row, start, end));
      const highRiseRows = dashboardData.region_share.high_rise.filter(row => monthInRange(row, start, end));
      const landedRows = dashboardData.region_share.landed_house.filter(row => monthInRange(row, start, end));
      updateStats(monthlyRows);
      applyChartSizing();
      renderAssetStack(monthlyRows);
      renderMomLine(monthlyRows);
      renderAssetLines(lineRows);
      renderRegionShare("chartHighRiseRegion", highRiseRows);
      renderRegionShare("chartLandedRegion", landedRows);
    }

    populateMonthSelects();
    wireStyleControls();
    document.getElementById("refreshBtn").addEventListener("click", renderAll);
    document.getElementById("startMonth").addEventListener("change", renderAll);
    document.getElementById("endMonth").addEventListener("change", renderAll);
    renderAll();
  </script>
</body>
</html>
"""


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    dataframe = load_secondary_all()
    dashboard_data = build_dashboard_data(dataframe)
    OUTPUT_JSON.write_text(
        json.dumps(dashboard_data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    html = HTML_TEMPLATE.replace(
        "__DASHBOARD_DATA__",
        json.dumps(dashboard_data, ensure_ascii=False),
    )
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"Wrote {OUTPUT_HTML}")
    print(f"Wrote {OUTPUT_JSON}")


if __name__ == "__main__":
    main()
