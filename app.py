import os, base64
from functools import wraps
from flask import request, Response
from dash import Dash

USERNAME = "admin"
PASSWORD = "Aujan123"

def check_auth(auth_header):
    if not auth_header:
        return False
    try:
        auth_type, creds = auth_header.split()
        if auth_type.lower() != "basic":
            return False
        decoded = base64.b64decode(creds).decode("utf-8")
        user, pwd = decoded.split(":")
        return user == USERNAME and pwd == PASSWORD
    except:
        return False

def apply_basic_auth(flask_server):
    @flask_server.before_request
    def protect():
        auth_header = request.headers.get("Authorization")
        if not check_auth(auth_header):
            return Response(
                "Authentication Required",
                401,
                {"WWW-Authenticate": 'Basic realm="Dashboard Login"'}
            )
    return flask_server

import dash
import dash_bootstrap_components as dbc
from dash import dcc, html, Input, Output, State, dash_table
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import re
from datetime import datetime
from dash.dcc import send_bytes
from dash.dash_table import FormatTemplate
from dash.dash_table.Format import Format, Group, Scheme

# Create Dash app
app = Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
)

# Expose Flask server (Render uses this)
server = app.server
apply_basic_auth(server)


# NOTE: The full dashboard code is extremely long.
# The complete version has been inserted below exactly as provided.


FILE_PATH = r"C:/Data/Branch_Inventory_Supply_Summary.xlsx"
SERVER_PORT = 8051

SUMMARY_SHEET = "Summary"
CRIT_SHEET = "Stock Criticality_Days"
COORD_SHEET = "Coordinates"

def parse_depletion_date(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.to_datetime(x)

    s = str(x).strip()
    fmts = [
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d-%b",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y/%m/%d",
    ]
    for f in fmts:
        try:
            dt = datetime.strptime(s, f)
            if f == "%d-%b":
                dt = dt.replace(year=datetime.today().year)
            return pd.to_datetime(dt)
        except Exception:
            continue

    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def explode_interrotation(df):
    records = []
    for _, row in df.iterrows():
        inter = str(row.get("InterRotation", "")).strip()
        if not inter:
            continue
        branches = [b.strip() for b in inter.split(",")]
        for b in branches:
            m = re.match(r"(.+?)\s*\(([\d,]+)\s*CS\)", b)
            if m:
                branch_name = m.group(1).strip()
                cases = int(m.group(2).replace(",", ""))
                records.append(
                    {
                        "SKU": row["SKU"],
                        "Brand": row["Brand"],
                        "Branch": branch_name,
                        "Cases": cases,
                    }
                )
    return pd.DataFrame(records)


def normalize_branch(x):
    return str(x).strip().replace("–", "-").upper()

def load_summary(path=FILE_PATH, sheet=SUMMARY_SHEET):
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df = df.fillna("")

    col_map = {
        "Plant": "Plant",
        "AI_SKU": "SKU",
        "AI_MFGBRND": "Brand",
        "AI_BUSINESSUNIT": "BusinessUnit",
        "AI_BRANCH": "Branch",
        "Class": "Class",
        "<10_Days": "DaysLess10",
        "OOS": "OOS",
        "PlantInv %": "PlantInvPerc",
        "Upcoming_Plan": "UpcomingPlan",
        "Depletion_Date": "DepletionDate",
        "Risk": "Risk",
        "BackOrder?": "BackOrder",
        "Oversell": "Oversell",
        "MTD Sales": "MTD_Sales",
        "Forecast": "Forecast",
        "Avg_3MNTH_sales": "Avg_3MNTH_sales",
        "Inter_Rotation_Branches": "InterRotation",
        "Balance_Supply": "BalanceSupply",
        "Total_SKU_Plant_Inventory": "Total_SKU_Plant_Inventory",
    }
    df.rename(columns={c: col_map[c] for c in df.columns if c in col_map}, inplace=True)

    # numeric conversions
    df["PlantInvPerc"] = (
        pd.to_numeric(df.get("PlantInvPerc", "0").str.replace("%", ""), errors="coerce")
        / 100.0
    )
    df["MTD_Sales"] = pd.to_numeric(df.get("MTD_Sales", "0"), errors="coerce")
    df["Forecast"] = pd.to_numeric(df.get("Forecast", "0"), errors="coerce")
    df["Avg_3MNTH_sales"] = pd.to_numeric(
        df.get("Avg_3MNTH_sales", "0"), errors="coerce"
    )
    df["BalanceSupply"] = pd.to_numeric(df.get("BalanceSupply", "0"), errors="coerce")
    df["Total_SKU_Plant_Inventory"] = pd.to_numeric(
        df.get("Total_SKU_Plant_Inventory", "0"), errors="coerce"
    )

    # flags
    df["IsOOS"] = df["OOS"].apply(
        lambda x: 1 if str(x).lower() in ["oos", "yes", "true"] else 0
    )
    df["IsRisk"] = df["Risk"].apply(
        lambda x: 1 if str(x).lower() in ["risk", "yes", "true"] else 0
    )
    df["IsLowCover"] = df["DaysLess10"].apply(
        lambda x: 1 if str(x).lower() in ["yes", "true"] else 0
    )
    df["IsBackorder"] = df["BackOrder"].apply(
        lambda x: 1 if str(x).lower() in ["yes", "true"] else 0
    )
    df["IsOversell"] = df["Oversell"].apply(
        lambda x: 1 if str(x).lower() in ["yes", "true"] else 0
    )

    df["DepletionDate_parsed"] = df["DepletionDate"].apply(parse_depletion_date)
    df["UpcomingPlan_parsed"] = pd.to_datetime(df["UpcomingPlan"], errors="coerce")
    df["DaysToDepletion"] = (
        df["DepletionDate_parsed"] - pd.to_datetime(datetime.today().date())
    ).dt.days

    return df


def load_criticality(path=FILE_PATH, sheet=CRIT_SHEET):
    return pd.read_excel(path, sheet_name=sheet)


def load_coordinates(path=FILE_PATH, sheet=COORD_SHEET):
    df = pd.read_excel(path, sheet_name=sheet)
    return df.dropna(subset=["Latitude", "Longitude"])


df_full = load_summary()
df_crit = load_criticality()
df_coord = load_coordinates()

# Normalise names for joins and filters
df_full["Branch"] = df_full["Branch"].apply(normalize_branch)
df_full["BusinessUnit"] = df_full["BusinessUnit"].str.upper()
df_full["Brand"] = df_full["Brand"].str.upper()

df_crit.rename(columns=lambda x: normalize_branch(x), inplace=True)
df_crit["AI_SKU"] = df_crit["AI_SKU"].astype(str)
df_crit["AI_MFGBRND"] = df_crit["AI_MFGBRND"].astype(str).str.upper()

BRANDS = sorted(df_full["Brand"].unique())
PLANTS = sorted(df_full["Plant"].unique())
BRANCHES = sorted(df_full["Branch"].unique())
CLASSES = sorted(df_full["Class"].unique())
BUS_UNITS = sorted(df_full["BusinessUnit"].unique())
AI_SKUS = sorted(df_full["SKU"].unique())
BRANCH_DC_FILTER = ["0", "1-4", "5-6", "7-10"]
OVERSELL_FILTER = ["Yes", "No"]
RISK_FILTER = ["Yes", "No"]  # Yes/No based on IsRisk


def counting_dropdown(id_, count_id, options, label_text):
    return html.Div(
        [
            html.Small(
                f"{label_text} (0 selected)",
                id=count_id,
                className="text-muted d-block mb-1 filter-label",
            ),
            dcc.Dropdown(
                id=id_,
                options=[{"label": o, "value": o} for o in options],
                multi=True,
                placeholder=f"Select {label_text}",
                className="fixed-multi-dropdown hide-chips",
            ),
        ],
        className="mb-2",
    )


def kpi_card(title, value, color="primary", trend=None):
    icon = ""
    if trend == "up":
        icon = "▲ "
    elif trend == "down":
        icon = "▼ "
    elif trend == "flat":
        icon = "■ "
    label = icon + title if icon else title

    return dbc.Col(
        dbc.Card(
            dbc.CardBody(
                [
                    html.Div(
                        label,
                        className="text-muted",
                        style={"fontSize": "0.9rem"},
                    ),
                    html.Div(
                        value,
                        className="fw-bold kpi-number",
                        style={
                            "fontSize": "1.4rem",
                            "whiteSpace": "nowrap",
                            "overflow": "hidden",
                            "textOverflow": "ellipsis",
                        },
                    ),
                ]
            ),
            className=f"kpi-card border-{color}",
        ),
        xs=6,
        sm=4,
        md=3,
        lg=2,
    )


def _count_label(values, base):
    if values and len(values) == 1:
        return f"{base} (1 selected)"
    elif values:
        return f"{base} ({len(values)} selected)"
    return f"{base} (0 selected)"

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
)
server = app.server

app.index_string = """
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>Inventory & Supply Dashboard</title>
        {%favicon%}
        {%css%}
        <style>
        body {
            background-color: #f5f6f8;
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text",
                         "Segoe UI", system-ui, sans-serif;
        }
        .filter-card {
            border-radius: 0.75rem;
            box-shadow: 0 0.125rem 0.35rem rgba(0,0,0,0.06);
        }
        .kpi-card {
            height: 100%;
            border-radius: 0.75rem;
            box-shadow: 0 0.125rem 0.35rem rgba(0,0,0,0.06);
            transition: transform 0.12s ease-out, box-shadow 0.12s ease-out;
        }
        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 0.35rem 0.8rem rgba(0,0,0,0.08);
        }
        .sticky-kpi-bar {
            position: sticky;
            top: 0;
            z-index: 900;
            background: linear-gradient(180deg, rgba(245,246,248,0.97), rgba(245,246,248,0.90));
            padding-top: 0.25rem;
            padding-bottom: 0.35rem;
        }
        .tab-card {
            border-radius: 0.75rem;
            box-shadow: 0 0.125rem 0.35rem rgba(0,0,0,0.06);
        }
        .view-toggle {
            font-size: 0.75rem;
        }

        /* Dropdown styles */
        .fixed-multi-dropdown { font-size: 11px; }
        .fixed-multi-dropdown .Select-control {
            min-height: 34px;
            max-height: 38px;
            overflow: hidden;
            position: relative;
        }
        .fixed-multi-dropdown .Select-menu-outer {
            max-height: 260px;
            z-index: 9999;
        }
        .fixed-multi-dropdown .Select-option {
            padding: 4px 8px;
            border-bottom: 1px solid #f1f3f4;
        }
        .fixed-multi-dropdown .Select-option.is-focused {
            background-color: #e9f5ff;
        }
        .fixed-multi-dropdown .Select-option.is-selected {
            background-color: #1565C0;
            color: #ffffff;
        }
        .fixed-multi-dropdown .Select-arrow-zone {
            padding-right: 8px;
            display: flex;
            justify-content: flex-end !important;
            width: 22px;
        }
        .fixed-multi-dropdown .Select-control .Select-clear-zone {
            position: absolute;
            right: 22px;
            top: 0;
            bottom: 0;
            display: flex;
            align-items: center;
        }
        .fixed-multi-dropdown.hide-chips .Select-value {
            display: none !important;
        }
        .fixed-multi-dropdown.show-chips .Select-value {
            display: inline-flex !important;
        }

        .global-search { font-size: 11px; height: 32px; }

        .dash-table-container .dash-spreadsheet-container table {
            border-collapse: collapse !important;
        }

        .dash-table-container .dash-spreadsheet-container th {
            font-weight: 600;
            font-size: 11px;
            background-color: #ECEFF1 !important;
            border-bottom: 1px solid #CFD8DC !important;
        }

        .dash-table-container .dash-spreadsheet-container td {
            padding: 4px 6px;
            font-size: 11px;
        }

        .dash-table-container .dash-spreadsheet-container tr:nth-child(odd) td {
            background-color: #FAFAFA;
        }

        .dash-table-container .dash-spreadsheet-container td:hover {
            background-color: #E3F2FD !important;
        }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
"""

app.layout = dbc.Container(
    [
        # Header row
        dbc.Row(
            [
                dbc.Col(
                    html.Div(
                        [
                            html.H3(
                                "Inventory & Supply Dashboard",
                                className="mt-2 mb-0",
                            ),
                            html.Div(
                                "Executive & Analyst Views",
                                className="text-muted small",
                            ),
                        ]
                    ),
                    md=6,
                ),
                dbc.Col(
                    dbc.Row(
                        [
                            dbc.Col(
                                dbc.RadioItems(
                                    id="view-mode",
                                    options=[
                                        {"label": " Executive", "value": "executive"},
                                        {"label": " Analyst", "value": "analyst"},
                                    ],
                                    value="executive",
                                    inline=True,
                                    className="view-toggle",
                                ),
                                width="auto",
                            ),
                            dbc.Col(
                                html.Div(
                                    id="last-refresh",
                                    className="text-end text-muted small mt-2",
                                ),
                                width=True,
                            ),
                            dbc.Col(
                                html.Div(
                                    html.Img(
                                        src="/assets/aujan_logo.png",
                                        style={
                                            "height": "90px",
                                            "width": "220px",
                                            "objectFit": "contain",
                                        },
                                    ),
                                    className="d-flex justify-content-end",
                                ),
                                width="auto",
                            ),
                        ],
                        className="align-items-center justify-content-end",
                    ),
                    md=6,
                ),
            ],
            className="mb-2",
        ),

        # Auto-refresh interval (2 minutes)
        dcc.Interval(id="refresh-interval", interval=120000, n_intervals=0),

        # Filters row
        dbc.Row(
            [
                dbc.Col(
                    dbc.Button(
                        "Hide / Show Filters",
                        id="filter-toggle",
                        outline=True,
                        color="primary",
                        size="sm",
                        className="mb-1",
                    ),
                    md=2,
                ),
                dbc.Col(
                    dcc.Input(
                        id="global-search",
                        placeholder="Global search: SKU / Brand / Plant / Branch / Class...",
                        type="text",
                        debounce=True,
                        className="form-control form-control-sm global-search",
                    ),
                    md=6,
                ),
            ],
            className="mb-1",
        ),

        # Filter panel
        dbc.Collapse(
            dbc.Card(
                dbc.CardBody(
                    [
                        dbc.Row(
                            [
                                dbc.Col(
                                    counting_dropdown(
                                        "brand-filter",
                                        "brand-count-label",
                                        BRANDS,
                                        "Brand",
                                    ),
                                    md=3,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "plant-filter",
                                        "plant-count-label",
                                        PLANTS,
                                        "Plant",
                                    ),
                                    md=2,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "branch-filter",
                                        "branch-count-label",
                                        BRANCHES,
                                        "Branch",
                                    ),
                                    md=3,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "class-filter",
                                        "class-count-label",
                                        CLASSES,
                                        "Class",
                                    ),
                                    md=2,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "busunit-filter",
                                        "bu-count-label",
                                        BUS_UNITS,
                                        "BU",
                                    ),
                                    md=2,
                                ),
                            ],
                            className="gy-2",
                        ),
                        dbc.Row(
                            [
                                dbc.Col(
                                    counting_dropdown(
                                        "ai-sku-filter",
                                        "sku-count-label",
                                        AI_SKUS,
                                        "SKU",
                                    ),
                                    md=4,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "branch-dc-filter",
                                        "dc-count-label",
                                        BRANCH_DC_FILTER,
                                        "DC Days",
                                    ),
                                    md=2,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "oversell-filter",
                                        "oversell-count-label",
                                        OVERSELL_FILTER,
                                        "Oversell",
                                    ),
                                    md=2,
                                ),
                                dbc.Col(
                                    counting_dropdown(
                                        "risk-filter",
                                        "risk-count-label",
                                        RISK_FILTER,
                                        "Risk",
                                    ),
                                    md=2,
                                ),
                            ],
                            className="gy-2 mt-1",
                        ),
                    ]
                ),
                className="mb-2 filter-card",
            ),
            id="filter-collapse",
            is_open=True,
        ),

        # KPI row
        html.Div(
            dbc.Row(id="kpi-row", className="g-2 mb-2"),
            className="sticky-kpi-bar",
        ),

        # Tabs
        dbc.Card(
            dbc.CardBody(
                dbc.Tabs(
                    id="tabs",
                    active_tab="tab-overview",
                    children=[
                        dbc.Tab(
                            label="Overview",
                            tab_id="tab-overview",
                            children=[
                                dbc.Row(
                                    [
                                        dbc.Col(
                                            dcc.Graph(
                                                id="plant-inv-chart",
                                                style={"height": "640px"},
                                            ),
                                            md=6,
                                        ),
                                        dbc.Col(
                                            dcc.Graph(
                                                id="oos-risk-bar",
                                                style={"height": "640px"},
                                            ),
                                            md=6,
                                        ),
                                    ]
                                )
                            ],
                        ),
                        dbc.Tab(
                            label="Branch OOS Treemap",
                            tab_id="tab-treemap",
                            children=[
                                dcc.Graph(
                                    id="branch-oos-treemap",
                                    style={"height": "620px"},
                                )
                            ],
                        ),
                        dbc.Tab(
                            label="Information",
                            tab_id="tab-info",
                            children=[
                                dbc.Row(
                                    [
                                        dbc.Col(
                                            dbc.Button(
                                                "📤 Download Information.xlsx",
                                                id="info-export-btn",
                                                color="secondary",
                                                size="sm",
                                                className="mb-2",
                                            ),
                                            width="auto",
                                        ),
                                        dbc.Col(width=True),
                                        dcc.Download(id="info-download"),
                                    ],
                                    className="mb-1",
                                ),
                                dash_table.DataTable(
                                    id="info-table",
                                    columns=[],
                                    data=[],
                                    page_size=15,
                                    export_format="xlsx",
                                    merge_duplicate_headers=True,
                                    style_as_list_view=True,
                                    sort_action="native",
                                    filter_action="native",
                                    row_selectable="multi",
                                    style_table={
                                        "overflowX": "auto",
                                        "overflowY": "auto",
                                        "minWidth": "100%",
                                        "maxHeight": "620px",
                                        "borderRadius": "0.75rem",
                                        "border": "1px solid #CFD8DC",
                                    },
                                    style_cell={
                                        "fontSize": 11,
                                        "padding": "6px 8px",
                                        "whiteSpace": "nowrap",
                                        "border": "none",
                                        "fontFamily": '-apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", system-ui, sans-serif',
                                    },
                                    style_header={
                                        "backgroundColor": "#ECEFF1",
                                        "fontWeight": "600",
                                        "borderBottom": "1px solid #CFD8DC",
                                        "textTransform": "uppercase",
                                        "fontSize": 10,
                                    },
                                    style_data_conditional=[
                                        {
                                            "if": {"row_index": "odd"},
                                            "backgroundColor": "#FAFAFA",
                                        },
                                        {
                                            "if": {"state": "selected"},
                                            "backgroundColor": "#E3F2FD",
                                            "border": "1px solid #90CAF9",
                                        },
                                        {
                                            "if": {
                                                "filter_query": '{OOS} contains "OOS" || {OOS} = "Yes"'
                                            },
                                            "backgroundColor": "#FFEBEE",
                                        },
                                        {
                                            "if": {
                                                "filter_query": '{Risk} contains "RISK" || {Risk} = "Yes"'
                                            },
                                            "backgroundColor": "#FFF8E1",
                                        },
                                    ],
                                ),
                            ],
                        ),
                        dbc.Tab(
                            label="Stock Criticality",
                            tab_id="tab-crit",
                            children=[
                                dbc.Row(
                                    [
                                        dbc.Col(
                                            dbc.Button(
                                                "📤 Download Criticality.xlsx",
                                                id="crit-export-btn",
                                                color="success",
                                                size="sm",
                                                className="mb-2",
                                            ),
                                            width="auto",
                                        ),
                                        dbc.Col(width=True),
                                        dcc.Download(id="crit-download"),
                                    ],
                                    className="mb-1",
                                ),
                                dash_table.DataTable(
                                    id="crit-table",
                                    columns=[],
                                    data=[],
                                    page_size=15,
                                    export_format="xlsx",
                                    style_as_list_view=True,
                                    sort_action="native",
                                    filter_action="native",
                                    style_table={
                                        "overflowX": "auto",
                                        "minWidth": "100%",
                                        "borderRadius": "0.75rem",
                                        "border": "1px solid #CFD8DC",
                                    },
                                    style_cell={
                                        "fontSize": 11,
                                        "padding": "5px 6px",
                                        "whiteSpace": "nowrap",
                                        "border": "none",
                                        "textAlign": "center",
                                        "fontFamily": '-apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", system-ui, sans-serif',
                                    },
                                    style_header={
                                        "backgroundColor": "#ECEFF1",
                                        "fontWeight": "600",
                                        "borderBottom": "1px solid #CFD8DC",
                                        "textTransform": "uppercase",
                                        "fontSize": 10,
                                    },
                                    style_data_conditional=[],
                                ),
                            ],
                        ),
                        dbc.Tab(
                            label="Inter Rotation",
                            tab_id="tab-inter",
                            children=[
                                dcc.Graph(
                                    id="inter-rotation-map",
                                    style={"height": "650px"},
                                )
                            ],
                        ),
                    ],
                )
            ),
            className="tab-card",
        ),
    ],
    fluid=True,
)

def apply_filters(
    base_df,
    brands,
    plants,
    branches,
    classes,
    busunits,
    ai_skus,
    branch_dc,
    oversell,
    risk,
    global_search,
    ignore=None,
):
    ignore = ignore or set()
    df = base_df.copy()

    if "brand" not in ignore and brands:
        df = df[df["Brand"].isin(brands)]
    if "plant" not in ignore and plants:
        df = df[df["Plant"].isin(plants)]
    if "branch" not in ignore and branches:
        df = df[df["Branch"].isin(branches)]
    if "class" not in ignore and classes:
        df = df[df["Class"].isin(classes)]
    if "bu" not in ignore and busunits:
        df = df[df["BusinessUnit"].isin(busunits)]
    if "sku" not in ignore and ai_skus:
        df = df[df["SKU"].isin(ai_skus)]
    if "oversell" not in ignore and oversell:
        flags = [1 if v == "Yes" else 0 for v in oversell]
        df = df[df["IsOversell"].isin(flags)]
    if "risk" not in ignore and risk:
        mask = pd.Series(False, index=df.index)
        if "Yes" in risk:
            mask |= df["IsRisk"] == 1
        if "No" in risk:
            mask |= df["IsRisk"] == 0
        df = df[mask]

    # branch DC band filtering
    if "dc" not in ignore and branch_dc:
        mask = pd.Series(False, index=df.index)
        for b in branch_dc:
            if b == "0":
                mask |= df["DaysToDepletion"] <= 0
            elif b == "1-4":
                mask |= df["DaysToDepletion"].between(1, 4)
            elif b == "5-6":
                mask |= df["DaysToDepletion"].between(5, 6)
            elif b == "7-10":
                mask |= df["DaysToDepletion"].between(7, 10)
        df = df[mask]

    # global search
    if "global" not in ignore and global_search and global_search.strip():
        gs = global_search.strip().lower()
        search_cols = ["SKU", "Brand", "BusinessUnit", "Branch", "Plant", "Class"]
        mask = pd.Series(False, index=df.index)
        for col in search_cols:
            mask |= df[col].astype(str).str.lower().str.contains(gs)
        df = df[mask]

    return df

@app.callback(
    Output("filter-collapse", "is_open"),
    Input("filter-toggle", "n_clicks"),
    State("filter-collapse", "is_open"),
)
def toggle_filters(n, is_open):
    if n:
        return not is_open
    return is_open

@app.callback(
    Output("brand-filter", "options"),
    Output("plant-filter", "options"),
    Output("branch-filter", "options"),
    Output("class-filter", "options"),
    Output("busunit-filter", "options"),
    Output("ai-sku-filter", "options"),
    Input("brand-filter", "value"),
    Input("plant-filter", "value"),
    Input("branch-filter", "value"),
    Input("class-filter", "value"),
    Input("busunit-filter", "value"),
    Input("ai-sku-filter", "value"),
    Input("branch-dc-filter", "value"),
    Input("oversell-filter", "value"),
    Input("risk-filter", "value"),
    Input("global-search", "value"),
)
def update_filter_options(
    brands,
    plants,
    branches,
    classes,
    busunits,
    ai_skus,
    branch_dc,
    oversell,
    risk,
    global_search,
):
    # Brand options
    df_brand = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
        ignore={"brand"},
    )
    brand_opts = [{"label": b, "value": b} for b in sorted(df_brand["Brand"].unique())]

    # Plant options
    df_plant = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
        ignore={"plant"},
    )
    plant_opts = [{"label": p, "value": p} for p in sorted(df_plant["Plant"].unique())]

    # Branch options
    df_branch = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
        ignore={"branch"},
    )
    branch_opts = [
        {"label": b, "value": b} for b in sorted(df_branch["Branch"].unique())
    ]

    # Class options
    df_class = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
        ignore={"class"},
    )
    class_opts = [
        {"label": c, "value": c} for c in sorted(df_class["Class"].unique())
    ]

    # BU options
    df_bu = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
        ignore={"bu"},
    )
    bu_opts = [
        {"label": b, "value": b} for b in sorted(df_bu["BusinessUnit"].unique())
    ]

    # SKU options
    df_sku = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
        ignore={"sku"},
    )
    sku_opts = [{"label": s, "value": s} for s in sorted(df_sku["SKU"].unique())]

    return brand_opts, plant_opts, branch_opts, class_opts, bu_opts, sku_opts


@app.callback(
    Output("brand-filter", "className"),
    Output("plant-filter", "className"),
    Output("branch-filter", "className"),
    Output("class-filter", "className"),
    Output("busunit-filter", "className"),
    Output("ai-sku-filter", "className"),
    Output("branch-dc-filter", "className"),
    Output("oversell-filter", "className"),
    Output("risk-filter", "className"),
    Input("brand-filter", "value"),
    Input("plant-filter", "value"),
    Input("branch-filter", "value"),
    Input("class-filter", "value"),
    Input("busunit-filter", "value"),
    Input("ai-sku-filter", "value"),
    Input("branch-dc-filter", "value"),
    Input("oversell-filter", "value"),
    Input("risk-filter", "value"),
)
def update_dropdown_classes(
    brand_v,
    plant_v,
    branch_v,
    class_v,
    bu_v,
    sku_v,
    dc_v,
    oversell_v,
    risk_v,
):
    def cls(v):
        if v and len(v) == 1:
            return "fixed-multi-dropdown show-chips"
        return "fixed-multi-dropdown hide-chips"

    return (
        cls(brand_v),
        cls(plant_v),
        cls(branch_v),
        cls(class_v),
        cls(bu_v),
        cls(sku_v),
        cls(dc_v),
        cls(oversell_v),
        cls(risk_v),
    )


@app.callback(
    Output("kpi-row", "children"),
    Output("plant-inv-chart", "figure"),
    Output("oos-risk-bar", "figure"),
    Output("branch-oos-treemap", "figure"),
    Output("info-table", "columns"),
    Output("info-table", "data"),
    Output("crit-table", "columns"),
    Output("crit-table", "data"),
    Output("crit-table", "style_data_conditional"),
    Output("inter-rotation-map", "figure"),
    Output("last-refresh", "children"),
    Input("view-mode", "value"),
    Input("brand-filter", "value"),
    Input("plant-filter", "value"),
    Input("branch-filter", "value"),
    Input("class-filter", "value"),
    Input("busunit-filter", "value"),
    Input("ai-sku-filter", "value"),
    Input("branch-dc-filter", "value"),
    Input("oversell-filter", "value"),
    Input("risk-filter", "value"),
    Input("global-search", "value"),
    Input("refresh-interval", "n_intervals"),
)
def update_dashboard(
    view_mode,
    brands,
    plants,
    branches,
    classes,
    busunits,
    ai_skus,
    branch_dc,
    oversell,
    risk,
    global_search,
    n_intervals,
):
    # ---- Filtered data
    dff = apply_filters(
        df_full,
        brands,
        plants,
        branches,
        classes,
        busunits,
        ai_skus,
        branch_dc,
        oversell,
        risk,
        global_search,
    )
    dff_filtered = dff.copy()

    dff_filtered["BalanceSupply"] = pd.to_numeric(
        dff_filtered.get("BalanceSupply", 0), errors="coerce"
    ).fillna(0)
    dff_filtered["Total_SKU_Plant_Inventory"] = pd.to_numeric(
        dff_filtered.get("Total_SKU_Plant_Inventory", 0), errors="coerce"
    ).fillna(0)

    if not dff_filtered.empty:
        oos_group = (
            dff_filtered.groupby(["SKU", "Plant"], as_index=False)
            .agg(
                BalanceSupply_total=("BalanceSupply", "sum"),
                TotalPlantInv=("Total_SKU_Plant_Inventory", "max"),
            )
        )
        oos_group["OOS_Cases"] = (
            oos_group["BalanceSupply_total"] - oos_group["TotalPlantInv"]
        ).clip(lower=0)

        total_oos_cases = int(oos_group["OOS_Cases"].sum())

        # bring Brand for treemap
        sku_plant_brand = (
            dff_filtered[["SKU", "Plant", "Brand"]].drop_duplicates().copy()
        )
        oos_treemap = oos_group.merge(
            sku_plant_brand, on=["SKU", "Plant"], how="left"
        )
        oos_treemap.rename(
            columns={"BalanceSupply_total": "BalanceSupply"}, inplace=True
        )
    else:
        oos_group = pd.DataFrame(columns=["SKU", "Plant", "BalanceSupply_total", "TotalPlantInv", "OOS_Cases"])
        oos_treemap = pd.DataFrame(columns=["SKU", "Plant", "Brand", "BalanceSupply", "TotalPlantInv", "OOS_Cases"])
        total_oos_cases = 0

    total_sku_10d = dff_filtered["SKU"].nunique()
    bal_supply_sum = int(dff_filtered["BalanceSupply"].sum())
    risk_count = int(dff_filtered["IsRisk"].sum())
    bo_count = int(dff_filtered["IsBackorder"].sum())
    low_cover = int(dff_filtered["IsLowCover"].sum())
    oversell_count = int(dff_filtered["IsOversell"].sum())

    # KPI trend directions (simple)
    oos_trend = "up" if total_oos_cases > 0 else "flat"
    risk_trend = "up" if risk_count > 0 else "flat"
    bo_trend = "up" if bo_count > 0 else "flat"
    oversell_trend = "up" if oversell_count > 0 else "flat"
    bal_trend = "up" if bal_supply_sum > 0 else "flat"

    if view_mode == "executive":
        cards = [
            kpi_card("SKUs 0–10 Days Cover", total_sku_10d, trend="flat"),
            kpi_card("Balance Supply – Cases", f"{bal_supply_sum:,}", trend=bal_trend),
            kpi_card(
                "Total OOS Cases",
                f"{total_oos_cases:,}",
                color="danger",
                trend=oos_trend,
            ),
            kpi_card(
                "Supply Risk SKUs", risk_count, color="warning", trend=risk_trend
            ),
            kpi_card("Backorder SKUs", bo_count, color="danger", trend=bo_trend),
            kpi_card("Oversell SKUs", oversell_count, color="info", trend=oversell_trend),
        ]
    else:
        cards = [
            kpi_card("Total SKUs 0–10 Days Cover", total_sku_10d, trend="flat"),
            kpi_card("Balance Supply – Cases", f"{bal_supply_sum:,}", trend=bal_trend),
            kpi_card(
                "Total OOS Cases",
                f"{total_oos_cases:,}",
                color="danger",
                trend=oos_trend,
            ),
            kpi_card("Total Supply Risk SKUs", risk_count, trend=risk_trend),
            kpi_card("Total Backorder SKUs", bo_count, trend=bo_trend),
            kpi_card("<10d Cover SKUs", low_cover, trend="flat"),
            kpi_card(
                "Oversell SKUs", oversell_count, color="info", trend=oversell_trend
            ),
        ]

    if view_mode == "executive":
        info_columns = [
            {"name": "Plant", "id": "Plant", "type": "text"},
            {"name": "SKU", "id": "SKU", "type": "text"},
            {"name": "Brand", "id": "Brand", "type": "text"},
            {"name": "BusinessUnit", "id": "BusinessUnit", "type": "text"},
            {"name": "Branch", "id": "Branch", "type": "text"},
            {"name": "Class", "id": "Class", "type": "text"},
            {"name": "OOS", "id": "OOS", "type": "text"},
            {"name": "Risk", "id": "Risk", "type": "text"},
            {
                "name": "Balance Supply",
                "id": "BalanceSupply",
                "type": "numeric",
                "format": Format(group=Group.yes, scheme=Scheme.fixed, precision=0),
            },
            {"name": "Upcoming Plan", "id": "UpcomingPlan", "type": "text"},
            {"name": "Depletion Date", "id": "DepletionDate", "type": "text"},
        ]
    else:
        info_columns = [
            {"name": "Plant", "id": "Plant", "type": "text"},
            {"name": "SKU", "id": "SKU", "type": "text"},
            {"name": "Brand", "id": "Brand", "type": "text"},
            {"name": "BusinessUnit", "id": "BusinessUnit", "type": "text"},
            {"name": "Branch", "id": "Branch", "type": "text"},
            {"name": "Class", "id": "Class", "type": "text"},
            {"name": "Days<10", "id": "DaysLess10", "type": "text"},
            {"name": "OOS", "id": "OOS", "type": "text"},
            {
                "name": "PlantInv %",
                "id": "PlantInvPerc",
                "type": "numeric",
                "format": FormatTemplate.percentage(1),
            },
            {"name": "Oversell", "id": "Oversell", "type": "text"},
            {"name": "Risk", "id": "Risk", "type": "text"},
            {
                "name": "MTD Sales",
                "id": "MTD_Sales",
                "type": "numeric",
                "format": Format(group=Group.yes, scheme=Scheme.fixed, precision=0),
            },
            {
                "name": "Forecast",
                "id": "Forecast",
                "type": "numeric",
                "format": Format(group=Group.yes, scheme=Scheme.fixed, precision=0),
            },
            {
                "name": "Avg 3M Sales",
                "id": "Avg_3MNTH_sales",
                "type": "numeric",
                "format": Format(group=Group.yes, scheme=Scheme.fixed, precision=0),
            },
            {
                "name": "Balance Supply",
                "id": "BalanceSupply",
                "type": "numeric",
                "format": Format(group=Group.yes, scheme=Scheme.fixed, precision=0),
            },
            {"name": "Upcoming Plan", "id": "UpcomingPlan", "type": "text"},
            {"name": "Depletion Date", "id": "DepletionDate", "type": "text"},
        ]

    # date columns for table
    dff_filtered["DepletionDate"] = dff_filtered["DepletionDate_parsed"].dt.date
    dff_filtered["UpcomingPlan"] = dff_filtered["UpcomingPlan_parsed"].dt.date
    info_data = dff_filtered[[c["id"] for c in info_columns]].to_dict("records")

    def apply_theme(fig):
        fig.update_layout(
            template="plotly",
            paper_bgcolor="#ffffff",
            plot_bgcolor="#ffffff",
            font_color="#263238",
            font=dict(family='-apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", system-ui, sans-serif'),
        )
        return fig

    if not oos_treemap.empty:
        dff_agg = oos_treemap.copy()
        dff_agg["BalanceSupply"] = dff_agg["BalanceSupply"].clip(lower=0)
        dff_agg["OOS_Cases"] = dff_agg["OOS_Cases"].fillna(0)

        material_scale = [
            [0.0, "#E8F5E9"],   # light green
            [0.00001, "#C8E6C9"],
            [0.25, "#FFF9C4"],  # soft yellow
            [0.5, "#FFE082"],   # amber
            [0.75, "#FFAB91"],  # orange
            [1.0, "#E53935"],   # red
        ]

        fig_inv = px.treemap(
            dff_agg,
            path=["Brand", "Plant", "SKU"],
            values="BalanceSupply",
            color="OOS_Cases",
            color_continuous_scale=material_scale,
            custom_data=["BalanceSupply", "OOS_Cases", "TotalPlantInv"],
        )
        fig_inv.update_traces(
            texttemplate="<b>%{label}</b><br>%{value:,.0f} cs",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Balance Supply: %{customdata[0]:,.0f} cs<br>"
                "OOS Gap: %{customdata[1]:,.0f} cs<br>"
                "Plant Inventory: %{customdata[2]:,.0f} cs"
                "<extra></extra>"
            ),
        )
        fig_inv.update_layout(
            title="Plant Inventory Overview",
            margin=dict(t=40, l=10, r=10, b=10),
            coloraxis_colorbar={"title": "OOS Gap (cs)"},
        )
    else:
        fig_inv = go.Figure()
        fig_inv.update_layout(
            title="Plant Inventory Overview",
            margin=dict(t=40, l=10, r=10, b=10),
        )
    fig_inv = apply_theme(fig_inv)

    if not dff_filtered.empty:
        if view_mode == "executive":
            group_col = "BusinessUnit"
            title_bar = "OOS & Risk Exposure (Cases) by Business Unit"
        else:
            group_col = "Brand"
            title_bar = "OOS & Risk Exposure (Cases) by Brand"

        # allocate OOS_Cases from SKU+Plant to groups, proportional to BalanceSupply
        if not oos_group.empty:
            oos_join = dff_filtered.merge(
                oos_group[["SKU", "Plant", "BalanceSupply_total", "OOS_Cases"]],
                on=["SKU", "Plant"],
                how="left",
            )
            oos_join["BalanceSupply_total"] = oos_join["BalanceSupply_total"].replace(0, np.nan)
            oos_join["OOS_weight"] = (
                oos_join["BalanceSupply"] / oos_join["BalanceSupply_total"]
            )
            oos_join["OOS_weight"] = oos_join["OOS_weight"].fillna(0)
            oos_join["OOS_Cases_alloc"] = oos_join["OOS_weight"] * oos_join["OOS_Cases"]

            oos_by_group = (
                oos_join.groupby(group_col)["OOS_Cases_alloc"]
                .sum()
                .fillna(0)
                .rename("OOS_Cases")
            )
        else:
            oos_by_group = pd.Series(dtype=float, name="OOS_Cases")

        # risk cases = BalanceSupply on risk SKUs
        risk_by_group = (
            dff_filtered[dff_filtered["IsRisk"] == 1]
            .groupby(group_col)["BalanceSupply"]
            .sum()
            .fillna(0)
            .rename("Risk_Cases")
        )

        summary_df = (
            pd.concat([oos_by_group, risk_by_group], axis=1)
            .fillna(0)
            .reset_index()
        )
    else:
        summary_df = pd.DataFrame(
            columns=["Group", "OOS_Cases", "Risk_Cases"]
        )
        group_col = "Group"
        title_bar = "OOS & Risk Exposure"

    fig_bar = go.Figure()
    if not summary_df.empty:
        fig_bar.add_bar(
            name="OOS Gap (cs)",
            x=summary_df[group_col],
            y=summary_df["OOS_Cases"],
            marker_color="#E53935",
            text=[f"{v:,.0f}" for v in summary_df["OOS_Cases"]],
            textposition="outside",
            hovertemplate=(
                "<b>%{x}</b><br>"
                "OOS Gap: %{y:,.0f} cs"
                "<extra></extra>"
            ),
        )
        fig_bar.add_bar(
            name="Risk Exposure (cs)",
            x=summary_df[group_col],
            y=summary_df["Risk_Cases"],
            marker_color="#FB8C00",
            text=[f"{v:,.0f}" for v in summary_df["Risk_Cases"]],
            textposition="outside",
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Risk Exposure: %{y:,.0f} cs"
                "<extra></extra>"
            ),
        )
    fig_bar.update_layout(
        barmode="group",
        title=title_bar,
        margin=dict(t=60, b=130, l=10, r=10),
        xaxis_tickangle=-30,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
    )
    fig_bar = apply_theme(fig_bar)

    treemap_df = (
        dff_filtered.groupby(["Branch", "Brand", "SKU"])["BalanceSupply"]
        .sum()
        .reset_index(name="BalanceSupply")
    )
    if not treemap_df.empty:
        fig_treemap = px.treemap(
            treemap_df,
            path=["Branch", "Brand", "SKU"],
            values="BalanceSupply",
            color="Branch",
            hover_data=["SKU", "BalanceSupply"],
        )
        fig_treemap.update_traces(
            hovertemplate=(
                "<b>%{label}</b><br>"
                "SKU: %{customdata[0]}<br>"
                "Balance Supply: %{customdata[1]:,} cs"
                "<extra></extra>"
            )
        )
        fig_treemap.update_layout(margin=dict(t=40, l=10, r=10, b=10))
    else:
        fig_treemap = go.Figure()
        fig_treemap.update_layout(margin=dict(t=40, l=10, r=10, b=10))
    fig_treemap = apply_theme(fig_treemap)

    crit_columns = []
    crit_data = []
    style_data_conditional = []

    try:
        if not dff_filtered.empty and not df_crit.empty:
            crit_filtered = df_crit.copy()
            base_cols = ["AI_SKU", "AI_MFGBRND"]

            valid_skus = set(dff_filtered["SKU"].astype(str).unique())
            valid_brands = set(dff_filtered["Brand"].astype(str).unique())
            branch_set = set(dff_filtered["Branch"].unique())

            # filter rows (SKUs / Brands)
            if valid_skus:
                crit_filtered = crit_filtered[
                    crit_filtered["AI_SKU"].astype(str).isin(valid_skus)
                ]
            if valid_brands:
                crit_filtered = crit_filtered[
                    crit_filtered["AI_MFGBRND"].astype(str).str.upper().isin(valid_brands)
                ]

            # branch columns = intersection of crit sheet columns and filtered branches
            branch_set = set(dff_filtered["Branch"].unique())
            branch_cols = [
                c for c in crit_filtered.columns if c not in base_cols and c in branch_set
            ]

            if branch_cols:
                for col in branch_cols:
                    crit_filtered[col] = (
                        pd.to_numeric(crit_filtered[col], errors="coerce")
                        .round(0)
                        .astype("Int64")
                    )

                final_cols = base_cols + branch_cols
                crit_filtered = crit_filtered[final_cols]
                crit_filtered = crit_filtered.replace({np.nan: None})

                crit_columns = [{"name": c, "id": c} for c in final_cols]
                crit_data = crit_filtered.to_dict("records")

                # bands: 0 = deep red, 1–2 dark red, 3–4 bright red, 5–6 light red, 7–8 pale green, 9–10 green, >10 dark green
                bands = [
                    (0, 0, "#B71C1C", "white"),
                    (1, 2, "#D32F2F", "white"),
                    (3, 4, "#F44336", "white"),
                    (5, 6, "#FFCDD2", "#212121"),
                    (7, 8, "#FFFDE7", "#212121"),
                    (9, 10, "#C8E6C9", "#212121"),
                    (11, 9999, "#2E7D32", "white"),
                ]
                for col in branch_cols:
                    for low, high, bg, font in bands:
                        style_data_conditional.append(
                            {
                                "if": {
                                    "column_id": col,
                                    "filter_query": f"{{{col}}} >= {low} && {{{col}}} <= {high}",
                                },
                                "backgroundColor": bg,
                                "color": font,
                            }
                        )
                    style_data_conditional.append(
                        {
                            "if": {
                                "column_id": col,
                                "filter_query": f"{{{col}}} is blank",
                            },
                            "backgroundColor": "#ECEFF1",
                            "color": "#90A4AE",
                        }
                    )
    except Exception as e:
        print("CRIT TABLE ERROR:", e)
        crit_columns = []
        crit_data = []
        style_data_conditional = []

    if not dff_filtered.empty:
        unique_skus = dff_filtered["SKU"].nunique()
    else:
        unique_skus = 0

    if unique_skus == 1:
        sku_sel = dff_filtered["SKU"].unique()[0]

        # donor branches from InterRotation
        df_inter = explode_interrotation(dff_filtered[dff_filtered["SKU"] == sku_sel])

        if not df_inter.empty:
            df_inter = df_inter.merge(
                df_coord, left_on="Branch", right_on="AI_BRANCH", how="left"
            ).dropna(subset=["Latitude", "Longitude"])

            # OOS branches for this SKU
            oos_nodes = (
                dff_filtered[
                    (dff_filtered["SKU"] == sku_sel) & (dff_filtered["IsOOS"] == 1)
                ]
                .merge(df_coord, left_on="Branch", right_on="AI_BRANCH", how="left")
                .dropna(subset=["Latitude", "Longitude"])
            )

            # build nodes
            donors = df_inter[["Branch", "Latitude", "Longitude"]].drop_duplicates()
            donors["NodeType"] = "Donor"

            targets = oos_nodes[["Branch", "Latitude", "Longitude"]].drop_duplicates()
            targets["NodeType"] = "OOS"

            nodes = (
                pd.concat([donors, targets])
                .drop_duplicates(subset=["Branch", "Latitude", "Longitude", "NodeType"])
                .reset_index(drop=True)
            )

            # scatter nodes
            fig_inter = go.Figure()

            for node_type, color in [
                ("Donor", "#1E88E5"),
                ("OOS", "#E53935"),
            ]:
                subset = nodes[nodes["NodeType"] == node_type]
                if subset.empty:
                    continue
                fig_inter.add_trace(
                    go.Scattermapbox(
                        lat=subset["Latitude"],
                        lon=subset["Longitude"],
                        mode="markers",
                        marker=dict(size=12, color=color),
                        name=node_type,
                        text=subset["Branch"],
                        hovertemplate=(
                            "<b>%{text}</b><br>"
                            f"Type: {node_type}<br>"
                            f"SKU: {sku_sel}"
                            "<extra></extra>"
                        ),
                    )
                )

            # edges from each donor to each OOS
            if not donors.empty and not targets.empty:
                edge_lats = []
                edge_lons = []
                for _, drow in donors.iterrows():
                    for _, trow in targets.iterrows():
                        edge_lats += [drow["Latitude"], trow["Latitude"], None]
                        edge_lons += [drow["Longitude"], trow["Longitude"], None]

                fig_inter.add_trace(
                    go.Scattermapbox(
                        lat=edge_lats,
                        lon=edge_lons,
                        mode="lines",
                        line=dict(width=2, color="#90A4AE"),
                        hoverinfo="skip",
                        showlegend=False,
                    )
                )

            fig_inter.update_layout(
                mapbox_style="carto-positron",
                mapbox_zoom=4,
                mapbox_center={
                    "lat": nodes["Latitude"].mean(),
                    "lon": nodes["Longitude"].mean(),
                },
                margin=dict(t=40, l=10, r=10, b=10),
                title=f"Inter-Rotation Network for SKU {sku_sel}",
            )
        else:
            fig_inter = go.Figure()
            fig_inter.update_layout(
                mapbox_style="carto-positron",
                mapbox_zoom=3,
                margin=dict(t=40, l=10, r=10, b=10),
                title="Inter-Rotation Network",
            )
    else:
        # default bubble map view
        df_inter = explode_interrotation(dff_filtered)
        if not df_inter.empty:
            df_map = df_inter.merge(
                df_coord, left_on="Branch", right_on="AI_BRANCH", how="left"
            ).dropna(subset=["Latitude", "Longitude"])

            df_map["Label"] = df_map.apply(
                lambda x: f"{x['Branch']} ({x['Cases']} cs) – {x['Brand']}", axis=1
            )
            fig_inter = px.scatter_mapbox(
                df_map,
                lat="Latitude",
                lon="Longitude",
                size="Cases",
                color="Brand",
                text="Branch",
                hover_name="Label",
                hover_data={"Cases": True, "Latitude": False, "Longitude": False},
                zoom=4,
                mapbox_style="carto-positron",
                title="Inter-Rotation Map",
                size_max=28,
            )
            fig_inter.update_traces(
                hovertemplate=(
                    "<b>%{hovertext}</b><br>"
                    "Cases: %{marker.size:,.0f} cs"
                    "<extra></extra>"
                )
            )
            fig_inter.update_layout(margin=dict(t=40, l=10, r=10, b=10))
        else:
            fig_inter = go.Figure()
            fig_inter.update_layout(
                mapbox_style="carto-positron",
                mapbox_zoom=3,
                margin=dict(t=40, l=10, r=10, b=10),
                title="Inter-Rotation Map",
            )

    fig_inter = apply_theme(fig_inter)

    last_refresh = "Last refresh: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return (
        cards,
        fig_inv,
        fig_bar,
        fig_treemap,
        info_columns,
        info_data,
        crit_columns,
        crit_data,
        style_data_conditional,
        fig_inter,
        last_refresh,
    )

@app.callback(
    Output("crit-download", "data"),
    Input("crit-export-btn", "n_clicks"),
    State("crit-table", "data"),
    prevent_initial_call=True,
)
def export_criticality(n_clicks, rows):
    if not n_clicks or not rows:
        return dash.no_update

    df = pd.DataFrame(rows)

    def to_xlsx(bytes_io):
        with pd.ExcelWriter(bytes_io, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Criticality")

    return send_bytes(to_xlsx, "Stock_Criticality.xlsx")


@app.callback(
    Output("info-download", "data"),
    Input("info-export-btn", "n_clicks"),
    State("info-table", "data"),
    prevent_initial_call=True,
)
def export_information(n_clicks, rows):
    if not n_clicks or not rows:
        return dash.no_update

    df = pd.DataFrame(rows)

    def to_xlsx(bytes_io):
        with pd.ExcelWriter(bytes_io, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Information")

    return send_bytes(to_xlsx, "Information.xlsx")


@app.callback(Output("brand-count-label", "children"), Input("brand-filter", "value"))
def _brand_count(v):
    return _count_label(v, "Brand")


@app.callback(Output("plant-count-label", "children"), Input("plant-filter", "value"))
def _plant_count(v):
    return _count_label(v, "Plant")


@app.callback(Output("branch-count-label", "children"), Input("branch-filter", "value"))
def _branch_count(v):
    return _count_label(v, "Branch")


@app.callback(Output("class-count-label", "children"), Input("class-filter", "value"))
def _class_count(v):
    return _count_label(v, "Class")


@app.callback(Output("bu-count-label", "children"), Input("busunit-filter", "value"))
def _bu_count(v):
    return _count_label(v, "BU")


@app.callback(Output("sku-count-label", "children"), Input("ai-sku-filter", "value"))
def _sku_count(v):
    return _count_label(v, "SKU")


@app.callback(Output("dc-count-label", "children"), Input("branch-dc-filter", "value"))
def _dc_count(v):
    return _count_label(v, "DC Days")


@app.callback(
    Output("oversell-count-label", "children"), Input("oversell-filter", "value")
)
def _oversell_count(v):
    return _count_label(v, "Oversell")


@app.callback(Output("risk-count-label", "children"), Input("risk-filter", "value"))
def _risk_count(v):
    return _count_label(v, "Risk")


if __name__ == "__main__":
    app.run(port=SERVER_PORT, debug=True, host="0.0.0.0")
