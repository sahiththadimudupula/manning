import math
import re
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Spinning Manpower Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PRIMARY_SPINNING_PATH = Path(r"Spinning.xlsx")
FALLBACK_SPINNING_PATHS = [
    PRIMARY_SPINNING_PATH,
    Path(__file__).resolve().parent / "Spinning.xlsx",
    Path("Spinning.xlsx"),
    Path("/mnt/data/Spinning.xlsx"),
]

DISPLAY_COLUMNS = [
    "Location",
    "Business",
    "Section",
    "Sr_No",
    "Department",
    "Designation",
    "Machine_Count",
    "Workload",
    "Formulas",
    "BE_Scientific_Manpower",
    "BE_Final_Manpower",
    "General_Shift",
    "Shift_A",
    "Shift_B",
    "Shift_C",
    "Reliever",
    "Remarks",
]

EDITABLE_COLUMNS = [
    "Formulas",
    "BE_Final_Manpower",
    "General_Shift",
    "Shift_A",
    "Shift_B",
    "Shift_C",
    "Reliever",
    "Remarks",
]

NUMERIC_COLUMNS = [
    "Sr_No",
    "Machine_Count",
    "BE_Scientific_Manpower",
    "BE_Final_Manpower",
    "General_Shift",
    "Shift_A",
    "Shift_B",
    "Shift_C",
    "Reliever",
]

TEXT_COLUMNS = [
    "Location",
    "Business",
    "Section",
    "Department",
    "Designation",
    "Workload",
    "Formulas",
    "Remarks",
]

st.markdown(
    """
    <style>
        .stApp {
            background: linear-gradient(180deg, #edf3fb 0%, #f7faff 42%, #ffffff 100%);
        }

        .block-container {
            max-width: 1540px;
            padding-top: 1rem;
            padding-bottom: 1.25rem;
        }

        label[data-testid="stWidgetLabel"] p {
            color: #0f172a !important;
            font-weight: 700 !important;
            font-size: 0.92rem !important;
        }

        div[data-testid="stTabs"] {
            margin-top: 0.2rem;
        }

        div[data-testid="stTabs"] button {
            font-weight: 700;
            font-size: 0.98rem;
            padding: 0.8rem 1rem;
            color: #334155 !important;
            border-radius: 12px 12px 0 0;
        }

        div[data-testid="stTabs"] button[aria-selected="true"] {
            color: #0f4c81 !important;
        }

        div[data-testid="stTabs"] button p {
            color: inherit !important;
        }

        div[data-baseweb="select"] > div {
            min-height: 48px !important;
            background: #ffffff !important;
            border: 1px solid #b8c5d6 !important;
            border-radius: 14px !important;
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.04);
        }

        div[data-baseweb="select"] > div * {
            color: #0f172a !important;
            opacity: 1 !important;
        }

        div[data-baseweb="select"] input {
            color: #0f172a !important;
            -webkit-text-fill-color: #0f172a !important;
        }

        div[data-baseweb="select"] svg {
            color: #334155 !important;
            fill: #334155 !important;
        }

        ul[role="listbox"] {
            background: #ffffff !important;
        }

        ul[role="listbox"] li {
            color: #0f172a !important;
            background: #ffffff !important;
        }

        ul[role="listbox"] li[aria-selected="true"] {
            background: #dbeafe !important;
            color: #0f172a !important;
        }

        div[data-baseweb="tag"] {
            background: linear-gradient(135deg, #0f4c81 0%, #1d5f97 100%) !important;
            border: none !important;
            border-radius: 999px !important;
            padding: 3px 8px !important;
            box-shadow: 0 4px 10px rgba(15, 76, 129, 0.18);
        }

        div[data-baseweb="tag"] span,
        div[data-baseweb="tag"] svg {
            color: #ffffff !important;
            fill: #ffffff !important;
            font-weight: 700 !important;
        }

        div[data-baseweb="popover"] {
            border-radius: 14px !important;
            overflow: hidden !important;
        }

        .stButton > button,
        .stDownloadButton > button {
            width: 100%;
            min-height: 48px;
            border-radius: 14px;
            border: none;
            background: linear-gradient(135deg, #0f4c81 0%, #1d5f97 100%);
            color: #ffffff !important;
            font-weight: 700;
            font-size: 0.96rem;
            box-shadow: 0 10px 22px rgba(15, 76, 129, 0.18);
            transition: all 0.18s ease;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 14px 24px rgba(15, 76, 129, 0.24);
            color: #ffffff !important;
        }

        .stButton > button:focus,
        .stDownloadButton > button:focus {
            color: #ffffff !important;
            outline: none !important;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.22);
        }

        div[data-testid="stExpander"] details {
            border-radius: 16px;
            border: 1px solid rgba(15, 76, 129, 0.12);
            background: linear-gradient(135deg, #0f4c81 0%, #1d5f97 100%);
            overflow: hidden;
        }

        div[data-testid="stExpander"] details summary p,
        div[data-testid="stExpander"] details summary span {
            color: #ffffff !important;
            font-weight: 700;
        }

        div[data-testid="stExpanderDetails"] {
            background: #ffffff;
            border-radius: 0 0 16px 16px;
            padding: 0.85rem 0.3rem 0.4rem 0.3rem;
        }

        div[data-testid="stExpanderDetails"] * {
            color: #111827 !important;
        }

        div[data-testid="stDataEditor"] {
            border-radius: 16px;
            overflow: hidden;
            border: 1px solid #dbe4f0;
        }

        .hero-wrap {
            background: linear-gradient(135deg, #0f4c81 0%, #1d5f97 46%, #4f90c8 100%);
            color: white;
            padding: 1.5rem 1.6rem;
            border-radius: 24px;
            box-shadow: 0 16px 34px rgba(15, 76, 129, 0.18);
            margin-bottom: 1rem;
        }

        .hero-title {
            font-size: 2.05rem;
            font-weight: 800;
            margin-bottom: 0.18rem;
            letter-spacing: 0.2px;
        }

        .hero-subtitle {
            font-size: 0.98rem;
            opacity: 0.98;
        }

        .metric-card {
            background: white;
            border-radius: 18px;
            padding: 1rem 1.1rem;
            box-shadow: 0 10px 28px rgba(16, 24, 40, 0.07);
            border: 1px solid rgba(15, 76, 129, 0.08);
            min-height: 122px;
        }

        .metric-label {
            color: #475467;
            font-size: 0.88rem;
            font-weight: 700;
            margin-bottom: 0.45rem;
        }

        .metric-value {
            color: #0f172a;
            font-size: 2rem;
            font-weight: 800;
            line-height: 1.05;
            margin-bottom: 0.3rem;
        }

        .metric-note {
            color: #64748b;
            font-size: 0.84rem;
        }

        .panel-card {
            background: white;
            border-radius: 22px;
            padding: 1rem 1.15rem 1.15rem 1.15rem;
            box-shadow: 0 10px 28px rgba(16, 24, 40, 0.07);
            border: 1px solid rgba(15, 76, 129, 0.08);
            margin-bottom: 1rem;
        }

        .section-title {
            font-size: 1.12rem;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 0.1rem;
        }

        .section-subtitle {
            color: #475467;
            font-size: 0.9rem;
            margin-bottom: 0.9rem;
        }

        .small-note {
            color: #475467;
            font-size: 0.83rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_spinning_master() -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    for candidate in FALLBACK_SPINNING_PATHS:
        if candidate.exists():
            source_df = pd.read_excel(candidate, sheet_name="Spinning")
            source_df["Excel_Row_No"] = range(2, len(source_df) + 2)
            return source_df, str(candidate)
    return None, None


def safe_float(value, default=0.0):
    try:
        if value is None or value == "":
            return default
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def round_2(value):
    return round(safe_float(value), 2)


def excel_round(value, digits=0):
    return round(value, digits)


def excel_roundup(value, digits=0):
    factor = 10 ** digits
    return math.ceil(value * factor) / factor


def get_initial_tfo_data():
    data = [
        {"Count": "2/24 KW", "Customer": "Vapi", "Count2": 12.00, "Speed": 9000.00, "TPI": 8.50, "Utilization": 0.98, "Efficiency": 0.88, "Production Required / day Kgs": 9386.50, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/24 KW", "Customer": "Vapi", "Count2": 12.00, "Speed": 6000.00, "TPI": 8.50, "Utilization": 0.97, "Efficiency": 0.85, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 0.00, "Eff": 0.00, "Machine divisor": 72.00},
        {"Count": "2/30 CW bci Anjar", "Customer": "Anjar", "Count2": 15.00, "Speed": 10000.00, "TPI": 11.00, "Utilization": 0.98, "Efficiency": 0.90, "Production Required / day Kgs": 1408.00, "TFO divisor": 240.00, "mpm": 950.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/20 KW", "Customer": "Vapi", "Count2": 10.00, "Speed": 8500.00, "TPI": 8.50, "Utilization": 0.98, "Efficiency": 0.87, "Production Required / day Kgs": 11968.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/20 KW BCI Anjar", "Customer": "Anjar", "Count2": 10.00, "Speed": 7000.00, "TPI": 8.50, "Utilization": 0.98, "Efficiency": 0.87, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/12 K bci Anjar", "Customer": "Anjar", "Count2": 6.00, "Speed": 6000.00, "TPI": 6.50, "Utilization": 0.98, "Efficiency": 0.80, "Production Required / day Kgs": 1315.14, "TFO divisor": 240.00, "mpm": 800.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/9 KC", "Customer": "Vapi", "Count2": 4.50, "Speed": 4000.00, "TPI": 6.50, "Utilization": 0.98, "Efficiency": 0.75, "Production Required / day Kgs": 3301.00, "TFO divisor": 240.00, "mpm": 700.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/6 kc", "Customer": "Rugs", "Count2": 3.00, "Speed": 4000.00, "TPI": 6.50, "Utilization": 0.98, "Efficiency": 0.80, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/6 kc", "Customer": "Rugs", "Count2": 3.00, "Speed": 3500.00, "TPI": 6.50, "Utilization": 0.98, "Efficiency": 0.80, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "2/19 C hygro", "Customer": "Vapi", "Count2": 9.50, "Speed": 6000.00, "TPI": 6.50, "Utilization": 0.98, "Efficiency": 0.85, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "9 com +80 PVA", "Customer": "Vapi", "Count2": 8.09, "Speed": 7000.00, "TPI": 13.00, "Utilization": 0.98, "Efficiency": 0.85, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 700.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "DYED/ Others OS", "Customer": "Vapi", "Count2": 13.33, "Speed": 8000.00, "TPI": 15.00, "Utilization": 0.98, "Efficiency": 0.85, "Production Required / day Kgs": 2323.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "16+106 80 PVA OS", "Customer": "Vapi", "Count2": 13.33, "Speed": 8000.00, "TPI": 19.00, "Utilization": 0.98, "Efficiency": 0.90, "Production Required / day Kgs": 4287.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "13 hygro+80 PVA", "Customer": "Vapi", "Count2": 11.18, "Speed": 8500.00, "TPI": 13.00, "Utilization": 0.98, "Efficiency": 0.90, "Production Required / day Kgs": 3051.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "13 CC +80 PVA", "Customer": "Vapi", "Count2": 11.18, "Speed": 8500.00, "TPI": 13.00, "Utilization": 0.98, "Efficiency": 0.90, "Production Required / day Kgs": 4358.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "14 CC EGY +80 PVA", "Customer": "Vapi", "Count2": 11.18, "Speed": 7000.00, "TPI": 13.00, "Utilization": 0.98, "Efficiency": 0.90, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "13 CC +80 PVA OS", "Customer": "Vapi", "Count2": 11.18, "Speed": 7000.00, "TPI": 13.00, "Utilization": 0.98, "Efficiency": 0.90, "Production Required / day Kgs": 0.00, "TFO divisor": 240.00, "mpm": 850.00, "Eff": 0.85, "Machine divisor": 72.00},
        {"Count": "4/2/6 K", "Customer": "Rugs", "Count2": 0.75, "Speed": 2000.00, "TPI": 2.80, "Utilization": 0.98, "Efficiency": 0.70, "Production Required / day Kgs": 2713.78, "TFO divisor": 80.00, "mpm": 400.00, "Eff": 0.70, "Machine divisor": 72.00},
    ]
    return pd.DataFrame(data)


if "tfo_input_df" not in st.session_state:
    st.session_state.tfo_input_df = get_initial_tfo_data()

if "full_spinning_editor_version" not in st.session_state:
    st.session_state.full_spinning_editor_version = 0


def calculate_upper_tfo_metrics(df: pd.DataFrame) -> pd.DataFrame:
    output_rows = []

    for _, row in df.iterrows():
        count2 = safe_float(row.get("Count2"))
        speed = safe_float(row.get("Speed"))
        tpi = safe_float(row.get("TPI"))
        efficiency = safe_float(row.get("Efficiency"))
        production_required_day = safe_float(row.get("Production Required / day Kgs"))
        tfo_divisor = safe_float(row.get("TFO divisor"), 240)
        mpm = safe_float(row.get("mpm"))
        eff_value = safe_float(row.get("Eff"))
        machine_divisor = safe_float(row.get("Machine divisor"), 72)

        production_required_month = production_required_day * 30 if production_required_day > 0 else 0.0

        if count2 > 0 and tpi > 0:
            production_per_drum_day = (
                (speed * 60 * 8 * efficiency * 2)
                / (tpi * 36 * 840 * count2 * 2.202)
            ) * 3
        else:
            production_per_drum_day = 0.0

        if production_per_drum_day > 0:
            no_of_drums_required = production_required_day / production_per_drum_day
        else:
            no_of_drums_required = 0.0

        if tfo_divisor > 0:
            no_of_tfo_required_shift = no_of_drums_required / tfo_divisor
        else:
            no_of_tfo_required_shift = 0.0

        if count2 > 0:
            kgs_per_drum_day = (
                (mpm * eff_value * 8 * 60 * 1.09)
                / (count2 * 840 * 2.202)
            ) * 3
        else:
            kgs_per_drum_day = 0.0

        if kgs_per_drum_day > 0:
            no_of_drums = production_required_day / kgs_per_drum_day
        else:
            no_of_drums = 0.0

        if machine_divisor > 0:
            no_of_machines = no_of_drums / machine_divisor
        else:
            no_of_machines = 0.0

        new_row = row.to_dict()
        new_row["Production Required / Month Kgs"] = round_2(production_required_month)
        new_row["Production per Drum/day"] = round_2(production_per_drum_day)
        new_row["No. of Drums Required"] = round_2(no_of_drums_required)
        new_row["No. of TFO Required / shift"] = round_2(no_of_tfo_required_shift)
        new_row["kgs/drum/day"] = round_2(kgs_per_drum_day)
        new_row["No. of Drums"] = round_2(no_of_drums)
        new_row["no. of machines"] = round_2(no_of_machines)
        output_rows.append(new_row)

    result_df = pd.DataFrame(output_rows)

    numeric_cols = [
        "Count2", "Speed", "TPI", "Utilization", "Efficiency",
        "Production Required / day Kgs", "Production Required / Month Kgs",
        "TFO divisor", "mpm", "Eff", "Machine divisor",
        "Production per Drum/day", "No. of Drums Required",
        "No. of TFO Required / shift", "kgs/drum/day", "No. of Drums",
        "no. of machines",
    ]

    for col in numeric_cols:
        result_df[col] = pd.to_numeric(result_df[col], errors="coerce").fillna(0).round(2)

    result_df["Upper_Row_No"] = range(2, len(result_df) + 2)
    return result_df


def build_upper_total_row(df: pd.DataFrame) -> pd.DataFrame:
    total_row = {
        "Count": "TOTAL",
        "Customer": "",
        "Count2": round_2(df["Count2"].sum()),
        "Speed": 0.00,
        "TPI": 0.00,
        "Utilization": 0.00,
        "Efficiency": 0.00,
        "Production Required / day Kgs": round_2(df["Production Required / day Kgs"].sum()),
        "Production Required / Month Kgs": round_2(df["Production Required / Month Kgs"].sum()),
        "Production per Drum/day": round_2(df["Production per Drum/day"].sum()),
        "No. of Drums Required": round_2(df["No. of Drums Required"].sum()),
        "No. of TFO Required / shift": round_2(df["No. of TFO Required / shift"].sum()),
        "mpm": 0.00,
        "Eff": 0.00,
        "kgs/drum/day": round_2(df["kgs/drum/day"].sum()),
        "No. of Drums": round_2(df["No. of Drums"].sum()),
        "no. of machines": round_2(df["no. of machines"].sum()),
        "TFO divisor": 0.00,
        "Machine divisor": 0.00,
    }
    return pd.DataFrame([total_row])


def split_shift(total_manpower, mode="three_shift"):
    total_manpower = int(round(total_manpower))

    if mode == "general":
        return total_manpower, 0, 0, 0

    shift_a = math.ceil(total_manpower / 3)
    shift_b = math.ceil((total_manpower - shift_a) / 2)
    shift_c = total_manpower - shift_a - shift_b
    return 0, shift_a, shift_b, shift_c


def calculate_lower_tfo_manpower(upper_df: pd.DataFrame):
    sum_no_of_drums_total = round_2(upper_df["No. of Drums"].sum())
    sum_tfo_required_shift_total = round_2(upper_df["No. of TFO Required / shift"].sum())

    count_426k_mask = upper_df["Count"].astype(str).str.strip().str.upper() == "4/2/6 K"
    no_of_drums_426k = round_2(upper_df.loc[count_426k_mask, "No. of Drums"].sum())

    assembly_winding = excel_roundup((sum_no_of_drums_total / 36) * 3, 0)
    jumbo_assembly_winding = excel_round(excel_roundup(no_of_drums_426k, 0) / 16, 0) * 2
    tfo_operator = excel_round(sum_tfo_required_shift_total / 6, 0) * 3
    tfo_doffer = excel_roundup(sum_tfo_required_shift_total / 4, 0) * 3

    rows = [
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 1, "Department": "Assembly winding", "Designation": "operator", "Machine_Count": "", "Workload": "", "Formulas": "ROUNDUP((SUM(T2:T18)/36)*3,0)", "BE_Scientific_Manpower": round_2(assembly_winding), "BE_Final_Manpower": round_2(assembly_winding), **dict(zip(["General_Shift", "Shift_A", "Shift_B", "Shift_C"], split_shift(assembly_winding, "three_shift"))), "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 2, "Department": "Jumbo Assembly Winding", "Designation": "operator", "Machine_Count": "", "Workload": "", "Formulas": "ROUND(ROUNDUP(T19,0)/16,0)*2", "BE_Scientific_Manpower": round_2(jumbo_assembly_winding), "BE_Final_Manpower": round_2(jumbo_assembly_winding), **dict(zip(["General_Shift", "Shift_A", "Shift_B", "Shift_C"], split_shift(jumbo_assembly_winding, "three_shift"))), "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 3, "Department": "TFO", "Designation": "TFO Operator", "Machine_Count": "", "Workload": "", "Formulas": "ROUND(SUM(N2:N18)/6,0)*3", "BE_Scientific_Manpower": round_2(tfo_operator), "BE_Final_Manpower": round_2(tfo_operator), **dict(zip(["General_Shift", "Shift_A", "Shift_B", "Shift_C"], split_shift(tfo_operator, "three_shift"))), "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 4, "Department": "TFO", "Designation": "TFO Operator (Doffer)", "Machine_Count": "", "Workload": "", "Formulas": "ROUNDUP(SUM(N2:N18)/4,0)*3", "BE_Scientific_Manpower": round_2(tfo_doffer), "BE_Final_Manpower": round_2(tfo_doffer), **dict(zip(["General_Shift", "Shift_A", "Shift_B", "Shift_C"], split_shift(tfo_doffer, "three_shift"))), "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 5, "Department": "Jumbo TFO", "Designation": "TFO Operator", "Machine_Count": "", "Workload": "", "Formulas": "2*3", "BE_Scientific_Manpower": 6.00, "BE_Final_Manpower": 6.00, "General_Shift": 0, "Shift_A": 2, "Shift_B": 2, "Shift_C": 2, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 6, "Department": "Vijaylakshmi TFO", "Designation": "TFO Operator", "Machine_Count": "", "Workload": "", "Formulas": "2*3", "BE_Scientific_Manpower": 6.00, "BE_Final_Manpower": 6.00, "General_Shift": 0, "Shift_A": 2, "Shift_B": 2, "Shift_C": 2, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 7, "Department": "Jobber", "Designation": "Jobber", "Machine_Count": "", "Workload": "", "Formulas": "1*3", "BE_Scientific_Manpower": 3.00, "BE_Final_Manpower": 3.00, "General_Shift": 0, "Shift_A": 1, "Shift_B": 1, "Shift_C": 1, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 8, "Department": "Cone Carrier", "Designation": "Cone carrier", "Machine_Count": "", "Workload": "", "Formulas": "1*3", "BE_Scientific_Manpower": 3.00, "BE_Final_Manpower": 3.00, "General_Shift": 0, "Shift_A": 1, "Shift_B": 1, "Shift_C": 1, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 9, "Department": "Cone Checker", "Designation": "Cone checker", "Machine_Count": "", "Workload": "", "Formulas": "2*3", "BE_Scientific_Manpower": 6.00, "BE_Final_Manpower": 6.00, "General_Shift": 0, "Shift_A": 2, "Shift_B": 2, "Shift_C": 2, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 10, "Department": "Cone tipping", "Designation": "", "Machine_Count": "", "Workload": "", "Formulas": "1*3", "BE_Scientific_Manpower": 3.00, "BE_Final_Manpower": 3.00, "General_Shift": 0, "Shift_A": 1, "Shift_B": 1, "Shift_C": 1, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 11, "Department": "Fork lift opt", "Designation": "", "Machine_Count": "", "Workload": "", "Formulas": "1*1", "BE_Scientific_Manpower": 1.00, "BE_Final_Manpower": 1.00, "General_Shift": 1, "Shift_A": 0, "Shift_B": 0, "Shift_C": 0, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 12, "Department": "Packer", "Designation": "", "Machine_Count": "", "Workload": "", "Formulas": "4*3", "BE_Scientific_Manpower": 12.00, "BE_Final_Manpower": 12.00, "General_Shift": 0, "Shift_A": 4, "Shift_B": 4, "Shift_C": 4, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 13, "Department": "Packing Jobber", "Designation": "Jobber", "Machine_Count": "", "Workload": "", "Formulas": "0*3", "BE_Scientific_Manpower": 0.00, "BE_Final_Manpower": 0.00, "General_Shift": 0, "Shift_A": 0, "Shift_B": 0, "Shift_C": 0, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 15, "Department": "DEO", "Designation": "DEO", "Machine_Count": "", "Workload": "", "Formulas": "1*1", "BE_Scientific_Manpower": 1.00, "BE_Final_Manpower": 1.00, "General_Shift": 1, "Shift_A": 0, "Shift_B": 0, "Shift_C": 0, "Reliever": 0, "Remarks": ""},
        {"Location": "Vapi", "Business": "Spinning", "Section": "TFO", "Sr_No": 16, "Department": "House Keeper", "Designation": "Contractors", "Machine_Count": "", "Workload": "", "Formulas": "2*3", "BE_Scientific_Manpower": 6.00, "BE_Final_Manpower": 6.00, "General_Shift": 0, "Shift_A": 2, "Shift_B": 2, "Shift_C": 2, "Reliever": 0, "Remarks": ""},
    ]

    lower_df = pd.DataFrame(rows)
    numeric_cols = [
        "BE_Scientific_Manpower", "BE_Final_Manpower", "General_Shift",
        "Shift_A", "Shift_B", "Shift_C", "Reliever",
    ]
    for col in numeric_cols:
        lower_df[col] = pd.to_numeric(lower_df[col], errors="coerce").fillna(0).round(2)

    return lower_df, {
        "sum_no_of_drums_total": round_2(sum_no_of_drums_total),
        "sum_tfo_required_shift_total": round_2(sum_tfo_required_shift_total),
        "no_of_drums_426k": round_2(no_of_drums_426k),
    }


def sanitize_editor_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    safe_df = df.copy()
    safe_df = safe_df.replace([math.inf, -math.inf], pd.NA)

    for col in NUMERIC_COLUMNS:
        if col in safe_df.columns:
            safe_df[col] = pd.to_numeric(safe_df[col], errors="coerce")

    for col in TEXT_COLUMNS:
        if col in safe_df.columns:
            safe_df[col] = safe_df[col].fillna("").astype(str)

    return safe_df


def normalize_compare_df(df: pd.DataFrame) -> pd.DataFrame:
    compare_df = df.copy()

    for col in compare_df.columns:
        if col in NUMERIC_COLUMNS or col == "BE_Scientific_Manpower":
            compare_df[col] = pd.to_numeric(compare_df[col], errors="coerce").fillna(0).round(4)
        else:
            compare_df[col] = compare_df[col].fillna("").astype(str).str.strip()

    return compare_df.reset_index(drop=True)


def dataframes_equal_for_ui(left_df: pd.DataFrame, right_df: pd.DataFrame) -> bool:
    return normalize_compare_df(left_df).equals(normalize_compare_df(right_df))


def get_initial_full_spinning_df(source_df: pd.DataFrame, lower_tfo_df: pd.DataFrame) -> pd.DataFrame:
    non_tfo_df = source_df.loc[source_df["Section"].astype(str).str.upper() != "TFO"].copy()
    original_tfo_rows = source_df.loc[source_df["Section"].astype(str).str.upper() == "TFO", "Excel_Row_No"].tolist()

    dynamic_tfo_df = lower_tfo_df.copy()
    if len(original_tfo_rows) >= len(dynamic_tfo_df):
        dynamic_tfo_df["Excel_Row_No"] = original_tfo_rows[: len(dynamic_tfo_df)]
    else:
        start_row = int(source_df["Excel_Row_No"].max()) + 1
        dynamic_tfo_df["Excel_Row_No"] = list(range(start_row, start_row + len(dynamic_tfo_df)))

    non_tfo_df["Row_Key"] = [f"BASE_{row_no}" for row_no in non_tfo_df["Excel_Row_No"]]
    dynamic_tfo_df["Row_Key"] = [f"TFO_{i + 1}" for i in range(len(dynamic_tfo_df))]

    full_df = pd.concat([non_tfo_df, dynamic_tfo_df], ignore_index=True)
    return sanitize_editor_dataframe(full_df)


def assign_tfo_row_metadata(
    source_df: pd.DataFrame,
    lower_tfo_df: pd.DataFrame,
) -> pd.DataFrame:
    prepared_tfo_df = lower_tfo_df.copy()

    original_tfo_rows = source_df.loc[
        source_df["Section"].astype(str).str.upper() == "TFO",
        "Excel_Row_No",
    ].tolist()

    if len(original_tfo_rows) >= len(prepared_tfo_df):
        prepared_tfo_df["Excel_Row_No"] = original_tfo_rows[: len(prepared_tfo_df)]
    else:
        start_row = int(source_df["Excel_Row_No"].max()) + 1
        prepared_tfo_df["Excel_Row_No"] = list(range(start_row, start_row + len(prepared_tfo_df)))

    prepared_tfo_df["Row_Key"] = [f"TFO_{i + 1}" for i in range(len(prepared_tfo_df))]
    return prepared_tfo_df


def sum_spinning_reference(current_df: pd.DataFrame, col_letter: str, start_row: int, end_row: int) -> float:
    column_map = {
        "K": "BE_Scientific_Manpower",
    }
    column_name = column_map.get(col_letter.upper())
    if not column_name or column_name not in current_df.columns:
        return 0.0

    mask = current_df["Excel_Row_No"].between(start_row, end_row, inclusive="both")
    values = pd.to_numeric(current_df.loc[mask, column_name], errors="coerce").fillna(0)
    return float(values.sum())


def upper_tfo_reference_value(upper_df: pd.DataFrame, col_letter: str, row_no: int) -> float:
    column_map = {
        "N": "No. of TFO Required / shift",
        "T": "No. of Drums",
    }
    column_name = column_map.get(col_letter.upper())
    if not column_name or column_name not in upper_df.columns:
        return 0.0

    matched = upper_df.loc[upper_df["Upper_Row_No"] == row_no, column_name]
    if matched.empty:
        return 0.0
    return safe_float(matched.iloc[0], 0.0)


def evaluate_formula(formula_text: str, current_df: pd.DataFrame, upper_tfo_df: pd.DataFrame):
    if formula_text is None:
        return None

    expression = str(formula_text).strip()
    if expression == "" or expression.lower() == "nan":
        return None

    expression = expression.lstrip("=")
    expression = expression.replace("^", "**")
    expression = re.sub(
        r"(\d+(?:\.\d+)?)\s*%",
        lambda match: f"({match.group(1)}/100)",
        expression,
    )

    range_pattern = re.compile(r"SUM\(\+?([A-Z]+)(\d+):([A-Z]+)(\d+)\)", re.IGNORECASE)

    def range_replacer(match):
        start_col = match.group(1).upper()
        end_col = match.group(3).upper()
        start_row = int(match.group(2))
        end_row = int(match.group(4))

        if start_col != end_col:
            return "0"

        if start_col in {"K"}:
            return str(sum_spinning_reference(current_df, start_col, start_row, end_row))

        if start_col in {"N", "T"}:
            return str(sum(upper_tfo_reference_value(upper_tfo_df, start_col, row_no) for row_no in range(start_row, end_row + 1)))

        return "0"

    expression = range_pattern.sub(range_replacer, expression)

    single_ref_pattern = re.compile(r"(?<![A-Z])([KNT])(\d+)(?![A-Z])", re.IGNORECASE)

    def single_ref_replacer(match):
        col_letter = match.group(1).upper()
        row_no = int(match.group(2))

        if col_letter == "K":
            return str(sum_spinning_reference(current_df, col_letter, row_no, row_no))

        return str(upper_tfo_reference_value(upper_tfo_df, col_letter, row_no))

    expression = single_ref_pattern.sub(single_ref_replacer, expression)
    expression = re.sub(r"\bROUNDUP\s*\(", "excel_roundup(", expression, flags=re.IGNORECASE)
    expression = re.sub(r"\bROUND\s*\(", "excel_round(", expression, flags=re.IGNORECASE)

    safe_namespace = {
        "excel_round": excel_round,
        "excel_roundup": excel_roundup,
        "math": math,
        "abs": abs,
        "min": min,
        "max": max,
    }

    try:
        result = eval(expression, {"__builtins__": {}}, safe_namespace)
        result = safe_float(result, 0.0)
        if math.isfinite(result):
            return round(result, 2)
        return 0.0
    except Exception:
        return None


def recalculate_scientific_manpower(full_df: pd.DataFrame, upper_tfo_df: pd.DataFrame) -> pd.DataFrame:
    recalculated_df = full_df.copy()
    recalculated_df = recalculated_df.sort_values("Excel_Row_No").reset_index(drop=True)

    for index, row in recalculated_df.iterrows():
        formula_value = row.get("Formulas", "")
        evaluated_value = evaluate_formula(formula_value, recalculated_df, upper_tfo_df)

        if evaluated_value is not None:
            recalculated_df.at[index, "BE_Scientific_Manpower"] = round_2(evaluated_value)
        else:
            existing_value = row.get("BE_Scientific_Manpower")
            if pd.isna(existing_value) or existing_value == "":
                recalculated_df.at[index, "BE_Scientific_Manpower"] = pd.NA
            else:
                recalculated_df.at[index, "BE_Scientific_Manpower"] = round_2(existing_value)

    return sanitize_editor_dataframe(recalculated_df)


def build_summary_table(full_spinning_df: pd.DataFrame) -> pd.DataFrame:
    summary_df = full_spinning_df.groupby(["Location", "Business", "Section"], as_index=False, dropna=False)["BE_Final_Manpower"].sum()
    summary_df["BE_Final_Manpower"] = pd.to_numeric(summary_df["BE_Final_Manpower"], errors="coerce").fillna(0).round(0).astype(int)
    return summary_df.sort_values(["Location", "Business", "Section"]).reset_index(drop=True)


def apply_editor_changes(master_df: pd.DataFrame, edited_view_df: pd.DataFrame) -> pd.DataFrame:
    updated_df = master_df.copy()
    editable_plus_key = ["Row_Key"] + EDITABLE_COLUMNS
    changed_df = edited_view_df[editable_plus_key].copy()
    changed_df = sanitize_editor_dataframe(changed_df)
    updated_df = updated_df.drop(columns=EDITABLE_COLUMNS, errors="ignore").merge(changed_df, on="Row_Key", how="left")
    return sanitize_editor_dataframe(updated_df)


def rebuild_full_spinning_with_tfo(
    source_df: pd.DataFrame,
    base_full_df: pd.DataFrame,
    upper_tfo_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, dict]:
    fresh_lower_tfo_df, driver_values = calculate_lower_tfo_manpower(upper_tfo_df)
    fresh_lower_tfo_df["BE_Final_Manpower"] = fresh_lower_tfo_df["BE_Scientific_Manpower"]
    fresh_lower_tfo_df = assign_tfo_row_metadata(source_df, fresh_lower_tfo_df)

    non_tfo_df = base_full_df.loc[
        base_full_df["Section"].astype(str).str.upper() != "TFO"
    ].copy()

    updated_full_df = pd.concat([non_tfo_df, fresh_lower_tfo_df], ignore_index=True)
    updated_full_df = sanitize_editor_dataframe(updated_full_df)
    updated_full_df = recalculate_scientific_manpower(updated_full_df, upper_tfo_df)

    tfo_mask = updated_full_df["Section"].astype(str).str.upper() == "TFO"
    updated_full_df.loc[tfo_mask, "BE_Final_Manpower"] = updated_full_df.loc[tfo_mask, "BE_Scientific_Manpower"]

    updated_full_df = sanitize_editor_dataframe(updated_full_df)
    current_lower_df = updated_full_df.loc[tfo_mask, DISPLAY_COLUMNS].copy()

    return updated_full_df, current_lower_df, driver_values


def create_download_workbook(
    summary_df: pd.DataFrame,
    full_spinning_df: pd.DataFrame,
    upper_tfo_df: pd.DataFrame,
    lower_tfo_df: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    export_spinning_df = full_spinning_df[DISPLAY_COLUMNS].copy()
    export_tfo_upper_df = upper_tfo_df.drop(columns=["Upper_Row_No"], errors="ignore").copy()
    export_tfo_lower_df = lower_tfo_df[DISPLAY_COLUMNS].copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        export_spinning_df.to_excel(writer, sheet_name="Spinning", index=False)

        tfo_sheet_name = "TFO"
        export_tfo_upper_df.to_excel(writer, sheet_name=tfo_sheet_name, index=False, startrow=1)
        lower_start_row = len(export_tfo_upper_df) + 5
        export_tfo_lower_df.to_excel(writer, sheet_name=tfo_sheet_name, index=False, startrow=lower_start_row)

        tfo_sheet = writer.sheets[tfo_sheet_name]
        summary_sheet = writer.sheets["Summary"]
        spinning_sheet = writer.sheets["Spinning"]

        header_fill = PatternFill("solid", fgColor="0F4C81")
        header_font = Font(color="FFFFFF", bold=True)
        title_fill = PatternFill("solid", fgColor="D9EAF7")
        title_font = Font(color="0F172A", bold=True)

        tfo_sheet["A1"] = "Upper TFO Production Table"
        tfo_sheet["A1"].font = title_font
        tfo_sheet["A1"].fill = title_fill

        lower_title_row = lower_start_row + 1
        tfo_sheet.cell(row=lower_title_row, column=1, value="Lower TFO Manpower Table")
        tfo_sheet.cell(row=lower_title_row, column=1).font = title_font
        tfo_sheet.cell(row=lower_title_row, column=1).fill = title_fill

        for sheet in [summary_sheet, spinning_sheet, tfo_sheet]:
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
                for cell in row:
                    if cell.row in {1, 2, lower_start_row + 1, lower_start_row + 2}:
                        if isinstance(cell.value, str) and cell.value not in {"Upper TFO Production Table", "Lower TFO Manpower Table"}:
                            cell.fill = header_fill
                            cell.font = header_font

            for column_cells in sheet.columns:
                max_length = 0
                column_letter = get_column_letter(column_cells[0].column)
                for cell in column_cells:
                    cell_value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(cell_value))
                sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 34)

    output.seek(0)
    return output.getvalue()


def render_metric_card(title: str, value: str, note: str):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


source_spinning_df, source_path = load_spinning_master()
if source_spinning_df is None:
    st.error("Spinning.xlsx file was not found. Please keep the file in the configured path or in the same folder as the app.")
    st.stop()

current_upper_df = calculate_upper_tfo_metrics(st.session_state.tfo_input_df)

if "full_spinning_df" not in st.session_state:
    initial_lower_tfo_df, _ = calculate_lower_tfo_manpower(current_upper_df)
    st.session_state.full_spinning_df = get_initial_full_spinning_df(source_spinning_df, initial_lower_tfo_df)
    st.session_state.full_spinning_df = recalculate_scientific_manpower(st.session_state.full_spinning_df, current_upper_df)
    tfo_mask_init = st.session_state.full_spinning_df["Section"].astype(str).str.upper() == "TFO"
    st.session_state.full_spinning_df.loc[tfo_mask_init, "BE_Final_Manpower"] = st.session_state.full_spinning_df.loc[tfo_mask_init, "BE_Scientific_Manpower"]

st.session_state.full_spinning_df = recalculate_scientific_manpower(st.session_state.full_spinning_df, current_upper_df)
full_spinning_df = st.session_state.full_spinning_df.copy()
summary_df = build_summary_table(full_spinning_df)

st.markdown(
    """
    <div class="hero-wrap">
        <div class="hero-title">Spinning Manpower Dashboard</div>
        <div class="hero-subtitle">Executive view of section-wise manpower, editable planning tables, and the TFO manpower engine.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

kpi_col_1, kpi_col_2, kpi_col_3, kpi_col_4 = st.columns(4)
with kpi_col_1:
    render_metric_card(
        "Total Final Manpower",
        f"{int(round(pd.to_numeric(full_spinning_df['BE_Final_Manpower'], errors='coerce').fillna(0).sum())):,}",
        "Across all spinning sections",
    )
with kpi_col_2:
    render_metric_card(
        "Total Scientific Manpower",
        f"{pd.to_numeric(full_spinning_df['BE_Scientific_Manpower'], errors='coerce').fillna(0).sum():,.2f}",
        "Driven by formulas and workload logic",
    )
with kpi_col_3:
    render_metric_card("Sections Covered", f"{full_spinning_df['Section'].nunique():,}", "Operational coverage in the dashboard")
with kpi_col_4:
    render_metric_card("Locations", f"{full_spinning_df['Location'].nunique():,}", "Current active location scope")

summary_tab, spinning_tab, tfo_tab = st.tabs(["Summary", "Entire Spinning Table", "TFO"])

with summary_tab:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Executive filters</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Use the filters below to focus the summary view without changing the underlying data.</div>', unsafe_allow_html=True)

    filter_col_1, filter_col_2, filter_col_3 = st.columns([1, 1, 2])

    with filter_col_1:
        selected_location = st.selectbox(
            "Location",
            options=sorted(summary_df["Location"].dropna().unique().tolist()),
        )

    with filter_col_2:
        business_options = sorted(
            summary_df.loc[summary_df["Location"] == selected_location, "Business"].dropna().unique().tolist()
        )
        selected_business = st.selectbox("Business", options=business_options)

    with filter_col_3:
        section_options = sorted(
            summary_df.loc[
                (summary_df["Location"] == selected_location)
                & (summary_df["Business"] == selected_business),
                "Section",
            ].dropna().unique().tolist()
        )
        selected_sections = st.multiselect(
            "Section filter",
            options=section_options,
            default=section_options,
        )

    st.markdown('</div>', unsafe_allow_html=True)

    filtered_summary_df = summary_df.loc[
        (summary_df["Location"] == selected_location)
        & (summary_df["Business"] == selected_business)
    ].copy()

    if selected_sections:
        filtered_summary_df = filtered_summary_df.loc[filtered_summary_df["Section"].isin(selected_sections)].copy()
    else:
        filtered_summary_df = filtered_summary_df.iloc[0:0].copy()

    summary_total = int(round(pd.to_numeric(filtered_summary_df["BE_Final_Manpower"], errors="coerce").fillna(0).sum()))
    summary_sections = int(filtered_summary_df["Section"].nunique())

    metric_col_1, metric_col_2, metric_col_3 = st.columns(3)

    with metric_col_1:
        render_metric_card("Filtered Final Manpower", f"{summary_total:,}", "Visible summary selection")
    with metric_col_2:
        render_metric_card("Visible Sections", f"{summary_sections:,}", "Current filter output")
    with metric_col_3:
        top_section = "-"
        if not filtered_summary_df.empty:
            top_row = filtered_summary_df.sort_values("BE_Final_Manpower", ascending=False).iloc[0]
            top_section = f"{top_row['Section']} ({int(top_row['BE_Final_Manpower'])})"
        render_metric_card("Top Section", top_section, "Highest final manpower in current view")

    chart_col_1, chart_col_2 = st.columns([1.4, 1])

    with chart_col_1:
        st.markdown('<div class="panel-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Section-wise final manpower</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-subtitle">Quick comparison of final manpower across the selected sections.</div>', unsafe_allow_html=True)

        if filtered_summary_df.empty:
            st.info("No sections selected. Please choose one or more sections.")
        else:
            bar_chart = (
                alt.Chart(filtered_summary_df)
                .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
                .encode(
                    x=alt.X("Section:N", sort="-y", title="Section"),
                    y=alt.Y("BE_Final_Manpower:Q", title="Final manpower"),
                    tooltip=["Location", "Business", "Section", "BE_Final_Manpower"],
                )
                .properties(height=360)
            )
            st.altair_chart(bar_chart, use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True)

    with chart_col_2:
        st.markdown('<div class="panel-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Shift mix</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-subtitle">Plant-level shift allocation from the current full table.</div>', unsafe_allow_html=True)

        shift_summary = pd.DataFrame(
            {
                "Shift": ["General", "A", "B", "C"],
                "Manpower": [
                    pd.to_numeric(full_spinning_df["General_Shift"], errors="coerce").fillna(0).sum(),
                    pd.to_numeric(full_spinning_df["Shift_A"], errors="coerce").fillna(0).sum(),
                    pd.to_numeric(full_spinning_df["Shift_B"], errors="coerce").fillna(0).sum(),
                    pd.to_numeric(full_spinning_df["Shift_C"], errors="coerce").fillna(0).sum(),
                ],
            }
        )

        if shift_summary["Manpower"].sum() <= 0:
            st.info("Shift data is not available.")
        else:
            shift_chart = (
                alt.Chart(shift_summary)
                .mark_arc(innerRadius=55, outerRadius=105)
                .encode(
                    theta=alt.Theta("Manpower:Q"),
                    color=alt.Color("Shift:N", legend=alt.Legend(orient="bottom")),
                    tooltip=["Shift", "Manpower"],
                )
                .properties(height=350)
            )
            st.altair_chart(shift_chart, use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Summary table</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Section-wise final manpower ready for leadership review and download.</div>', unsafe_allow_html=True)

    if filtered_summary_df.empty:
        st.info("No rows to display for the current filter selection.")
    else:
        st.dataframe(
            filtered_summary_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "BE_Final_Manpower": st.column_config.NumberColumn("BE_Final_Manpower", format="%d"),
            },
        )

    current_lower_for_download = full_spinning_df.loc[
        full_spinning_df["Section"].astype(str).str.upper() == "TFO",
        DISPLAY_COLUMNS,
    ].copy()

    st.download_button(
        "Download final Excel",
        data=create_download_workbook(
            summary_df=build_summary_table(full_spinning_df),
            full_spinning_df=full_spinning_df,
            upper_tfo_df=current_upper_df,
            lower_tfo_df=current_lower_for_download,
        ),
        file_name="Spinning_Manpower_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown('</div>', unsafe_allow_html=True)

with spinning_tab:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Editable spinning manpower table</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Edit formulas and manpower fields. As soon as you update a formula and press Enter, BE_Scientific_Manpower refreshes automatically.</div>', unsafe_allow_html=True)

    spin_col_1, spin_col_2, spin_col_3 = st.columns([1.1, 1.3, 1.8])

    with spin_col_1:
        detail_location = st.selectbox(
            "Location filter",
            options=sorted(full_spinning_df["Location"].dropna().astype(str).unique().tolist()),
            key="detail_location",
        )

    with spin_col_2:
        detail_section_options = sorted(
            full_spinning_df.loc[
                full_spinning_df["Location"].astype(str) == str(detail_location),
                "Section",
            ].dropna().astype(str).unique().tolist()
        )
        detail_sections = st.multiselect(
            "Section filter",
            options=detail_section_options,
            default=detail_section_options,
            key="detail_section_filter",
        )

    with spin_col_3:
        detail_search = st.text_input(
            "Search department / designation",
            placeholder="Type to filter",
            key="detail_search",
        )

    editor_view_df = full_spinning_df.loc[
        full_spinning_df["Location"].astype(str) == str(detail_location)
    ].copy()

    if detail_sections:
        editor_view_df = editor_view_df.loc[
            editor_view_df["Section"].astype(str).isin(detail_sections)
        ].copy()
    else:
        editor_view_df = editor_view_df.iloc[0:0].copy()

    if detail_search.strip():
        search_mask = (
            editor_view_df["Department"].astype(str).str.contains(detail_search, case=False, na=False)
            | editor_view_df["Designation"].astype(str).str.contains(detail_search, case=False, na=False)
        )
        editor_view_df = editor_view_df.loc[search_mask].copy()

    if editor_view_df.empty:
        st.info("No rows available for the current filter selection.")
    else:
        editor_display_df = editor_view_df[["Row_Key"] + DISPLAY_COLUMNS].copy()
        editor_widget_key = f"full_spinning_editor_{st.session_state.full_spinning_editor_version}"

        edited_view_df = st.data_editor(
            editor_display_df,
            use_container_width=True,
            hide_index=True,
            height=680,
            key=editor_widget_key,
            disabled=[
                "Row_Key",
                "Location",
                "Business",
                "Section",
                "Sr_No",
                "Department",
                "Designation",
                "Machine_Count",
                "Workload",
                "BE_Scientific_Manpower",
            ],
            column_config={
                "Row_Key": st.column_config.TextColumn("Row_Key", width="small"),
                "BE_Scientific_Manpower": st.column_config.NumberColumn("BE_Scientific_Manpower", format="%.2f"),
                "BE_Final_Manpower": st.column_config.NumberColumn("BE_Final_Manpower", format="%.2f"),
                "General_Shift": st.column_config.NumberColumn("General_Shift", format="%.2f"),
                "Shift_A": st.column_config.NumberColumn("Shift_A", format="%.2f"),
                "Shift_B": st.column_config.NumberColumn("Shift_B", format="%.2f"),
                "Shift_C": st.column_config.NumberColumn("Shift_C", format="%.2f"),
                "Reliever": st.column_config.NumberColumn("Reliever", format="%.2f"),
            },
        )

        updated_master_df = apply_editor_changes(full_spinning_df, edited_view_df)
        updated_master_df = recalculate_scientific_manpower(updated_master_df, current_upper_df)

        before_compare_df = full_spinning_df[["Row_Key"] + EDITABLE_COLUMNS + ["BE_Scientific_Manpower"]].sort_values("Row_Key")
        after_compare_df = updated_master_df[["Row_Key"] + EDITABLE_COLUMNS + ["BE_Scientific_Manpower"]].sort_values("Row_Key")

        if not dataframes_equal_for_ui(before_compare_df, after_compare_df):
            st.session_state.full_spinning_df = sanitize_editor_dataframe(updated_master_df)
            st.session_state.full_spinning_editor_version += 1
            st.rerun()

        st.session_state.full_spinning_df = sanitize_editor_dataframe(updated_master_df)
        full_spinning_df = st.session_state.full_spinning_df.copy()

        st.markdown(
            '<div class="small-note">Tip: formula edits now trigger a clean rerender, so the Scientific Manpower column updates immediately after you confirm the edit.</div>',
            unsafe_allow_html=True,
        )

    st.markdown('</div>', unsafe_allow_html=True)

with tfo_tab:
    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">TFO planning and manpower engine</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Editable TFO production inputs with automatic roll-through into the entire spinning table and final summary.</div>', unsafe_allow_html=True)

    input_columns = [
        "Count",
        "Customer",
        "Count2",
        "Speed",
        "TPI",
        "Utilization",
        "Efficiency",
        "Production Required / day Kgs",
        "TFO divisor",
        "mpm",
        "Eff",
        "Machine divisor",
    ]

    edited_tfo_input_df = st.data_editor(
        st.session_state.tfo_input_df[input_columns],
        use_container_width=True,
        num_rows="dynamic",
        key="tfo_editor",
        hide_index=True,
        column_config={
            "Count": st.column_config.TextColumn("Count"),
            "Customer": st.column_config.TextColumn("Customer"),
            "Count2": st.column_config.NumberColumn("Count2", format="%.2f"),
            "Speed": st.column_config.NumberColumn("Speed", format="%.2f"),
            "TPI": st.column_config.NumberColumn("TPI", format="%.2f"),
            "Utilization": st.column_config.NumberColumn("Utilization", format="%.2f"),
            "Efficiency": st.column_config.NumberColumn("Efficiency", format="%.2f"),
            "Production Required / day Kgs": st.column_config.NumberColumn("Production Required / day Kgs", format="%.2f"),
            "TFO divisor": st.column_config.NumberColumn("TFO divisor", format="%.2f"),
            "mpm": st.column_config.NumberColumn("mpm", format="%.2f"),
            "Eff": st.column_config.NumberColumn("Eff", format="%.2f"),
            "Machine divisor": st.column_config.NumberColumn("Machine divisor", format="%.2f"),
        },
    )

    for col in [
        "Count2",
        "Speed",
        "TPI",
        "Utilization",
        "Efficiency",
        "Production Required / day Kgs",
        "TFO divisor",
        "mpm",
        "Eff",
        "Machine divisor",
    ]:
        edited_tfo_input_df[col] = pd.to_numeric(edited_tfo_input_df[col], errors="coerce").fillna(0).round(2)

    st.session_state.tfo_input_df = edited_tfo_input_df.copy()

    current_upper_df = calculate_upper_tfo_metrics(edited_tfo_input_df)
    current_upper_total_df = build_upper_total_row(current_upper_df)
    current_upper_final_df = pd.concat(
        [current_upper_df.drop(columns=["Upper_Row_No"], errors="ignore"), current_upper_total_df],
        ignore_index=True,
    )

    updated_full_df, current_lower_df, current_driver_values = rebuild_full_spinning_with_tfo(
        source_df=source_spinning_df,
        base_full_df=st.session_state.full_spinning_df,
        upper_tfo_df=current_upper_df,
    )

    st.session_state.full_spinning_df = updated_full_df
    full_spinning_df = updated_full_df.copy()

    tfo_metric_1, tfo_metric_2, tfo_metric_3, tfo_metric_4 = st.columns(4)

    with tfo_metric_1:
        render_metric_card("Total No. of Drums", f"{current_driver_values['sum_no_of_drums_total']:.2f}", "Calculated from TFO production inputs")
    with tfo_metric_2:
        render_metric_card("TFO Required / Shift", f"{current_driver_values['sum_tfo_required_shift_total']:.2f}", "Based on No. of Drums Required and divisor")
    with tfo_metric_3:
        render_metric_card("Drums for 4/2/6 K", f"{current_driver_values['no_of_drums_426k']:.2f}", "Used in Jumbo Assembly Winding")
    with tfo_metric_4:
        render_metric_card(
            "Lower TFO Final Manpower",
            f"{int(round(pd.to_numeric(current_lower_df['BE_Final_Manpower'], errors='coerce').fillna(0).sum())):,}",
            "Current TFO rows in final table",
        )

    st.markdown("#### Upper TFO Production Table")
    upper_display_columns = [
        "Count",
        "Customer",
        "Count2",
        "Speed",
        "TPI",
        "Utilization",
        "Efficiency",
        "Production per Drum/day",
        "Production Required / day Kgs",
        "Production Required / Month Kgs",
        "No. of Drums Required",
        "No. of TFO Required / shift",
        "mpm",
        "Eff",
        "kgs/drum/day",
        "No. of Drums",
        "no. of machines",
    ]
    st.dataframe(
        current_upper_final_df[upper_display_columns],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("#### Lower TFO Manpower Table")
    st.dataframe(
        current_lower_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "BE_Scientific_Manpower": st.column_config.NumberColumn("BE_Scientific_Manpower", format="%.2f"),
            "BE_Final_Manpower": st.column_config.NumberColumn("BE_Final_Manpower", format="%.2f"),
        },
    )

    with st.expander("Formula logic"):
        st.markdown(
            f"""
            **Upper table**
            - Production Required / Month Kgs = Production Required / day Kgs × 30
            - Production per Drum/day = `((Speed × 60 × 8 × Efficiency × 2) / (TPI × 36 × 840 × Count2 × 2.202)) × 3`
            - No. of Drums Required = Production Required / day Kgs / Production per Drum/day
            - No. of TFO Required / shift = No. of Drums Required / TFO divisor
            - kgs/drum/day = `((mpm × Eff × 8 × 60 × 1.09) / (Count2 × 840 × 2.202)) × 3`
            - No. of Drums = Production Required / day Kgs / kgs/drum/day
            - no. of machines = No. of Drums / Machine divisor

            **Lower table**
            - Assembly winding = `ROUNDUP((SUM(T2:T18)/36)*3,0)`
            - Jumbo Assembly Winding = `ROUND(ROUNDUP(T19,0)/16,0)*2`
            - TFO Operator = `ROUND(SUM(N2:N18)/6,0)*3`
            - TFO Operator (Doffer) = `ROUNDUP(SUM(N2:N18)/4,0)*3`

            **Current driver values**
            - Total No. of Drums = {current_driver_values['sum_no_of_drums_total']:.2f}
            - Total No. of TFO Required / shift = {current_driver_values['sum_tfo_required_shift_total']:.2f}
            - No. of Drums for 4/2/6 K = {current_driver_values['no_of_drums_426k']:.2f}
            """
        )

    action_col_1, action_col_2 = st.columns(2)

    with action_col_1:
        if st.button("Reset TFO Table"):
            st.session_state.tfo_input_df = get_initial_tfo_data()
            reset_upper_df = calculate_upper_tfo_metrics(st.session_state.tfo_input_df)
            reset_full_df, _, _ = rebuild_full_spinning_with_tfo(
                source_df=source_spinning_df,
                base_full_df=st.session_state.full_spinning_df,
                upper_tfo_df=reset_upper_df,
            )
            st.session_state.full_spinning_df = reset_full_df
            st.session_state.full_spinning_editor_version += 1
            st.rerun()

    with action_col_2:
        if st.button("Reset Full Spinning Table from Source"):
            fresh_upper_df = calculate_upper_tfo_metrics(st.session_state.tfo_input_df)
            fresh_lower_df, _ = calculate_lower_tfo_manpower(fresh_upper_df)
            st.session_state.full_spinning_df = get_initial_full_spinning_df(source_spinning_df, fresh_lower_df)
            st.session_state.full_spinning_df = recalculate_scientific_manpower(st.session_state.full_spinning_df, fresh_upper_df)
            tfo_mask = st.session_state.full_spinning_df["Section"].astype(str).str.upper() == "TFO"
            st.session_state.full_spinning_df.loc[tfo_mask, "BE_Final_Manpower"] = st.session_state.full_spinning_df.loc[tfo_mask, "BE_Scientific_Manpower"]
            st.session_state.full_spinning_editor_version += 1
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)