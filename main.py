import streamlit as st
import base64
import pandas as pd
from pathlib import Path
import streamlit.components.v1 as components

st.set_page_config(page_title="Adevali Insurance", layout="wide")

APP_DIR = Path(__file__).resolve().parent

BACKGROUND_FILE = APP_DIR / "Copy of Flyer Life Insurance.png"
SIDE_IMAGE_FILE = APP_DIR / "logow.png"
MORTALITY_FILE = APP_DIR / "Mortality_Table.xlsx"


# =========================
# HELPER FUNCTIONS
# =========================

def get_base64_image(image_path):
    image_path = Path(image_path)

    with open(image_path, "rb") as image_file:
        encoded_image = base64.b64encode(image_file.read()).decode()

    return encoded_image


def rupiah(value):
    return f"Rp{value:,.2f}"


def rupiah_no_decimal(value):
    return f"Rp{value:,.0f}"


def standardize_mortality_columns(table):
    df = table.copy()
    df.columns = [str(col).strip() for col in df.columns]

    rename_map = {}

    for col in df.columns:
        cleaned = (
            str(col)
            .strip()
            .lower()
            .replace(" ", "")
            .replace("_", "")
            .replace("-", "")
        )

        if cleaned in ["age", "x", "age(x)"]:
            rename_map[col] = "age (x)"

        elif cleaned in ["qxtau", "qx(tau)", "qtau", "qx"]:
            rename_map[col] = "q_x(tau)"

        elif cleaned in ["qxaccident", "q_xaccident", "qx(accident)", "accident"]:
            rename_map[col] = "q_x(accident)"

        elif cleaned in ["qxillness", "q_xillness", "qx(illness)", "illness"]:
            rename_map[col] = "q_x(illness)"

        elif cleaned in [
            "qother",
            "qxother",
            "qx(othercauses)",
            "q(othercauses)",
            "q_(othercauses)",
            "other",
            "othercauses"
        ]:
            rename_map[col] = "q_(other causes)"

    df = df.rename(columns=rename_map)

    return df


@st.cache_data
def load_mortality_tables(file_path):
    mortality_tables = pd.read_excel(
        file_path,
        sheet_name=None,
        engine="openpyxl"
    )

    standardized_tables = {}

    for sheet_name, table in mortality_tables.items():
        standardized_tables[sheet_name] = standardize_mortality_columns(table)

    sheet_names = list(standardized_tables.keys())

    male_table = None
    female_table = None

    for sheet_name in sheet_names:
        lower_name = sheet_name.lower()

        if "male" in lower_name and "female" not in lower_name:
            male_table = standardized_tables[sheet_name]

        if "female" in lower_name:
            female_table = standardized_tables[sheet_name]

    if male_table is None and len(sheet_names) >= 1:
        male_table = standardized_tables[sheet_names[0]]

    if female_table is None and len(sheet_names) >= 2:
        female_table = standardized_tables[sheet_names[1]]

    if male_table is None or female_table is None:
        raise ValueError(
            "Could not find male and female mortality tables. "
            "Your Excel file should have two sheets, preferably named Male and Female."
        )

    return male_table, female_table


def clean_table(table, required_columns):
    df = table.copy()

    missing_columns = []

    for column in required_columns:
        if column not in df.columns:
            missing_columns.append(column)

    if missing_columns:
        raise ValueError(
            "Missing columns in Mortality_Table.xlsx: "
            + ", ".join(missing_columns)
        )

    for column in required_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(subset=required_columns)
    df["age (x)"] = df["age (x)"].astype(int)
    df = df.sort_values("age (x)")

    return df


# =========================
# LOAD MORTALITY TABLE SILENTLY
# =========================

try:
    male_table, female_table = load_mortality_tables(MORTALITY_FILE)

except Exception as error:
    st.error("Mortality_Table.xlsx could not be read.")

    st.write("Python is trying to read the file from this exact location:")
    st.code(str(MORTALITY_FILE))

    st.write("Actual error:")
    st.code(str(error))

    st.write("Files found in the same folder as main.py:")
    st.code("\n".join([file.name for file in APP_DIR.iterdir()]))

    st.stop()


# =========================
# ACTUARIAL ASSUMPTIONS
# =========================

i = 0.0525
deferred_period = 5

tiers = {
    "Premium": {
        "accident_bt": 2_000_000_000,
        "other_bt": 1_500_000_000,
        "illness_bt": 1_000_000_000
    },
    "Standard": {
        "accident_bt": 1_000_000_000,
        "other_bt": 950_000_000,
        "illness_bt": 750_000_000
    },
    "Basic": {
        "accident_bt": 500_000_000,
        "other_bt": 475_000_000,
        "illness_bt": 300_000_000
    }
}

e0 = 1_810_000
e1 = 60_000
c0 = 0.60
c1 = 0.30
settlement_cost = 250_000

smoker_loading_table = {
    "none": 0.00,
    "light": 0.15,
    "heavy": 0.25
}

alcohol_loading_table = {
    "none": 0.00,
    "light": 0.20,
    "moderate": 0.25
}

dangerous_hobby_loading_table = {
    "no": 0.00,
    "yes": 0.50
}


# =========================
# PREMIUM FORMULAS
# =========================

def calculate_apv_total_whole_life(table, x, accident_bt, other_bt, illness_bt):
    required_columns = [
        "age (x)",
        "q_x(tau)",
        "q_x(accident)",
        "q_x(illness)",
        "q_(other causes)"
    ]

    df = clean_table(table, required_columns)

    if x not in df["age (x)"].values:
        raise ValueError(f"Age {x} is not found in the mortality table.")

    v = 1 / (1 + i)

    apv_accident = 0
    apv_other = 0
    apv_illness = 0

    survival_probability = 1
    max_age = int(df["age (x)"].max())

    for age in range(x, max_age):
        row = df[df["age (x)"] == age]

        if row.empty:
            break

        row = row.iloc[0]

        k = age - x

        q_accident = row["q_x(accident)"]
        q_illness = row["q_x(illness)"]
        q_other = row["q_(other causes)"]
        q_tau = row["q_x(tau)"]

        discount_factor = v ** (k + 1)

        apv_accident = (
            apv_accident
            +
            accident_bt
            *
            discount_factor
            *
            survival_probability
            *
            q_accident
        )

        apv_other = (
            apv_other
            +
            other_bt
            *
            discount_factor
            *
            survival_probability
            *
            q_other
        )

        if k >= deferred_period:
            apv_illness = (
                apv_illness
                +
                illness_bt
                *
                discount_factor
                *
                survival_probability
                *
                q_illness
            )

        survival_probability = survival_probability * (1 - q_tau)

    apv_total = apv_accident + apv_other + apv_illness

    return apv_accident, apv_other, apv_illness, apv_total


def calculate_annuity_whole_life(table, x):
    required_columns = [
        "age (x)",
        "q_x(tau)"
    ]

    df = clean_table(table, required_columns)

    if x not in df["age (x)"].values:
        raise ValueError(f"Age {x} is not found in the mortality table.")

    v = 1 / (1 + i)

    omega = int(df["age (x)"].max())

    annuity_annual = 0
    annuity_monthly = 0

    survival_probability = 1

    for k in range(0, omega - x):
        age = x + k

        row = df[df["age (x)"] == age]

        if row.empty:
            break

        row = row.iloc[0]

        q_tau = row["q_x(tau)"]

        annual_contribution = (
            v ** k
            *
            survival_probability
        )

        annuity_annual = annuity_annual + annual_contribution

        for r in range(0, 12):
            monthly_fraction = r / 12

            monthly_discount_factor = v ** (k + monthly_fraction)

            monthly_survival_probability = (
                survival_probability
                *
                (1 - monthly_fraction * q_tau)
            )

            monthly_contribution = (
                (1 / 12)
                *
                monthly_discount_factor
                *
                monthly_survival_probability
            )

            annuity_monthly = annuity_monthly + monthly_contribution

        survival_probability = survival_probability * (1 - q_tau)

    return annuity_annual, annuity_monthly


def calculate_net_and_gross_premiums(
    apv_accident,
    apv_other,
    apv_illness,
    apv_total,
    annuity_annual,
    annuity_monthly,
    accident_bt,
    other_bt,
    illness_bt
):
    net_yearly_premium = apv_total / annuity_annual

    net_annualized_monthly_premium = apv_total / annuity_monthly
    net_monthly_premium = net_annualized_monthly_premium / 12

    A_accident = apv_accident / accident_bt
    A_other = apv_other / other_bt
    A_illness = apv_illness / illness_bt

    A_total_claim = A_accident + A_other + A_illness

    apv_settlement = settlement_cost * A_total_claim

    gross_yearly_numerator = (
        apv_total
        +
        apv_settlement
        +
        e0
        +
        e1 * (annuity_annual - 1)
    )

    gross_yearly_denominator = (
        annuity_annual
        -
        c0
        -
        c1 * (annuity_annual - 1)
    )

    gross_yearly_premium = (
        gross_yearly_numerator
        /
        gross_yearly_denominator
    )

    gross_monthly_numerator = (
        apv_total
        +
        apv_settlement
        +
        e0
        +
        e1 * (annuity_annual - 1)
    )

    gross_monthly_denominator = (
        annuity_monthly
        -
        c0
        -
        c1 * (annuity_annual - 1)
    )

    gross_annualized_monthly_premium = (
        gross_monthly_numerator
        /
        gross_monthly_denominator
    )

    gross_monthly_premium = gross_annualized_monthly_premium / 12

    return (
        net_yearly_premium,
        net_monthly_premium,
        gross_yearly_premium,
        gross_monthly_premium,
        gross_annualized_monthly_premium,
        apv_settlement
    )


def apply_underwriting_loading(
    gross_yearly_premium,
    gross_monthly_premium,
    gross_annualized_monthly_premium,
    underwriting_info
):
    underwriting_factor = underwriting_info["underwriting_factor"]

    adjusted_gross_yearly_premium = (
        gross_yearly_premium
        *
        underwriting_factor
    )

    adjusted_gross_monthly_premium = (
        gross_monthly_premium
        *
        underwriting_factor
    )

    adjusted_gross_annualized_monthly_premium = (
        gross_annualized_monthly_premium
        *
        underwriting_factor
    )

    return (
        adjusted_gross_yearly_premium,
        adjusted_gross_monthly_premium,
        adjusted_gross_annualized_monthly_premium
    )


def make_underwriting_info(smoker_level, alcohol_level, dangerous_hobby):
    smoker_loading = smoker_loading_table[smoker_level]
    alcohol_loading = alcohol_loading_table[alcohol_level]
    dangerous_hobby_loading = dangerous_hobby_loading_table[dangerous_hobby]

    total_loading = (
        smoker_loading
        +
        alcohol_loading
        +
        dangerous_hobby_loading
    )

    underwriting_factor = 1 + total_loading

    return {
        "smoker_level": smoker_level,
        "alcohol_level": alcohol_level,
        "dangerous_hobby": dangerous_hobby,
        "smoker_loading": smoker_loading,
        "alcohol_loading": alcohol_loading,
        "dangerous_hobby_loading": dangerous_hobby_loading,
        "total_loading": total_loading,
        "underwriting_factor": underwriting_factor
    }


# =========================
# WEBSITE DESIGN
# =========================

background_image = get_base64_image(BACKGROUND_FILE)
side_image = get_base64_image(SIDE_IMAGE_FILE)

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Josefin+Sans:wght@400;700;800&display=swap');

    .stApp {{
        background-image: url("data:image/png;base64,{background_image}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }}

    [data-testid="stHeader"] {{
        background-color: transparent;
    }}

    .stApp h1, .stApp h2, .stApp h3, .stApp p, .stApp label {{
        color: white;
        font-family: 'Josefin Sans', sans-serif;
    }}

    .st-key-premium_calculator_box {{
        max-width: 1180px;
        margin: 30px auto 20px auto;
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: #238EDB !important;
        padding: 25px;
        border-radius: 20px;
        border: none !important;
        font-family: 'Josefin Sans', sans-serif;
        box-sizing: border-box;
    }}

    .st-key-premium_calculator_box > div {{
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        border-radius: 20px;
    }}

    .st-key-premium_calculator_box h1,
    .st-key-premium_calculator_box h2,
    .st-key-premium_calculator_box h3,
    .st-key-premium_calculator_box p,
    .st-key-premium_calculator_box label {{
        color: #238EDB !important;
        font-family: 'Josefin Sans', sans-serif;
    }}

    .st-key-premium_calculator_box [role="radiogroup"] label p {{
        color: #238EDB !important;
    }}

    .premium-calculator-title {{
        color: #238EDB;
        font-size: 32px;
        font-weight: 800;
        margin-bottom: 20px;
        font-family: 'Josefin Sans', sans-serif;
    }}

    div[data-testid="stDataFrame"] {{
        background-color: white;
        border-radius: 20px;
        padding: 10px;
    }}

    div.stButton > button {{
        background-color: #FFFFFF !important;
        color: #238EDB !important;
        border: 1px solid #238EDB !important;
        border-radius: 15px;
        padding: 12px 25px;
        font-weight: 800;
        font-family: 'Josefin Sans', sans-serif;
    }}

    div.stButton > button:hover {{
        background-color: #F2F8FF !important;
        color: #238EDB !important;
        border: 1px solid #238EDB !important;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

components.html(
    f"""
    <style>
    body {{
        margin: 0;
        background-color: transparent;
        font-family: "Source Sans Pro", Arial, Helvetica, sans-serif;
    }}

    .text-box {{
        height: 375px;
        padding: 60px;
        display: flex;
        justify-content: flex-start;
        align-items: center;
        box-sizing: border-box;
    }}

    .content-row {{
        display: flex;
        align-items: center;
        gap: 35px;
    }}

    .left-side {{
        text-align: center;
    }}

    .side-img {{
