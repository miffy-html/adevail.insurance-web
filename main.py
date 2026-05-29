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
        font-family: 'Josefin Sans', sans-serif;
    }}

    [data-testid="stHeader"] {{
        background-color: transparent;
    }}

    .block-container {{
        max-width: 1220px;
        padding-top: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }}

    .stApp h1,
    .stApp h2,
    .stApp h3,
    .stApp p,
    .stApp label,
    .stApp li {{
        font-family: 'Josefin Sans', sans-serif;
    }}

    .st-key-hero_logo,
    .st-key-hero_text {{
        background: transparent !important;
        border: none !important;
    }}

    .hero-title {{
        color: white;
        font-size: 67px;
        font-weight: 800;
        line-height: 1.15;
        letter-spacing: 2px;
        margin-top: 35px;
    }}

    .hero-tagline {{
        color: yellow;
        font-size: 26px;
        font-weight: 800;
        letter-spacing: 2px;
        margin-top: 15px;
        margin-bottom: 35px;
    }}

    .logo-name {{
        color: white;
        font-size: 22px;
        font-weight: 700;
        text-align: center;
        margin-top: 10px;
        text-shadow: 2px 2px 6px black;
    }}

    .st-key-about_box {{
        background-color: #238EDB !important;
        color: white !important;
        padding: 25px;
        border-radius: 20px;
        min-height: 500px;
        box-sizing: border-box;
    }}

    .st-key-about_box h1,
    .st-key-about_box h2,
    .st-key-about_box h3,
    .st-key-about_box p,
    .st-key-about_box li {{
        color: white !important;
        font-family: 'Josefin Sans', sans-serif;
        font-size: 22px;
        line-height: 1.35;
    }}

    .st-key-about_box h3 {{
        font-size: 30px !important;
        font-weight: 800 !important;
        margin-bottom: 12px !important;
    }}

    .st-key-contact_box {{
        background-color: white !important;
        color: #238EDB !important;
        padding: 25px;
        border-radius: 20px;
        min-height: 500px;
        box-sizing: border-box;
    }}

    .st-key-contact_box h1,
    .st-key-contact_box h2,
    .st-key-contact_box h3,
    .st-key-contact_box p,
    .st-key-contact_box li {{
        color: #238EDB !important;
        font-family: 'Josefin Sans', sans-serif;
        font-size: 22px;
        line-height: 1.35;
    }}

    .st-key-contact_box h3 {{
        font-size: 30px !important;
        font-weight: 800 !important;
        margin-bottom: 12px !important;
    }}

    .st-key-mission_box {{
        background-color: white !important;
        color: #238EDB !important;
        padding: 25px;
        border-radius: 20px;
        margin-top: 30px;
        box-sizing: border-box;
    }}

    .st-key-mission_box h1,
    .st-key-mission_box h2,
    .st-key-mission_box h3,
    .st-key-mission_box p,
    .st-key-mission_box li {{
        color: #238EDB !important;
        font-family: 'Josefin Sans', sans-serif;
        font-size: 22px;
        line-height: 1.35;
    }}

    .st-key-mission_box h3 {{
        font-size: 30px !important;
        font-weight: 800 !important;
        margin-bottom: 12px !important;
    }}

    .st-key-products_box {{
        background-color: #238EDB !important;
        color: white !important;
        padding: 25px;
        border-radius: 20px;
        margin-top: 30px;
        box-sizing: border-box;
    }}

    .st-key-products_box h1,
    .st-key-products_box h2,
    .st-key-products_box h3,
    .st-key-products_box p,
    .st-key-products_box li {{
        color: white !important;
        font-family: 'Josefin Sans', sans-serif;
        font-size: 22px;
        line-height: 1.35;
    }}

    .st-key-products_box h3 {{
        font-size: 30px !important;
        font-weight: 800 !important;
        margin-bottom: 12px !important;
    }}

    .map-frame {{
        width: 100%;
        height: 220px;
        border: 0;
        border-radius: 15px;
        margin-top: 15px;
        pointer-events: none;
        overflow: hidden;
    }}

    .st-key-premium_calculator_box {{
        max-width: 1180px;
        margin: 30px auto 20px auto;
        background-color: white !important;
        color: #238EDB !important;
        padding: 25px;
        border-radius: 20px;
        border: none !important;
        box-sizing: border-box;
    }}

    .st-key-premium_calculator_box > div {{
        background-color: white !important;
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
    }}

    div.stButton > button {{
        background-color: white !important;
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

    div[data-testid="stDataFrame"] {{
        background-color: white;
        border-radius: 20px;
        padding: 10px;
    }}

    @media (max-width: 768px) {{
        .block-container {{
            padding-left: 1rem;
            padding-right: 1rem;
            padding-top: 0.5rem;
        }}

        .hero-title {{
            font-size: 36px;
            text-align: center;
            line-height: 1.1;
            letter-spacing: 1px;
            margin-top: 10px;
        }}

        .hero-tagline {{
            font-size: 16px;
            text-align: center;
            letter-spacing: 1px;
            margin-top: 10px;
            margin-bottom: 20px;
        }}

        .logo-name {{
            font-size: 17px;
        }}

        .st-key-about_box,
        .st-key-contact_box,
        .st-key-mission_box,
        .st-key-products_box,
        .st-key-premium_calculator_box {{
            padding: 18px;
            border-radius: 18px;
            min-height: auto;
            margin-top: 20px;
        }}

        .st-key-about_box h3,
        .st-key-contact_box h3,
        .st-key-mission_box h3,
        .st-key-products_box h3 {{
            font-size: 24px !important;
        }}

        .st-key-about_box p,
        .st-key-about_box li,
        .st-key-contact_box p,
        .st-key-contact_box li,
        .st-key-mission_box p,
        .st-key-mission_box li,
        .st-key-products_box p,
        .st-key-products_box li {{
            font-size: 16px !important;
            line-height: 1.4 !important;
        }}

        .map-frame {{
            height: 200px;
        }}

        .premium-calculator-title {{
            font-size: 26px;
        }}

        .st-key-premium_calculator_box h2,
        .st-key-premium_calculator_box h3 {{
            font-size: 21px;
        }}

        .st-key-premium_calculator_box label,
        .st-key-premium_calculator_box p {{
            font-size: 15px;
        }}

        .st-key-premium_calculator_box [role="radiogroup"] {{
            flex-wrap: wrap;
            gap: 8px;
        }}

        div.stButton > button {{
            width: 100%;
            padding: 12px 16px;
        }}
    }}

    @media (max-width: 480px) {{
        .hero-title {{
            font-size: 30px;
        }}

        .hero-tagline {{
            font-size: 14px;
        }}

        .st-key-about_box p,
        .st-key-about_box li,
        .st-key-contact_box p,
        .st-key-contact_box li,
        .st-key-mission_box p,
        .st-key-mission_box li,
        .st-key-products_box p,
        .st-key-products_box li {{
            font-size: 15px !important;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# =========================
# HERO SECTION
# =========================

hero_logo_col, hero_text_col = st.columns([1, 4], vertical_alignment="center")

with hero_logo_col:
    with st.container(key="hero_logo"):
        st.image(SIDE_IMAGE_FILE, width=175)

        st.markdown(
            """
            <div class="logo-name">
                Adevali<br>Insurance
            </div>
            """,
            unsafe_allow_html=True
        )

with hero_text_col:
    with st.container(key="hero_text"):
        st.markdown(
            """
            <div class="hero-title">
                Smart Life<br>
                Protection Solutions
            </div>

            <div class="hero-tagline">
                SECURE LIFE, SECURE LOVE
            </div>
            """,
            unsafe_allow_html=True
        )


# =========================
# INFORMATION BOXES
# =========================

about_col, contact_col = st.columns([2.2, 1], gap="large")

with about_col:
    with st.container(key="about_box"):
        st.markdown(
            """
            ### About Us

            Adevali Insurance is a student-created insurance company designed to provide simple, affordable, and reliable life insurance protection. Our goal is to help customers feel safer about the future by offering insurance products that are easy to understand and suitable for different needs.

            Adevali Insurance was developed as a university project by four students from President University, namely:

            - Alexandra Mercy Christani (021202400008)
            - Defira Lubnaziza (021202400004)
            - Liska Desryani Purba (021202400027)
            - Valerion Theodore Chandratama (021202400017)

            Through this project, we aim to combine actuarial knowledge, financial planning, and customer-focused service to create an insurance product that is both useful for policyholders and sustainable for the company.
            """
        )

with contact_col:
    with st.container(key="contact_box"):
        st.markdown(
            """
            ### Contact Us

            ☏ 0878-9191-7372

            ✉︎ adevali.help@gmail.com

            𖤣 President University,  
            Jl. Ki Hajar Dewantara, Mekarmukti, Kec. Cikarang Utara, Kabupaten Bekasi, Jawa Barat 17530.
            """
        )

        st.markdown(
            """
            <iframe
                class="map-frame"
                src="https://www.google.com/maps?q=President%20University%20Jl.%20Ki%20Hajar%20Dewantara%20Kota%20Jababeka%20Cikarang%20Baru%20Bekasi&output=embed"
                allowfullscreen
                loading="lazy"
                scrolling="no">
            </iframe>
            """,
            unsafe_allow_html=True
        )


with st.container(key="mission_box"):
    st.markdown(
        """
        ### Our Mission

        To provide life insurance protection that is simple, affordable, and trustworthy. Adevali Insurance aims to help individuals and families prepare for unexpected risks by offering clear benefits, fair premiums, and reliable financial protection.

        We are committed to creating insurance products that are easy to understand, accessible to customers, and designed with care to support a safer and more secure future.
        """
    )


with st.container(key="products_box"):
    st.markdown(
        """
        ### Our Products

        Adevali Insurance offers three life insurance protection tiers designed to match different customer needs and budgets. Each tier provides protection for accident, illness, and other causes, with different benefit amounts depending on the selected package.

        **Tier 1: Premium Protection**  
        The strongest package offered by Adevali Insurance, providing the highest level of financial protection. This tier gives an accident benefit of Rp2,000,000,000, an illness benefit of Rp1,000,000,000 after the policy has been active for 5 years, and an other causes benefit of Rp1,500,000,000. It is suitable for customers who want maximum protection for themselves and their loved ones.

        **Tier 2: Standard Protection**  
        A balanced level of protection with higher benefits than Tier 3. This package is suitable for customers who want stronger financial security while still maintaining an affordable premium. This tier provides an accident benefit of Rp1,000,000,000, an illness benefit of Rp750,000,000 after the policy has been active for 5 years, and an other causes benefit of Rp950,000,000.

        **Tier 3: Basic Protection**  
        The most affordable package, suitable for customers who want essential protection at a lower premium. This tier provides an accident benefit of Rp500,000,000, an illness benefit of Rp300,000,000 after the policy has been active for 5 years, and an other causes benefit of Rp475,000,000.
        """
    )


# =========================
# PREMIUM CALCULATOR
# =========================

with st.container(border=True, key="premium_calculator_box"):

    st.markdown(
        """
        <div class="premium-calculator-title">
            Premium Calculator
        </div>
        """,
        unsafe_allow_html=True
    )

    st.subheader("Input Information")

    user_name = st.text_input("Enter your full name")

    gender = st.radio(
        "Choose your gender",
        ["male", "female"],
        horizontal=True
    )

    x = st.number_input(
        "Enter your age",
        min_value=0,
        max_value=120,
        value=25,
        step=1
    )

    st.subheader("Underwriting Information")

    smoker_answer = st.radio(
        "Smoker status",
        ["non-smoker", "smoker"],
        horizontal=True
    )

    if smoker_answer == "smoker":
        smoker_level = st.radio(
            "Smoker level",
            ["light", "heavy"],
            horizontal=True
        )
    else:
        smoker_level = "none"

    alcohol_answer = st.radio(
        "Alcohol status",
        ["non-drinker", "drinker"],
        horizontal=True
    )

    if alcohol_answer == "drinker":
        alcohol_level = st.radio(
            "Alcohol level",
            ["light", "moderate"],
            horizontal=True
        )
    else:
        alcohol_level = "none"

    dangerous_hobby_answer = st.radio(
        "Dangerous hobby",
        ["no", "yes"],
        horizontal=True
    )

    if dangerous_hobby_answer == "yes":
        dangerous_hobby_detail = st.text_input(
            "Please type your dangerous hobby"
        )
    else:
        dangerous_hobby_detail = "None"

    show_actuarial_calculation = st.toggle(
        "Show actuarial calculation",
        value=False
    )

    calculate_button = st.button("Calculate your Premium")

    if calculate_button:
        try:
            x = int(x)

            if gender == "male":
                selected_table = male_table
            else:
                selected_table = female_table

            underwriting_info = make_underwriting_info(
                smoker_level=smoker_level,
                alcohol_level=alcohol_level,
                dangerous_hobby=dangerous_hobby_answer
            )

            customer_summary = pd.DataFrame(
                [
                    {
                        "Name": user_name,
                        "Gender": gender,
                        "Age": x
                    }
                ]
            )

            st.subheader("Upcoming Policyholder Information")

            st.dataframe(
                customer_summary,
                use_container_width=True,
                hide_index=True
            )

            underwriting_summary = pd.DataFrame(
                [
                    {
                        "Smoker status": underwriting_info["smoker_level"],
                        "Alcohol status": underwriting_info["alcohol_level"],
                        "Dangerous hobby": underwriting_info["dangerous_hobby"],
                        "Dangerous hobby detail": dangerous_hobby_detail,
                        "Smoker loading": f"{underwriting_info['smoker_loading']:.0%}",
                        "Alcohol loading": f"{underwriting_info['alcohol_loading']:.0%}",
                        "Dangerous hobby loading": f"{underwriting_info['dangerous_hobby_loading']:.0%}",
                        "Total premium loading": f"{underwriting_info['total_loading']:.0%}",
                        "Underwriting factor": underwriting_info["underwriting_factor"]
                    }
                ]
            )

            st.subheader("Underwriting Summary")

            st.dataframe(
                underwriting_summary,
                use_container_width=True,
                hide_index=True
            )

            result_rows = []

            for tier_name, benefits in tiers.items():
                accident_bt = benefits["accident_bt"]
                other_bt = benefits["other_bt"]
                illness_bt = benefits["illness_bt"]

                apv_accident, apv_other, apv_illness, apv_total = calculate_apv_total_whole_life(
                    table=selected_table,
                    x=x,
                    accident_bt=accident_bt,
                    other_bt=other_bt,
                    illness_bt=illness_bt
                )

                annuity_annual, annuity_monthly = calculate_annuity_whole_life(
                    table=selected_table,
                    x=x
                )

                (
                    net_yearly_premium,
                    net_monthly_premium,
                    gross_yearly_premium,
                    gross_monthly_premium,
                    gross_annualized_monthly_premium,
                    apv_settlement
                ) = calculate_net_and_gross_premiums(
                    apv_accident=apv_accident,
                    apv_other=apv_other,
                    apv_illness=apv_illness,
                    apv_total=apv_total,
                    annuity_annual=annuity_annual,
                    annuity_monthly=annuity_monthly,
                    accident_bt=accident_bt,
                    other_bt=other_bt,
                    illness_bt=illness_bt
                )

                (
                    adjusted_gross_yearly_premium,
                    adjusted_gross_monthly_premium,
                    adjusted_gross_annualized_monthly_premium
                ) = apply_underwriting_loading(
                    gross_yearly_premium=gross_yearly_premium,
                    gross_monthly_premium=gross_monthly_premium,
                    gross_annualized_monthly_premium=gross_annualized_monthly_premium,
                    underwriting_info=underwriting_info
                )

                result_rows.append(
                    {
                        "Tier": tier_name,
                        "Accident Benefit": rupiah_no_decimal(accident_bt),
                        "Illness Benefit": rupiah_no_decimal(illness_bt),
                        "Other Causes Benefit": rupiah_no_decimal(other_bt),
                        "Gross Monthly Premium": rupiah(adjusted_gross_monthly_premium),
                        "Gross Yearly Premium": rupiah(adjusted_gross_yearly_premium)
                    }
                )

            result_table = pd.DataFrame(result_rows)

            st.subheader("Premium (Level) Results for a Yearly and Monthly Payment:")

            st.dataframe(
                result_table,
                use_container_width=True,
                hide_index=True
            )

        except Exception as error:
            st.error(f"Calculation error: {error}")
