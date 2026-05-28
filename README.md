# Adevali Insurance Streamlit Website

Adevali Insurance is a Streamlit-based life insurance website created for an actuarial science university project.  
The website includes company profile information, product tier descriptions, underwriting questions, and an automatic premium calculator.

## Project Features

- Company profile section
- Mission section
- Product tier explanation
- Premium calculator
- Gender, smoker, alcohol, and dangerous hobby underwriting inputs
- Premium loading calculation
- Monthly and yearly gross premium results
- Mortality table calculation using Excel data

## Repository Files

Make sure your GitHub repository contains these files:

```text
Adevali-Insurance/
│
├── streamlit_app.py
├── requirements.txt
├── Mortality_Table.xlsx
├── Copy of Flyer Life Insurance.png
├── logow.png
├── README.md
└── .gitignore
```

## Important File Names

The image and Excel file names must match the names used inside the Python code exactly:

```python
BACKGROUND_FILE = APP_DIR / "Copy of Flyer Life Insurance.png"
SIDE_IMAGE_FILE = APP_DIR / "logow.png"
MORTALITY_FILE = APP_DIR / "Mortality_Table.xlsx"
```

GitHub and Streamlit Cloud are case-sensitive, so `Mortality_Table.xlsx` is not the same as `mortality_table.xlsx`.

## Installation

Install the required packages using:

```bash
pip install -r requirements.txt
```

## Run the App Locally

Run the Streamlit app with:

```bash
streamlit run streamlit_app.py
```

## Deploy to Streamlit Community Cloud

1. Upload all project files to a GitHub repository.
2. Go to Streamlit Community Cloud.
3. Connect your GitHub account.
4. Choose this repository.
5. Set the main file path as:

```text
streamlit_app.py
```

6. Click Deploy.

## Notes

Do not upload your `venv` folder to GitHub.  
The `requirements.txt` file is enough for Streamlit Cloud to install the needed packages.
