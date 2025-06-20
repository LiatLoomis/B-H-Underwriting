import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

# Page setup
st.set_page_config(page_title="BH Underwriting AI", layout="centered")
st.title("🏢 BH Property Underwriting Tool")

# 🔐 Password protection
if "auth" not in st.session_state:
    pwd = st.text_input("Enter password to use the app:", type="password")
    if pwd != "1234":
        st.warning("Incorrect or missing password.")
        st.stop()
    else:
        st.session_state.auth = True

# Upload rent roll
st.header("📁 Upload Rent Roll (.xlsx)")
rent_roll_file = st.file_uploader("Upload Rent Roll", type=["xlsx"], help="Excel format only")

# Input fields
st.header("🏦 Property & Loan Info")
price = st.number_input("Property Purchase Price ($)", value=1000000)
loan = st.number_input("Loan Amount ($)", value=700000)
rate = st.number_input("Interest Rate (%)", value=6.0)
term = st.number_input("Loan Term (Years)", value=30)
vacancy = st.number_input("Vacancy Rate (%)", value=5.0)
lease_type = st.selectbox("Lease Type", ["Gross", "NNN"])

# Upload BH template
template_file = st.file_uploader("Upload BH Excel Template (.xlsx)", type=["xlsx"], help="Your blank underwriting model")

# Run button
run = st.button("Run Underwriting & Download")

# Main logic
if run:
    if not rent_roll_file or not template_file:
        st.error("Please upload both the rent roll and the template.")
    else:
        # Read rent roll
        df = pd.read_excel(rent_roll_file)

        # Try to detect rent column
        try:
            rent_column = df["Monthly Rent"]
        except KeyError:
            rent_column = df.iloc[:, -1]  # fallback: last column

        monthly_rent_total = rent_column.sum()
        annual_income = monthly_rent_total * 12 * (1 - vacancy / 100)

        # Expenses
        if lease_type == "NNN":
            expenses = 0
        else:
            expenses = (
                0.012 * price +  # taxes
                8000 +           # insurance
                10000 +          # maintenance
                6000             # utilities
            )

        noi = annual_income - expenses

        # Loan calculations
        r = rate / 100 / 12
        n = term * 12
        pmt = loan * r / (1 - (1 + r) ** -n)
        annual_debt = pmt * 12

        # Returns
        taxable = noi - annual_debt
        after_tax = taxable * (1 - 0.35)
        coc = after_tax / (price - loan)
        cap_rate = noi / price

        # Load template and fill in values
        wb = load_workbook(template_file)
        ws = wb.active

        ws["C24"] = annual_income
        ws["C31"] = expenses
        ws["C43"] = noi
        ws["C55"] = taxable
        ws["C56"] = after_tax
        ws["C61"] = coc
        ws["C62"] = cap_rate

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Underwriting completed!")
        st.download_button(
            label="📥 Download Completed Excel",
            data=output,
            file_name="BH_Underwriting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

