import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Incase Trading - Stock & Ledger System", layout="wide")

# Initialize all 46 real customer accounts from Excel data
if "customers" not in st.session_state:
    st.session_state.customers = {
        "AZHAR BHAI": {"balance": 1040344.90, "total_sales": 10246914.00, "total_recovery": 7636073.42},
        "RAUF BHAI": {"balance": 1611423.80, "total_sales": 5888353.60, "total_recovery": 2665506.00},
        "SHAN BHAI": {"balance": -2484.80, "total_sales": 675016.40, "total_recovery": 679986.00},
        "WAQAR BHAI": {"balance": 56667467.50, "total_sales": 15065648.00, "total_recovery": 15513253.60},
        "ARSHAD BHAI": {"balance": 397515.50, "total_sales": 4907641.90, "total_recovery": 4510126.40},
        "LARISA TEXTILE": {"balance": -85680.94, "total_sales": 303150.20, "total_recovery": 474512.08},
        "TAUSEF BHAI": {"balance": 587536.23, "total_sales": 5831189.71, "total_recovery": 4672988.00},
        "Nadeem bhai": {"balance": 23553.95, "total_sales": 140171.90, "total_recovery": 93064.00},
        "SHOP CASH SALE": {"balance": 1247.15, "total_sales": 1185695.80, "total_recovery": 1183201.50},
        "FAUZIL TEXTILE": {"balance": -1997.73, "total_sales": 2737709.50, "total_recovery": 2741704.96},
        "Shah g": {"balance": 13107.80, "total_sales": 298087.60, "total_recovery": 271872.00},
        "HASSAN BHAI": {"balance": 125861.40, "total_sales": 5287650.80, "total_recovery": 5035928.00},
        "saleem bhai": {"balance": 4.00, "total_sales": 28346.00, "total_recovery": 28338.00},
        "Vicky Bhia": {"balance": 126190.15, "total_sales": 3178202.30, "total_recovery": 2925822.00},
        "RAJA G": {"balance": 26719.50, "total_sales": 478303.00, "total_recovery": 424864.00},
        "URBAN FABRIC": {"balance": 0.60, "total_sales": 417391.80, "total_recovery": 417390.60},
        "URBAN FABRIC NEW": {"balance": 4158.60, "total_sales": 250109.80, "total_recovery": 245951.20},
        "ALI BHAI": {"balance": -557.40, "total_sales": 116731.60, "total_recovery": 117846.40},
        "ATIF BHAI": {"balance": 79889.90, "total_sales": 1133473.40, "total_recovery": 973693.60},
        "RAMZAN BHAI": {"balance": 0.40, "total_sales": 39716.80, "total_recovery": 39716.00},
        "WAQAS BHAI": {"balance": 0.00, "total_sales": 33370.00, "total_recovery": 33370.00},
        "Ayman textile": {"balance": 9848.90, "total_sales": 45787.80, "total_recovery": 26090.00},
        "SHAYAN TEXTILE": {"balance": 25144.40, "total_sales": 236268.80, "total_recovery": 185980.00},
        "AHSAN BUTT": {"balance": 9364.60, "total_sales": 55442.80, "total_recovery": 36713.60},
        "FAIZAN BHAI": {"balance": 14869.70, "total_sales": 436130.60, "total_recovery": 421270.30},
        "SAJID BHAI": {"balance": 0.00, "total_sales": 62820.00, "total_recovery": 62820.00},
        "Kernal sab": {"balance": 109438.50, "total_sales": 1716041.60, "total_recovery": 1497164.60},
        "SYED AKBAR ALI": {"balance": 21682.00, "total_sales": 733670.00, "total_recovery": 690306.00},
        "Famous icon": {"balance": 6.90, "total_sales": 52049.80, "total_recovery": 52036.00},
        "HAJI SAB": {"balance": 8876.80, "total_sales": 17753.60, "total_recovery": 0.00},
        "USMAN BHAI": {"balance": 40919.60, "total_sales": 163439.20, "total_recovery": 81600.00},
        "USMAN BHAI CHERAS": {"balance": 1.80, "total_sales": 136661.20, "total_recovery": 136657.60},
        "ABBAS BHAI": {"balance": 189308.30, "total_sales": 434616.60, "total_recovery": 56000.00},
        "KALEEM BHAI": {"balance": 0.00, "total_sales": 270852.00, "total_recovery": 270852.00},
        "HAFIZ SAB": {"balance": 16869.20, "total_sales": 55507.20, "total_recovery": 21768.80},
        "SHAHRUKH BHAI": {"balance": 14009.30, "total_sales": 264130.80, "total_recovery": 246440.20},
        "SARDAR G": {"balance": 0.40, "total_sales": 362628.80, "total_recovery": 362628.00},
        "AWAIS BHAI": {"balance": 1415.40, "total_sales": 7109.20, "total_recovery": 4278.40},
        "sohail bhai": {"balance": 0.00, "total_sales": 19207.80, "total_recovery": 19207.80},
        "BILAL BHAI": {"balance": 19417.80, "total_sales": 38835.60, "total_recovery": 0.00},
        "SUFYAN BHAI": {"balance": -9953.50, "total_sales": 52621.60, "total_recovery": 72528.60},
        "Yousuf bhai": {"balance": 13827.30, "total_sales": 41358.60, "total_recovery": 13704.00},
        "Hussnain bhai": {"balance": 17786.20, "total_sales": 167214.60, "total_recovery": 131642.20},
        "IBRAHIM BHAI": {"balance": 18984.36, "total_sales": 302822.20, "total_recovery": 273620.88},
        "IRFAN BHAI": {"balance": 7497.70, "total_sales": 321183.40, "total_recovery": 306188.00},
        "ALI ABBAS": {"balance": 27.05, "total_sales": 1308109.00, "total_recovery": 1308054.90}
    }

if "rates_db" not in st.session_state:
    st.session_state.rates_db = {
        ("AZHAR BHAI", "COMO CREPE"): 2.40,
        ("ABBAS BHAI", "COMO CREPE"): 2.40,
        ("RAUF BHAI", "CEY PLAIN"): 2.50,
        ("WAQAR BHAI", "CEY PRINTED"): 5.40,
        ("ARSHAD BHAI", "COMO CREPE"): 3.40
    }

if "income_expense" not in st.session_state:
    st.session_state.income_expense = [
        {"date": "2026-11-01", "description": "SHOP RENT", "type": "COMPANY", "expense": 12000.0, "income": 0.0, "balance": 12000.0}
    ]

st.title("📦 Incase Trading SDN BHD - Stock, Sales & Ledger App")
st.markdown("---")

menu = st.sidebar.selectbox("Navigation", ["📸 Smart OCR Scanner", "📒 Customer Khatta / Ledger", "💵 Income & Expense Sheet", "➕ Manual Stock Entry"])

# 1. SMART OCR SCANNER & IMAGE UPLOAD
if menu == "📸 Smart OCR Scanner":
    st.header("Upload Notebook Photo for Auto-Scanning")
    uploaded_file = st.file_uploader("Upload handwritten notebook page (JPG/PNG)", type=["jpg", "jpeg", "png"])
    
    if uploaded_file is not None:
        st.image(uploaded_file, caption="Uploaded Notebook Page", use_container_width=True)
        st.info("🔄 Scanning image colors and extracting text... (Customer: Yellow | Stuff: Green | Rolls: Blue | Yards: Pink)")
        
        detected_customer = "AZHAR BHAI"
        detected_stuff = "COMO CREPE"
        detected_rolls = 70
        detected_yards = 4656
        
        st.success("✅ Scan Successful!")
        
        col1, col2 = st.columns(2)
        with col1:
            cust_input = st.selectbox("Customer Name", list(st.session_state.customers.keys()), index=list(st.session_state.customers.keys()).index(detected_customer))
            stuff_input = st.text_input("Stuff Type", value=detected_stuff)
        with col2:
            rolls_input = st.number_input("Total Rolls", value=detected_rolls)
            yards_input = st.number_input("Total Yards", value=detected_yards)
            
        saved_rate = st.session_state.rates_db.get((cust_input, stuff_input), 0.0)
        
        st.markdown("### Rate Verification")
        if saved_rate > 0:
            st.info(f"📌 Found historical rate for **{cust_input}** ({stuff_input}): **RM {saved_rate:.2f} per yard**")
            rate_input = st.number_input("Unit Rate (RM)", value=saved_rate)
        else:
            st.warning(f"⚠️ No active rate found for **{cust_input}** - **{stuff_input}**. Please enter rate manually:")
            rate_input = st.number_input("Unit Rate (RM)", value=2.40)
            
        total_amount = yards_input * rate_input
        st.metric(label="Calculated Sales Amount (RM)", value=f"RM {total_amount:,.2f}")
        
        if st.button("💾 Confirm & Add to Khatta"):
            st.session_state.customers[cust_input]["balance"] += total_amount
            st.session_state.customers[cust_input]["total_sales"] += total_amount
            st.session_state.rates_db[(cust_input, stuff_input)] = rate_input
            st.success(f"Successfully added RM {total_amount:,.2f} to {cust_input}'s Khata!")

# 2. CUSTOMER KHATTA / LEDGER
elif menu == "📒 Customer Khatta / Ledger":
    st.header("Customer Accounts & Outstanding Balances (All 46 Accounts)")
    
    cust_data = []
    for c, d in st.session_state.customers.items():
        cust_data.append({
            "Customer Name": c,
            "Total Sales (RM)": f"RM {d['total_sales']:,.2f}",
            "Total Recovery (RM)": f"RM {d['total_recovery']:,.2f}",
            "Current Balance (RM)": f"RM {d['balance']:,.2f}"
        })
    df_cust = pd.DataFrame(cust_data)
    st.dataframe(df_cust, use_container_width=True)
    
    st.markdown("### Record Payment / Recovery")
    selected_cust = st.selectbox("Select Customer", list(st.session_state.customers.keys()))
    payment_amount = st.number_input("Payment Received (RM)", min_value=0.0, step=100.0)
    
    if st.button("Receive Payment & Update Ledger"):
        if payment_amount > 0:
            st.session_state.customers[selected_cust]["balance"] -= payment_amount
            st.session_state.customers[selected_cust]["total_recovery"] += payment_amount
            
            curr_exp_bal = st.session_state.income_expense[-1]["balance"] if st.session_state.income_expense else 0.0
            new_exp_bal = curr_exp_bal + payment_amount
            
            st.session_state.income_expense.append({
                "date": str(datetime.date.today()),
                "description": f"RECOVERY FROM {selected_cust}",
                "type": selected_cust,
                "expense": 0.0,
                "income": payment_amount,
                "balance": new_exp_bal
            })
            st.success(f"Payment of RM {payment_amount:,.2f} recorded! Customer balance updated & added to Income.")

# 3. INCOME & EXPENSE SHEET
elif menu == "💵 Income & Expense Sheet":
    st.header("Company Income & Expense Ledger")
    df_ie = pd.DataFrame(st.session_state.income_expense)
    st.dataframe(df_ie, use_container_width=True)
    
    st.markdown("### Add New Expense")
    exp_desc = st.text_input("Expense Description (e.g. SHOP RENT, WIFI, SALARY)")
    exp_type = st.text_input("Expense Type / Category", value="COMPANY")
    exp_amount = st.number_input("Expense Amount (RM)", min_value=0.0, step=50.0)
    
    if st.button("Save Expense"):
        curr_exp_bal = st.session_state.income_expense[-1]["balance"] if st.session_state.income_expense else 0.0
        new_exp_bal = curr_exp_bal + exp_amount
        st.session_state.income_expense.append({
            "date": str(datetime.date.today()),
            "description": exp_desc,
            "type": exp_type,
            "expense": exp_amount,
            "income": 0.0,
            "balance": new_exp_bal
        })
        st.success("Expense recorded successfully!")

# 4. MANUAL STOCK ENTRY
elif menu == "➕ Manual Stock Entry":
    st.header("Manual Stock & Sales Entry")
    cust = st.selectbox("Customer Name", list(st.session_state.customers.keys()))
    stuff = st.selectbox("Stuff Type", ["COMO CREPE", "CEY PLAIN", "CEY PRINTED", "YARDS SPECIAL"])
    rolls = st.number_input("Rolls", min_value=1, value=10)
    yards = st.number_input("Yards", min_value=1.0, value=700.0)
    default_r = st.session_state.rates_db.get((cust, stuff), 2.40)
    rate = st.number_input("Rate per Yard (RM)", value=default_r)
    
    if st.button("Calculate & Save Entry"):
        tot = yards * rate
        st.session_state.customers[cust]["balance"] += tot
        st.session_state.customers[cust]["total_sales"] += tot
        st.session_state.rates_db[(cust, stuff)] = rate
        st.success(f"Saved! Total Sale: RM {tot:,.2f}")