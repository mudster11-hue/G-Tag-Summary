import io
import os
import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================
# PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="State of Utah GA4 Governance Dashboard",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main { padding-top: 1rem; }
    .stMetric {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E88E5;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    </style>
""",
    unsafe_allow_html=True,
)


def convert_df_to_excel(df_to_convert):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_convert.to_excel(writer, index=False, sheet_name="Exported_Data")
    return output.getvalue()


# ==========================================
# DATA LOADING FUNCTION (STRICT MASTER SHEET)
# ==========================================
@st.cache_data
def load_data():
    file_name = "websites8.xlsx"
    
    if not os.path.exists(file_name):
        if os.path.exists("websites8.csv"):
            file_name = "websites8.csv"
        else:
            raise FileNotFoundError("Could not find 'websites8.xlsx' or 'websites8.csv'.")

    # Load ONLY Master Sheet
    if file_name.endswith(".csv"):
        df = pd.read_csv(file_name)
        sheet_used = "Master CSV"
    else:
        xl = pd.ExcelFile(file_name)
        sheet_names = xl.sheet_names
        # Lock strictly onto Master Sheet tab
        master_sheet_name = next((s for s in sheet_names if "master" in s.lower()), sheet_names[0])
        df = pd.read_excel(file_name, sheet_name=master_sheet_name)
        sheet_used = master_sheet_name

    df.columns = df.columns.astype(str).str.strip()

    # Dynamic Column Resolver
    def find_col(candidates, default):
        for c in df.columns:
            cleaned_c = c.lower().replace(" ", "").replace("_", "").replace("?", "")
            for x in candidates:
                cleaned_x = x.lower().replace(" ", "").replace("_", "").replace("?", "")
                if cleaned_c == cleaned_x:
                    return c
        return default

    agency_col = find_col(["Agency", "Department"], "Agency")
    url_col = find_col(["URL Path", "Path", "URL"], "URL Path")
    host_col = find_col(["Host Name", "Hostname", "Host"], "Host Name")
    gtag_col = find_col(["G-Tag", "GTag", "Measurement_ID"], "G-Tag")
    gtm_col = find_col(["GTM", "GTM Container", "GTM ID"], "GTM")
    state_acct_col = find_col(["State of Utah GA Account Y/N", "State GA Account", "GA Account Y/N"], "State of Utah GA Account Y/N")
    prop_col = find_col(["GA property Name", "GAProperty Name", "GA360_Property_Name"], "GA property Name")
    status_col = find_col(["Status", "HTTP Status"], "Status")

    df["Agency"] = df[agency_col] if agency_col in df.columns else "UNKNOWN"
    df["URL Path"] = df[url_col] if url_col in df.columns else "/"
    df["Host Name"] = df[host_col] if host_col in df.columns else "Unknown Host"
    df["G-Tag"] = df[gtag_col] if gtag_col in df.columns else "N/A"
    df["GTM"] = df[gtm_col] if gtm_col in df.columns else "N/A"
    df["State of Utah GA Account Y/N"] = df[state_acct_col] if state_acct_col in df.columns else "N"
    df["GA property Name"] = df[prop_col] if prop_col in df.columns else "N/A"
    df["Status"] = df[status_col] if status_col in df.columns else "Active"

    # Normalize values safely
    for c in ["Agency", "URL Path", "Host Name", "G-Tag", "GTM", "State of Utah GA Account Y/N", "GA property Name", "Status"]:
        df[c] = df[c].fillna("N/A").astype(str).str.strip()
        df[c] = df[c].replace("", "N/A")

    # =========================================================
    # ROBUST TAG DETECTION LOGIC (HANDLES TIMEOUT / LOAD FAILED)
    # =========================================================
    invalid_tag_values = [
        "N/A", "NAN", "NONE", "NULL", "NO TAG", "0", "", "UNKNOWN", 
        "NOT DETECTED", "NONE FOUNDTIMEOUT / LOAD FAILED", "NONE FOUND", 
        "TIMEOUT / LOAD FAILED", "LOAD FAILED", "TIMEOUT"
    ]

    # 1. Missing G-Tag: Value matches failed strings or does NOT start with 'G-'
    def check_missing_gtag(val):
        clean_val = str(val).strip().upper()
        if clean_val in invalid_tag_values or any(err in clean_val for err in ["TIMEOUT", "LOAD FAILED", "NONE FOUND"]):
            return True
        return not clean_val.startswith("G-")

    # 2. Missing GTM: Value matches failed strings or does NOT start with 'GTM-'
    def check_missing_gtm(val):
        clean_val = str(val).strip().upper()
        if clean_val in invalid_tag_values or any(err in clean_val for err in ["TIMEOUT", "LOAD FAILED", "NONE FOUND"]):
            return True
        return not clean_val.startswith("GTM-")

    # 3. GA360 Onboarding Candidate: Active G-Tag present AND State GA Account = N/NO
    def check_onboarding(row):
        gtag = str(row["G-Tag"]).strip().upper()
        state_acct = str(row["State of Utah GA Account Y/N"]).strip().upper()
        
        is_missing_tag = check_missing_gtag(gtag)
        has_active_gtag = (not is_missing_tag) and gtag.startswith("G-")
        is_not_state_acct = state_acct in ["N", "NO", "FALSE", "0", "N/A"]
        
        return has_active_gtag and is_not_state_acct

    df["Missing_GTag"] = df["G-Tag"].apply(check_missing_gtag)
    df["Missing_GTM"] = df["GTM"].apply(check_missing_gtm)
    df["Needs_GA360_Onboarding"] = df.apply(check_onboarding, axis=1)

    return df, file_name, sheet_used

try:
    df, loaded_file, loaded_sheet = load_data()
except Exception as e:
    st.error(f"🚨 Could not load data: {e}")
    st.stop()

# ==========================================
# SIDEBAR FILTERS
# ==========================================
st.sidebar.title("🎛️ Governance Filters")
st.sidebar.caption(f"📁 Source: `{loaded_file}` [{loaded_sheet}]")

agencies = ["All Agencies"] + sorted([a for a in df["Agency"].unique() if a and a != "N/A"])
selected_agency = st.sidebar.selectbox("Filter by Agency:", options=agencies)

df_view = df.copy()
if selected_agency != "All Agencies":
    df_view = df_view[df_view["Agency"] == selected_agency]

# ==========================================
# HEADER & KPI METRICS (AT A GLANCE)
# ==========================================
st.title("🏛️ State of Utah | Master Web Inventory Audit")
st.markdown("Operational tracking focusing strictly on records inside the **Master Sheet**.")

st.divider()

col1, col2, col3, col4 = st.columns(4)

total_assets = len(df_view)
onboarding_count = len(df_view[df_view["Needs_GA360_Onboarding"]])
missing_gtag_count = len(df_view[df_view["Missing_GTag"]])
missing_gtm_count = len(df_view[df_view["Missing_GTM"]])

col1.metric("Total Web Assets", f"{total_assets:,}")
col2.metric("GA360 Onboarding Queue", f"{onboarding_count:,}", delta="Active G-Tags Outside State GA Account", delta_color="normal")
col3.metric("Assets Missing G-Tags", f"{missing_gtag_count:,}", delta="Needs Analytics Tag", delta_color="inverse")
col4.metric("Assets Missing GTM", f"{missing_gtm_count:,}", delta="Needs Container Tag", delta_color="inverse")

st.divider()

# ==========================================
# WORK QUEUE TABS (WHAT NEEDS TO BE DONE)
# ==========================================
st.header("📌 Work Queue & Action Items")

tab1, tab2, tab3 = st.tabs([
    f"⚠️ Missing G-Tags ({missing_gtag_count})",
    f"🚀 GA360 Onboarding Candidates ({onboarding_count})",
    f"📦 Missing GTM ({missing_gtm_count})"
])

with tab1:
    st.subheader("Websites Missing Google Analytics Tracking (G-Tags)")
    st.markdown("Active web paths where no valid `G-XXXXXXXXXX` tag was found (includes timeouts and load failures).")
    df_no_gtag = df_view[df_view["Missing_GTag"]]
    if not df_no_gtag.empty:
        cols = ["Agency", "Host Name", "URL Path", "G-Tag", "GTM", "Status"]
        st.dataframe(df_no_gtag[cols], use_container_width=True, hide_index=True)
        st.download_button("📥 Export Missing G-Tags List to Excel", convert_df_to_excel(df_no_gtag[cols]), "Missing_GTags_Queue.xlsx")
    else:
        st.success("All sites have valid G-Tags assigned!")

with tab2:
    st.subheader("Websites Firing Active Tags Outside Official GA360 Account")
    st.markdown("These sites have active `G-Tags` but are marked **`State GA Account = N`** in Master Sheet.")
    df_onboard = df_view[df_view["Needs_GA360_Onboarding"]]
    if not df_onboard.empty:
        cols = ["Agency", "Host Name", "URL Path", "G-Tag", "State of Utah GA Account Y/N", "GA property Name"]
        st.dataframe(df_onboard[cols], use_container_width=True, hide_index=True)
        st.download_button("📥 Export Onboarding Queue to Excel", convert_df_to_excel(df_onboard[cols]), "GA360_Onboarding_Queue.xlsx")
    else:
        st.success("No pending onboarding items for this selection!")

with tab3:
    st.subheader("Websites Missing Google Tag Manager (GTM)")
    st.markdown("Active web paths where no valid `GTM-XXXXXXX` container ID was found.")
    df_no_gtm = df_view[df_view["Missing_GTM"]]
    if not df_no_gtm.empty:
        cols = ["Agency", "Host Name", "URL Path", "GTM", "G-Tag", "Status"]
        st.dataframe(df_no_gtm[cols], use_container_width=True, hide_index=True)
        st.download_button("📥 Export Missing GTM List to Excel", convert_df_to_excel(df_no_gtm[cols]), "Missing_GTM_Queue.xlsx")
    else:
        st.success("All sites have GTM containers assigned!")

st.divider()

# ==========================================
# AGENCY FOOTPRINT COMPARISON
# ==========================================
st.header("📊 Agency Portfolio Footprints")

agency_summary = (
    df_view[df_view["Agency"] != "N/A"]
    .groupby("Agency")
    .agg(
        Total_Assets=("URL Path", "count"),
        Missing_GTags=("Missing_GTag", "sum"),
        Onboarding_Needed=("Needs_GA360_Onboarding", "sum"),
        Missing_GTM=("Missing_GTM", "sum"),
    )
    .reset_index()
    .sort_values(by="Total_Assets", ascending=False)
)

fig = px.bar(
    agency_summary.head(15),
    x="Agency",
    y=["Total_Assets", "Missing_GTags", "Onboarding_Needed"],
    barmode="group",
    title="Top Agencies: Total Web Assets vs. Missing G-Tags & Onboarding Tasks",
    labels={"value": "Count", "variable": "Metric"}
)
st.plotly_chart(fig, use_container_width=True)

st.divider()

# ==========================================
# COMPLETE MASTER DATA VIEW
# ==========================================
st.header("📑 Complete Master Sheet Inventory Lookup")
search = st.text_input("🔍 Search Master Inventory (Host, URL, Agency, or Tag):")

df_master = df_view.copy()
if search.strip():
    df_master = df_master[df_master.apply(lambda r: r.astype(str).str.contains(search.strip(), case=False).any(), axis=1)]

clean_cols = [c for c in df_master.columns if not c.startswith("Needs_") and not c.startswith("Missing_")]
st.dataframe(df_master[clean_cols], use_container_width=True, hide_index=True)
st.download_button("📥 Download Filtered Master Inventory", convert_df_to_excel(df_master[clean_cols]), "Utah_Master_Web_Inventory.xlsx")