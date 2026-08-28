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
    .stAlert { border-radius: 8px; }
    </style>
""",
    unsafe_allow_html=True,
)


# Helper function to convert DataFrames to Excel bytes for download
def convert_df_to_excel(df_to_convert):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_to_convert.to_excel(writer, index=False, sheet_name="Exported_Data")
    return output.getvalue()


# ==========================================
# DATA LOADING FUNCTION
# ==========================================
@st.cache_data
def load_and_prep_data():
    primary_file = "Masterspreadsheet.xlsx"

    if not os.path.exists(primary_file):
        possible_files = [
            f
            for f in os.listdir(".")
            if f.endswith(".xlsx") and "master" in f.lower()
        ]
        if possible_files:
            primary_file = possible_files[0]
        else:
            raise FileNotFoundError(
                f"Could not find '{primary_file}' in the current folder."
            )

    xl = pd.ExcelFile(primary_file)
    sheet_names = xl.sheet_names

    main_sheet = sheet_names[0]
    for s in sheet_names:
        if "master" in s.lower() or "cleaned" in s.lower():
            main_sheet = s
            break

    df = pd.read_excel(primary_file, sheet_name=main_sheet)
    df.columns = df.columns.astype(str).str.strip()

    # Dynamic Column Resolver
    def find_col(candidates, default):
        for c in df.columns:
            if c.lower().replace(" ", "").replace("_", "") in [
                x.lower().replace(" ", "").replace("_", "") for x in candidates
            ]:
                return c
        return default

    host_col = find_col(["Host Name", "Hostname", "Host"], "Host Name")
    ga_prop_col = find_col(
        ["GA property Name", "GAProperty Name", "GA360_Property_Name"],
        "GA property Name",
    )
    gtag_col = find_col(["G-Tag", "GTag", "Measurement_ID"], "G-Tag")
    agency_col = find_col(["Agency", "Department"], "Agency")
    url_col = find_col(["URL Path", "Path", "URL"], "URL Path")
    status_col = find_col(["Status", "HTTP Status"], "Status")

    # Map standardized columns
    df["Host Name"] = df[host_col] if host_col in df.columns else "Unknown Host"
    df["GA property Name"] = (
        df[ga_prop_col] if ga_prop_col in df.columns else "Unmapped"
    )
    df["G-Tag"] = df[gtag_col] if gtag_col in df.columns else "No Tag"
    df["Agency"] = df[agency_col] if agency_col in df.columns else "Unknown Agency"
    df["URL Path"] = df[url_col] if url_col in df.columns else "/"
    df["Status"] = df[status_col] if status_col in df.columns else "OK"

    # Fill NA values safely
    for col in ["Host Name", "GA property Name", "G-Tag", "Agency", "URL Path", "Status"]:
        df[col] = df[col].fillna("N/A").astype(str).str.strip()
        df[col] = df[col].replace("", "N/A")

    # Identify broken link / error rows
    error_keywords = [
        "404",
        "500",
        "error",
        "broken",
        "403",
        "failed",
        "timed out",
        "400",
    ]

    def check_error(row):
        status_val = str(row.get("Status", "")).lower()
        url_val = str(row.get("URL Path", "")).lower()
        return any(
            err in status_val or err in url_val for err in error_keywords
        )

    df["Is_Broken_Or_Error"] = df.apply(check_error, axis=1)

    # Detect multi-tag hostnames
    host_tag_counts = (
        df.groupby("Host Name")["G-Tag"]
        .apply(lambda x: len(set(g for g in x if g not in ["N/A", ""])))
        .reset_index(name="Distinct_G_Tags")
    )
    df = df.merge(host_tag_counts, on="Host Name", how="left")
    df["Has_Multiple_GTags"] = df["Distinct_G_Tags"] > 1

    return df, primary_file, main_sheet


# Load Dataset safely
try:
    df_raw, loaded_file, loaded_sheet = load_and_prep_data()
except Exception as e:
    st.error(f"🚨 Error loading dataset: {e}")
    st.stop()

# ==========================================
# SIDEBAR FILTERS
# ==========================================
st.sidebar.title("🎛️ Dashboard Filters")
st.sidebar.caption(f"📁 Loaded: `{loaded_file}` [{loaded_sheet}]")

agencies = sorted([a for a in df_raw["Agency"].unique() if a and a != "N/A"])
selected_agencies = st.sidebar.multiselect("Filter by Agency:", options=agencies)

st.sidebar.markdown("---")
st.sidebar.subheader("Special Audit Toggles")

broken_only = st.sidebar.checkbox(
    "🚨 Show Only Broken Links / Errors",
    value=False,
    help="Isolates rows with 404, 500, or HTTP error statuses.",
)

candidates_only = st.sidebar.checkbox(
    "🟡 Show Only Onboarding Candidates",
    value=False,
    help="Show only rows where tags fire outside the state account.",
)

df_filtered = df_raw.copy()

if selected_agencies:
    df_filtered = df_filtered[df_filtered["Agency"].isin(selected_agencies)]

if broken_only:
    df_filtered = df_filtered[df_filtered["Is_Broken_Or_Error"] == True]

if candidates_only:
    df_filtered = df_filtered[
        df_filtered["GA property Name"].str.contains(
            "Tag Firing but Not Found", case=False, na=False
        )
    ]

# ==========================================
# HEADER & EXECUTIVE SUMMARY METRICS
# ==========================================
st.title("🏛️ State of Utah | GA4 Analytics Governance Dashboard")
st.markdown(
    "Central monitoring for State of Utah web analytics properties, account mappings, onboarding candidates, and site health."
)

st.markdown("---")

col1, col2, col3, col4 = st.columns(4)

total_paths = len(df_filtered)
unique_hosts_count = df_filtered["Host Name"].nunique()
broken_count = len(df_filtered[df_filtered["Is_Broken_Or_Error"] == True])
candidates_count = len(
    df_filtered[
        df_filtered["GA property Name"].str.contains(
            "Tag Firing but Not Found", case=False, na=False
        )
    ]
)

col1.metric("Total Web Paths", f"{total_paths:,}")
col2.metric("Unique Host Names", f"{unique_hosts_count:,}")
col3.metric(
    "Broken Links / Errors",
    f"{broken_count:,}",
    delta="Review Needed" if broken_count > 0 else "All Healthy",
    delta_color="inverse",
)
col4.metric(
    "GA360 Onboarding Candidates",
    f"{candidates_count:,}",
    delta="Ready to Add",
    delta_color="normal",
)

st.markdown("---")

# ==========================================
# SECTION 1: GA PROPERTY NAME, HOST NAME & G-TAG LOOKUP
# ==========================================
st.header("1. 🔗 GA Property Name, Host Name & Correlated G-Tag Lookup")
st.info(
    "**What you are looking at:** This lookup table correlates **Host Names** with their assigned **GA Property Name** "
    "and state-registered **Measurement_ID** (G-Tag) from our **State Profile on Google Analytics**. "
    "Use the drop-downs or search bar below to filter results."
)

mapping_3col = (
    df_filtered[["GA property Name", "Host Name", "G-Tag"]]
    .drop_duplicates()
    .sort_values(by="GA property Name")
)

unique_props = sorted([p for p in mapping_3col["GA property Name"].unique() if p and p != "N/A"])
unique_hosts = sorted([h for h in mapping_3col["Host Name"].unique() if h and h != "N/A"])

f_col1, f_col2, f_col3 = st.columns(3)

with f_col1:
    selected_prop = st.selectbox(
        "🔍 Filter by GA Property Name:",
        options=["All GA Properties"] + unique_props,
    )

with f_col2:
    selected_host = st.selectbox(
        "🔍 Filter by Host Name:",
        options=["All Host Names"] + unique_hosts,
    )

with f_col3:
    search_gtag = st.text_input(
        "🔍 Search Correlated G-Tag:",
        placeholder="e.g. G-12345678",
        help="Type any full or partial G-Tag / Measurement ID to search.",
    )

display_mapping = mapping_3col.copy()

if selected_prop != "All GA Properties":
    display_mapping = display_mapping[display_mapping["GA property Name"] == selected_prop]

if selected_host != "All Host Names":
    display_mapping = display_mapping[display_mapping["Host Name"] == selected_host]

if search_gtag.strip():
    display_mapping = display_mapping[
        display_mapping["G-Tag"].str.contains(search_gtag.strip(), case=False, na=False)
    ]

if not display_mapping.empty:
    st.dataframe(
        display_mapping,
        column_config={
            "GA property Name": st.column_config.TextColumn("GA Property Name"),
            "Host Name": st.column_config.TextColumn("Host Name"),
            "G-Tag": st.column_config.TextColumn("Correlated G-Tag (Measurement ID)"),
        },
        use_container_width=True,
        hide_index=True,
    )

    excel_data_sec1 = convert_df_to_excel(display_mapping)
    st.download_button(
        label="📥 Export Lookup Key to Excel (.xlsx)",
        data=excel_data_sec1,
        file_name="GA_Property_Host_GTag_Mapping_Key.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.warning("No mappings found matching the selected search criteria.")

st.markdown("---")

# ==========================================
# SECTION 2: AGENCY SIZE & PORTFOLIO COMPARISONS
# ==========================================
st.header("2. 📊 Agency Portfolio & Size Comparisons")
st.info(
    "**What you are looking at:** Visual comparisons showing the size and scope of each State Agency's digital footprint. "
    "Compare total managed web paths, hostnames, and GA Property allocations across agencies."
)

# Aggregate metrics per agency
agency_stats = (
    df_filtered[df_filtered["Agency"] != "N/A"]
    .groupby("Agency")
    .agg(
        Total_Paths=("URL Path", "count"),
        Unique_Hosts=("Host Name", lambda x: len(set(h for h in x if h != "N/A"))),
        GA_Properties=("GA property Name", lambda x: len(set(p for p in x if p != "N/A"))),
    )
    .reset_index()
    .sort_values(by="Total_Paths", ascending=False)
)

if not agency_stats.empty:
    chart_tab1, chart_tab2, chart_tab3 = st.tabs(
        ["📈 Web Paths per Agency", "🧱 Digital Footprint Treemap", "📋 Agency Summary Data"]
    )

    with chart_tab1:
        top_n = st.slider("Select number of agencies to display:", min_value=5, max_value=max(len(agency_stats), 5), value=15)
        top_agencies = agency_stats.head(top_n)

        fig_bar = px.bar(
            top_agencies,
            x="Total_Paths",
            y="Agency",
            orientation="h",
            color="Total_Paths",
            color_continuous_scale="Blues",
            title=f"Top {top_n} Agencies by Total Web Paths / URLs Managed",
            labels={"Total_Paths": "Total Web Paths / URLs", "Agency": "State Agency"},
            text_auto=True,
        )
        fig_bar.update_layout(yaxis={"categoryorder": "total ascending"}, height=500)
        st.plotly_chart(fig_bar, use_container_width=True)

    with chart_tab2:
        fig_tree = px.treemap(
            agency_stats,
            path=["Agency"],
            values="Total_Paths",
            color="Unique_Hosts",
            color_continuous_scale="Viridis",
            title="Agency Portfolio Composition (Box Size = Total Paths | Color = Unique Hosts)",
            labels={"Total_Paths": "Total Paths", "Unique_Hosts": "Unique Hosts"},
        )
        fig_tree.update_layout(height=500)
        st.plotly_chart(fig_tree, use_container_width=True)

    with chart_tab3:
        st.dataframe(
            agency_stats,
            column_config={
                "Agency": st.column_config.TextColumn("State Agency"),
                "Total_Paths": st.column_config.NumberColumn("Total Web Paths / URLs"),
                "Unique_Hosts": st.column_config.NumberColumn("Unique Host Names"),
                "GA_Properties": st.column_config.NumberColumn("Distinct GA Properties"),
            },
            use_container_width=True,
            hide_index=True,
        )

        excel_agency = convert_df_to_excel(agency_stats)
        st.download_button(
            label="📥 Export Agency Comparison Summary to Excel (.xlsx)",
            data=excel_agency,
            file_name="Agency_Portfolio_Size_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("No agency data available for the current filter criteria.")

st.markdown("---")

# ==========================================
# SECTION 3: BROKEN LINKS & HTTP ERRORS
# ==========================================
st.header("3. 🚨 Broken Links & HTTP Status Errors")
st.info(
    "**What you are looking at:** URL paths currently returning error statuses (404 Not Found, 500 Server Errors, or broken links). "
    "Use this view to keep site lists clean and healthy."
)

error_df = df_filtered[df_filtered["Is_Broken_Or_Error"] == True]

if not error_df.empty:
    st.warning(
        f"Found **{len(error_df)}** paths returning error statuses in the current selection."
    )

    err_cols = [
        c
        for c in [
            "Agency",
            "Host Name",
            "URL Path",
            "Status",
            "G-Tag",
            "GA property Name",
        ]
        if c in error_df.columns
    ]
    st.dataframe(error_df[err_cols], use_container_width=True, hide_index=True)
else:
    st.success("🎉 No broken links or status errors found in the current selection!")

st.markdown("---")

# ==========================================
# SECTION 4: ONBOARDING CANDIDATES
# ==========================================
st.header("4. 🟡 Onboarding Candidates (Tag Firing but Not Found)")
st.warning(
    "**What you are looking at:** Web paths firing active tags that are **not yet in the State Analytics Account**. "
    "The **URL Path column is highlighted in yellow** as candidates for future GA360 onboarding."
)

candidates_df = df_filtered[
    df_filtered["GA property Name"].str.contains(
        "Tag Firing but Not Found", case=False, na=False
    )
].copy()

if not candidates_df.empty:

    def highlight_url_path(val):
        return "background-color: #FFF3CD; color: #856404; font-weight: bold;"

    cand_cols = [
        c
        for c in [
            "Agency",
            "Host Name",
            "URL Path",
            "G-Tag",
            "GA property Name",
            "Status",
        ]
        if c in candidates_df.columns
    ]

    styled_candidates = candidates_df[cand_cols].style.map(
        highlight_url_path, subset=["URL Path"]
    )
    st.dataframe(styled_candidates, use_container_width=True, hide_index=True)

    excel_data_sec4 = convert_df_to_excel(candidates_df[cand_cols])
    st.download_button(
        label="📥 Export Onboarding Candidates to Excel (.xlsx)",
        data=excel_data_sec4,
        file_name="GA360_Onboarding_Candidates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("No onboarding candidate URLs found matching current filters.")

st.markdown("---")

# ==========================================
# SECTION 5: MASTER DATA TABLE
# ==========================================
st.header("5. 📑 Complete Master Data Sheet")
st.info(
    "**What you are looking at:** The complete dataset loaded from `Masterspreadsheet.xlsx`. Use sidebar controls to filter rows."
)


def highlight_candidate_rows(row):
    if "Tag Firing but Not Found" in str(row.get("GA property Name", "")):
        return ["background-color: #FFF3CD;"] * len(row)
    return [""] * len(row)


st.dataframe(
    df_filtered.style.apply(highlight_candidate_rows, axis=1),
    use_container_width=True,
    hide_index=True,
)

st.markdown("---")

# ==========================================
# SECTION 6: NEUTRAL OBSERVATIONAL INSIGHTS
# ==========================================
st.header("6. ℹ️ Additional Observation: Host Names with Multiple Active G-Tags")
st.caption(
    "**Note:** If a host already has an official State Google Tag active, removing additional existing tags is not required. "
    "This table is provided purely as a neutral informational breakdown of hostnames detecting multiple tags across sub-paths."
)

multi_tag_df = (
    df_filtered[df_filtered["Has_Multiple_GTags"]]
    .groupby(["Agency", "Host Name"])
    .agg(
        Total_Tags_Detected=("Distinct_G_Tags", "first"),
        Detected_G_Tags=(
            "G-Tag",
            lambda x: ", ".join(sorted(set(g for g in x if g != "N/A"))),
        ),
        Total_Sub_Paths=("URL Path", "count"),
    )
    .reset_index()
)

if not multi_tag_df.empty:
    st.dataframe(
        multi_tag_df,
        column_config={
            "Total_Tags_Detected": st.column_config.NumberColumn(
                "Distinct Tags Count", format="%d 🏷️"
            ),
            "Detected_G_Tags": st.column_config.TextColumn("Detected Tags"),
            "Total_Sub_Paths": st.column_config.NumberColumn("Total Paths"),
        },
        use_container_width=True,
        hide_index=True,
    )
else:
    st.write("No multi-tag hostnames in current view.")

st.caption(
    f"State of Utah GA360 Governance Dashboard | Source: `{loaded_file}` [{loaded_sheet}]"
)