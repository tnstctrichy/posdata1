# Import necessary libraries
import streamlit as st
import pandas as pd
import plotly.express as px
from google.oauth2.service_account import Credentials
import gspread
from io import BytesIO
from datetime import datetime, timedelta

# Streamlit configuration
st.set_page_config(
    page_title="POS Dashboard - Advanced Visualization",
    layout="wide",
    page_icon="🚌"
)

# Styled Title Section
st.markdown(f"""
    <div style="border: 4px solid #4CAF50; padding: 15px; border-radius: 15px; background-color: #E3F2FD; text-align: center;">
        <h1 style="color: #FF5722; font-family: 'Segoe UI', sans-serif; font-size: 45px; margin-bottom: 5px;">
            POS Dashboard - TNSTC[KUM]LTD., KUMBAKONAM
        </h1>
        <h2 style="color: #4CAF50; font-family: 'Arial', sans-serif; font-size: 36px; margin-bottom: 0;">
            Data as of {(datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")}
        </h2>
    </div>
    """, unsafe_allow_html=True)

# Branch-to-region mapping
branch_to_region = {
    "KM1": "Kumbakonam", "KM2": "Kumbakonam", "TYR": "Kumbakonam", "TJ1": "Kumbakonam", "TJ2": "Kumbakonam",
    "OND": "Kumbakonam", "PKT": "Kumbakonam", "PVR": "Kumbakonam",
    "RFT": "Trichy", "DCN": "Trichy", "TVK": "Trichy", "LAL": "Trichy", "MCR": "Trichy", "TMF": "Trichy",
    "CNT": "Trichy", "MNP": "Trichy", "TKI": "Trichy", "PBR": "Trichy", "JKM": "Trichy", "ALR": "Trichy",
    "UPM": "Trichy", "TRR": "Trichy", "KNM": "Trichy",
    "KR1": "Karur", "KR2": "Karur", "AKI": "Karur", "KLI": "Karur", "MSI": "Karur",
    "PDK": "Pudukottai", "ATQ": "Pudukottai", "ALD": "Pudukottai", "PTK": "Pudukottai", "TRY": "Pudukottai",
    "ILP": "Pudukottai", "GVK": "Pudukottai", "PON": "Pudukottai",
    "KKD": "Karaikudi", "TPR": "Karaikudi", "MDU": "Karaikudi", "SVG": "Karaikudi", "DVK": "Karaikudi",
    "RNM": "Karaikudi", "RNT": "Karaikudi", "RMM": "Karaikudi", "KMD": "Karaikudi", "MDK": "Karaikudi",
    "PMK": "Karaikudi",
    "NGT": "Nagappattinam", "KKL": "Nagappattinam", "PYR": "Nagappattinam", "MLD": "Nagappattinam",
    "SHY": "Nagappattinam", "CDM": "Nagappattinam", "TVR": "Nagappattinam", "NLM": "Nagappattinam",
    "MNG": "Nagappattinam", "TTP": "Nagappattinam"
}

# Function to load and preprocess data
def load_google_sheet(sheet_url, gid, json_keyfile_dict):
    """
    Load and preprocess data from a Google Sheet.
    """
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(json_keyfile_dict, scopes=scopes)
        gc = gspread.authorize(credentials)
        sheet = gc.open_by_url(sheet_url)
        worksheets = {ws.id: ws for ws in sheet.worksheets()}
        worksheet = worksheets.get(int(gid))
        if not worksheet:
            st.error(f"No worksheet found with GID {gid}.")
            return None

        raw_data = worksheet.get_all_values()
        if not raw_data or len(raw_data) < 2:
            st.error("The worksheet has insufficient data or is empty.")
            return None

        headers = raw_data[0]
        data_rows = raw_data[1:]
        data = pd.DataFrame(data_rows, columns=headers)
        data.dropna(how="all", inplace=True)
        data.replace("", None, inplace=True)
        data.columns = [col.strip() for col in data.columns]

        if "BRANCH" in data.columns:
            data["REGION"] = data["BRANCH"].map(branch_to_region)
        else:
            st.warning("BRANCH column not found in the data.")
            data["REGION"] = None

        numeric_cols = [
            "POS MOF Total", "POS MOF Issued", "MOF POS %", "POS TOWN Total",
            "POS TOWN  Issued", "TOWN POS %", "OVER ALL MOF+TOWN Total",
            "OVER ALL MOF+TOWN POS Issued", "OVER ALL MOF+TOWN POS %",
            "MOF POS Tickets", "MOF Pre Printed  Tickets", "TOWN POS Tickets", "TOWN Pre Printed  Tickets"
        ]
        for col in numeric_cols:
            if col in data.columns:
                data[col] = pd.to_numeric(data[col], errors="coerce")

        if "POS IMPLEMENTED DATE" in data.columns:
            data["POS IMPLEMENTED DATE"] = pd.to_datetime(data["POS IMPLEMENTED DATE"], errors="coerce")

        return data
    except Exception as e:
        st.error(f"Error loading Google Sheet: {e}")
        return None

google_credentials = {
    "type": "service_account",
    "project_id": "posdashboard-442201",
    "private_key_id": "78a8b4fe3f25a90176067dd6b46f45f2abbf87f8",
    "private_key": (
        "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDEtdSLGwNIJtXc\nCBCbf8L/oA8F9a4P+sBkz8lRZQcAPo2GPASPNXcg4GRNncGi0L2L2cMrGh9IYE6K\nH4SFvutG1Xe1WEbFi0U0/j+2V0ULpxENLPhJmnkBanMts6APd9WvPdxFagT7NUvq\nYIKGvwrpawsq5fwgjmuiKDWgzn7XaIB5duUwAReMbbXOfDoZKp7l3dp5gyXEzAxD\neMVhd2La7agblmzrRP0E+32ct4hPyLrGcBDkcwjG559Zmk+q5/2MbhDrIS0P5ux7\n+qftVUN5WwxE9UrJoZXHoSxKz96A8aHJMevK55zPEP1JjTR8LPsWNW3jl1IU8KyT\nLZyLz52dAgMBAAECggEAGMZmnW81XnDXDkONGpX1+y3Eu/VBoIKY7mQwkLSpiTws\n+qNJgbiep0CLxtjKgBVL94wMOuarL6RC5WjrJHszdR7tWPo5ePzoUQXVWW5WubfS\nJneA/VLcJbOrYQNWx/afA6zGDDoPpDdbYeAY4HFkUES/9H2138CASZKd5Sx3fkKo\nKcB497f7YpHg/mgN07wxhB29QpN35E5UB3XlAOdBGk0vIJ9V3Lvl/OHBc+EiXkzV\nqzQLI9kgIdOAWJ8o3NaGx6te2vy5TUzT0uWOL7/eOvDpI6UV9yy6K7xrg9XtSpu8\nY6DkUEMdCLlxVXcfM2+SZE5XiJ86E17njIuuLV5L4wKBgQD5jTeMqL7ptQV1t1J/\n9IvDg2jxuyJDFo77ifkklY3Sefe4j4ZNNv0799iVAmAMf1nUYk0bp+4Xns2AQjCx\nAKKYPOx6YfiaMorALkmrH5mXYX6o2Tgv0BnVNoifjuHPhSm2LEbzgUeDhAgEZN2i\nOJb684KJbhe15GPKEp44RDmcSwKBgQDJyxFiuerVXYri9c4mLxUJHyBWCN5Yib+H\nQj6UyN1r61girPMysC2jWJpZ4oU7nv/6IXwUBZQ3o4/fnxC02y3vRz3GFJkMtcqs\no+kfFgbKUqkH4nSsJBjZQW1DfC+xIIBdT8k7oD2/Ux119R/9NLNnD7iOvFKaC+7Y\n1RQHfMwstwKBgEsa4TkIIE0eGgKPpdi0tMum5RK7i1g9ldLGd6E3EXPjGVcGexkK\nD7TYpupRyK56NYLiAurr45BgTuDnCth6pHTFATbj/XoK9A9a3vkNjaAty3ztwydA\nrkWpH/1Fd1iJb0BQmxn2Mpu2RONtp/aGqYnld8f8xk4L6qyKZevxPJV5AoGBALSk\noM+sd1jCAI7kVMNB6qbbwmrCTakcxuQinTs8BVuStrdz89IwfOp5atOEQJj64VPd\nneGejOyx8x3Qm3gLrbdCIz6rOcdzBhg+M3aslS+Rh9eTFbb0KXpzY4jCJz99ROxD\nfHVwIVag5QKviQ92mhNss16zn45fmFVrih6ZzX1JAoGAaOChj3CbcuDmSYxfcV9C\nGYdFBvYwCCWeYMd6EWL+V12dV7Sqf0EpETsNQlvbcTweNWoo04fdURByV/t6pEkx\niaNhclcH7iCSu7t0mZVqo2QaUVqErNpVh4+pqQUplCwopLCQeJ3SvRX4OJr6VRha\noULq6+BNyVLozX4LdNPMutI=\n-----END PRIVATE KEY-----\n"
    ),
    "client_email": "datapos@posdashboard-442201.iam.gserviceaccount.com",
    "client_id": "116030208639366774536",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/datapos%40posdashboard-442201.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

sheet_url = "https://docs.google.com/spreadsheets/d/1zjn_8Qi0RAdffzsfOmYWK3cqqVGdveaIrDABzB6BcRk/edit"
sheet_gid = "1150984969"

data = load_google_sheet(sheet_url, sheet_gid, google_credentials)

if data is not None:
    # Custom CSS for tab styling
    st.markdown("""
    <style>
    /* Customize the tab menu styles */
    .stTabs button {
        font-size: 24px !important;
        font-weight: bold !important;
        color: #333333 !important;  /* Darker text for better readability */
        background-color: #E3F2FD !important;  /* Soft blue background */
        border: 2px solid #B0BEC5 !important;  /* Light grey border for contrast */
        border-radius: 10px;
        padding: 10px 20px;
        margin: 6px;
        transition: background-color 0.3s ease, color 0.3s ease;  /* Smooth transitions */
    }
    
    /* Customize the active tab */
    .stTabs button[aria-selected="true"] {
        color: white !important;
        background-color: #0078d4 !important;  /* Strong blue background for active tab */
        border: 2px solid #005B9A !important;  /* Darker blue border for active tab */
        border-radius: 12px;
    }
    
    /* Hover effect for tabs */
    .stTabs button:hover {
        background-color: #B3E5FC !important;  /* Slightly darker blue when hovered */
        color: #0078d4 !important;  /* Change text color to blue when hovered */
    }
    </style>
    """, unsafe_allow_html=True)

    # Reusable chart plotting functions
    def plot_bar_chart(df, x_col, y_cols, title, labels, color_scheme, container_width=True):
        """
        Function to plot bar charts using Plotly.
        """
        fig = px.bar(
            df, x=x_col, y=y_cols, title=title, labels=labels,
            barmode="group", color_discrete_sequence=color_scheme
        )
        st.plotly_chart(fig, use_container_width=container_width)

    def plot_pie_chart(df, names_col, values_col, title, color_scheme, container_width=True):
        """
        Function to plot pie charts using Plotly.
        """
        fig = px.pie(
            df, names=names_col, values=values_col, title=title,
            color_discrete_sequence=color_scheme
        )
        st.plotly_chart(fig, use_container_width=container_width)

    def display_region_metrics(region, region_data):
    """
    Function to display metrics and charts for a specific region.
    """
    st.markdown(f"### Region: {region}")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### MOF POS Services")
        plot_bar_chart(
            region_data, "BRANCH", ["POS MOF Total", "POS MOF Issued"],
            f"MOF POS Services in {region}",
            {"value": "POS Metrics", "BRANCH": "Branch"},
            px.colors.qualitative.Set2
        )
        st.markdown(f"**MOF POS %**: {region_data['MOF POS %'].mean():.2f}%")

    with col2:
        st.markdown("#### Town POS Services")
        plot_bar_chart(
            region_data, "BRANCH", ["POS TOWN Total", "POS TOWN  Issued"],
            f"Town POS Services in {region}",
            {"value": "POS Metrics", "BRANCH": "Branch"},
            px.colors.qualitative.Set3
        )
        st.markdown(f"**Town POS %**: {region_data['TOWN POS %'].mean():.2f}%")

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("#### MOF Ticket Distribution")
        plot_pie_chart(
            region_data, "BRANCH", "MOF POS Tickets",
            f"MOF POS Tickets Distribution in {region}",
            px.colors.qualitative.Pastel
        )
        plot_pie_chart(
            region_data, "BRANCH", "MOF Pre Printed  Tickets",
            f"MOF Preprinted Tickets Distribution in {region}",
            px.colors.qualitative.Bold
        )

    with col4:
        st.markdown("#### Town Ticket Distribution")
        plot_pie_chart(
            region_data, "BRANCH", "TOWN POS Tickets",
            f"Town POS Tickets Distribution in {region}",
            px.colors.qualitative.Set2
        )
        plot_pie_chart(
            region_data, "BRANCH", "TOWN Pre Printed  Tickets",
            f"Town Preprinted Tickets Distribution in {region}",
            px.colors.qualitative.Set1
        )

    # Additional Pie Chart: MOF POS Tickets vs MOF Preprinted Tickets
    st.markdown("#### MOF POS Tickets vs MOF Preprinted Tickets")
    mof_ticket_vs_preprint = region_data[["BRANCH", "MOF POS Tickets", "MOF Pre Printed  Tickets"]]
    mof_ticket_vs_preprint = mof_ticket_vs_preprint.melt(id_vars="BRANCH", var_name="Type", value_name="Count")
    pie_fig_mof = px.pie(
        mof_ticket_vs_preprint, names="Type", values="Count",
        title=f"MOF POS Tickets vs Preprinted Tickets in {region}",
        color_discrete_sequence=px.colors.qualitative.Bold
    )
    st.plotly_chart(pie_fig_mof, use_container_width=True)
        
    # Create the tabs
    tab3, tab1, tab2 = st.tabs([
        "📊 Kumbakonam Corporation - Consolidated View",
        "🔎 Kumbakonam - Regionwise View",
        "📥 Download - Excel Sheet Data"
    ])

    # Global filtered_data for consistent use
    filtered_data = data

    # Tab 1: Region-wise View with Filters
    with tab1:
        st.markdown("### Filters")
        regions = st.multiselect(
            "Select Regions", 
            options=data["REGION"].dropna().unique(), 
            default=data["REGION"].dropna().unique()
        )
        filtered_data = data[data["REGION"].isin(regions)] if regions else data

        # Ensure valid regions exist
        valid_regions = filtered_data["REGION"].dropna().unique().tolist()
        #valid_regions = filtered_data["REGION"].dropna().unique().tolist()
        #region_tabs = st.tabs(valid_regions)  # Works perfectly


        if len(valid_regions) > 0:
            st.markdown("### Key Metrics by Region")
            region_tabs = st.tabs(valid_regions)

            for tab, region in zip(region_tabs, valid_regions):
                region_data = filtered_data[filtered_data["REGION"] == region]
                with tab:
                    display_region_metrics(region, region_data)
        else:
            st.warning("No valid regions to display.")

    # Tab 2: Download Filtered Data
    with tab2:
        st.markdown("### Download Filtered Data")
        if not filtered_data.empty:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                filtered_data.to_excel(writer, index=False, sheet_name="Filtered Data")
            buffer.seek(0)
            st.download_button(
                label="Download Filtered Data (Excel)",
                data=buffer,
                file_name="filtered_pos_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data to download.")

    
   # Tab 3: Consolidated View
# Tab 3: Consolidated View
# Tab 3: Consolidated View
with tab3:
    st.markdown("### Consolidated View for All Regions")

    # Sort regions in a specific order
    region_order = ["Kumbakonam", "Trichy", "Karur", "Pudukottai", "Karaikudi", "Nagappattinam"]
    consolidated_data = filtered_data.groupby("REGION", as_index=False).sum(numeric_only=True)

    # Reorder the regions
    consolidated_data = consolidated_data.set_index("REGION").loc[region_order].reset_index()

    # Recalculate percentages
    consolidated_data["MOF POS %"] = (
        (consolidated_data["POS MOF Issued"] / consolidated_data["POS MOF Total"]) * 100
    ).fillna(0).round(2)
    consolidated_data["TOWN POS %"] = (
        (consolidated_data["POS TOWN  Issued"] / consolidated_data["POS TOWN Total"]) * 100
    ).fillna(0).round(2)

    # Add a record ID column starting from 1
    consolidated_data.insert(0, "ID", range(1, len(consolidated_data) + 1))

    # Display Key Metrics Table by Region
    st.markdown("#### Key Metrics Table by Region")
    st.dataframe(consolidated_data[[
        "ID", "REGION", "POS MOF Total", "POS MOF Issued", "MOF POS %",
        "POS TOWN Total", "POS TOWN  Issued", "TOWN POS %",
        "MOF POS Tickets", "MOF Pre Printed  Tickets",
        "TOWN POS Tickets", "TOWN Pre Printed  Tickets"
    ]], use_container_width=True, hide_index=True)

    # Consolidated MOF POS Bar Chart
    st.markdown("#### MOF POS Services - Consolidated")
    plot_bar_chart(
        consolidated_data, "REGION", ["POS MOF Total", "POS MOF Issued"],
        "Consolidated MOF POS Services",
        {"value": "POS Metrics", "REGION": "Region"},
        px.colors.qualitative.Set2
    )
    st.markdown(f"**MOF POS Services**: ==>> {consolidated_data['MOF POS %'].mean():.2f}%")

    # Consolidated Town POS Bar Chart
    st.markdown("#### Town POS Services - Consolidated")
    plot_bar_chart(
        consolidated_data, "REGION", ["POS TOWN Total", "POS TOWN  Issued"],
        "Consolidated Town POS Services",
        {"value": "POS Metrics", "REGION": "Region"},
        px.colors.qualitative.Set3
    )
    st.markdown(f"**Town POS Services**: ==>> {consolidated_data['TOWN POS %'].mean():.2f}%")

    # Additional Pie Charts in the new order
    st.markdown("#### Ticket Distribution: POS Tickets and Preprinted Tickets (Consolidated)")
    col1, col2 = st.columns(2)
    
    with col1:
        # MOF POS Tickets Distribution
        plot_pie_chart(
            consolidated_data, "REGION", "MOF POS Tickets",
            "MOF POS Tickets Distribution", px.colors.qualitative.Pastel
        )
        
        # MOF Preprinted Tickets Sales Distribution
        plot_pie_chart(
            consolidated_data, "REGION", "MOF Pre Printed  Tickets",
            "MOF Preprinted Tickets Sales Distribution", px.colors.qualitative.Bold
        )
    
    with col2:
        # Town POS Tickets Distribution
        plot_pie_chart(
            consolidated_data, "REGION", "TOWN POS Tickets",
            "Town POS Tickets Distribution", px.colors.qualitative.Set2
        )
        
        # Town Preprinted Tickets Sales Distribution
        plot_pie_chart(
            consolidated_data, "REGION", "TOWN Pre Printed  Tickets",
            "Town Preprinted Tickets Sales Distribution", px.colors.qualitative.Set1
        )

    # Additional Pie Charts: POS Tickets vs Preprinted Tickets
    st.markdown("#### Comparison: POS Tickets vs Preprinted Tickets")
    col3, col4 = st.columns(2)
    
    with col3:
        st.markdown("##### MOF POS Tickets vs MOF Preprinted Tickets")
        mof_ticket_vs_preprint = consolidated_data[["REGION", "MOF POS Tickets", "MOF Pre Printed  Tickets"]]
        mof_ticket_vs_preprint = mof_ticket_vs_preprint.melt(id_vars="REGION", var_name="Type", value_name="Count")
        pie_fig = px.pie(
            mof_ticket_vs_preprint, names="Type", values="Count",
            #title="MOF POS Tickets vs Preprinted Tickets",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        st.plotly_chart(pie_fig, use_container_width=True)
    
    with col4:
        st.markdown("##### Town POS Tickets vs Town Preprinted Tickets")
        town_ticket_vs_preprint = consolidated_data[["REGION", "TOWN POS Tickets", "TOWN Pre Printed  Tickets"]]
        town_ticket_vs_preprint = town_ticket_vs_preprint.melt(id_vars="REGION", var_name="Type", value_name="Count")
        pie_fig_town = px.pie(
            town_ticket_vs_preprint, names="Type", values="Count",
            #title="Town POS Tickets vs Preprinted Tickets",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        st.plotly_chart(pie_fig_town, use_container_width=True)
