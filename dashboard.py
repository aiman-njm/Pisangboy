import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import requests
from io import BytesIO
import os

# Page settings
st.set_page_config(page_title="Project Tracker Dashboard", layout="wide")
st.title("📊 Project Tracker Dashboard")

# Excel file local path
EXCEL_FILE = "Project_Dashboard_Data.xlsx"
# Optional: SharePoint/OneDrive link
SHAREPOINT_URL = os.environ.get("PROJECT_DASHBOARD_URL")

@st.cache_data(ttl=300)
def load_excel():
    if os.path.exists(EXCEL_FILE):
        return pd.ExcelFile(EXCEL_FILE)
    if SHAREPOINT_URL:
        resp = requests.get(SHAREPOINT_URL)
        if resp.status_code == 200:
            return pd.ExcelFile(BytesIO(resp.content))
    return None

xls = load_excel()

if not xls:
    st.error("❌ Excel data could not be loaded.")
    st.stop()

# Exclude unwanted sheets
valid_sheets = [s for s in xls.sheet_names if s not in ['April 2025(Demo)', 'Skill Set']]
selected_sheets = st.sidebar.multiselect("🗂️ Select Sheets", valid_sheets, default=valid_sheets)

if selected_sheets:
    combined_df = pd.DataFrame()
    for sheet in selected_sheets:
        df = pd.read_excel(xls, sheet_name=sheet)
        if 'Start Date' not in df.columns:
            st.warning(f"⚠️ Skipping sheet '{sheet}' (missing 'Start Date').")
            continue
        df['Sheet'] = sheet
        df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce').dt.strftime('%d %b %Y')
        if 'End Date' in df.columns:
            df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce').dt.strftime('%d %b %Y')
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    combined_df.dropna(axis=1, how='all', inplace=True)

    # Sidebar filters
    st.sidebar.header("🔎 Filter Options")
    status_filter = st.sidebar.multiselect(
        "✅ Filter by Status",
        combined_df['Status'].dropna().unique(),
        default=combined_df['Status'].dropna().unique(),
    )
    health_filter = st.sidebar.multiselect(
        "🩺 Filter by Health",
        combined_df['Health'].dropna().unique(),
        default=combined_df['Health'].dropna().unique(),
    )
    manager_filter = st.sidebar.multiselect(
        "👤 Filter by Project Manager",
        combined_df['Project Manager'].dropna().unique(),
        default=combined_df['Project Manager'].dropna().unique(),
    )
    show_raw = st.sidebar.checkbox("📂 Show Raw Data")

    filtered_df = combined_df[
        combined_df['Status'].isin(status_filter)
        & combined_df['Health'].isin(health_filter)
        & combined_df['Project Manager'].isin(manager_filter)
    ]

    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        "⬇️ Download Filtered Data",
        data=csv,
        file_name="filtered_projects.csv",
        mime="text/csv",
    )

    if show_raw:
        st.markdown("### 🗃️ Raw Data (Unfiltered)")
        st.dataframe(combined_df)

    st.markdown("### 📌 Filtered Project Data")
    st.dataframe(filtered_df, use_container_width=True)

    st.markdown("### 👥 Project Count by Project Manager")
    pm_counts = filtered_df['Project Manager'].value_counts()
    col1, col2 = st.columns(2)

    with col1:
        st.dataframe(pm_counts.rename_axis('Project Manager').reset_index(name='Total Projects'))

    with col2:
        st.markdown("#### 🧁 Projects by Manager")
        fig_pie, ax_pie = plt.subplots(figsize=(4, 2.5))
        wedges, texts, autotexts = ax_pie.pie(
            pm_counts,
            labels=pm_counts.index,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 6},
            pctdistance=0.8,
        )
        for text in texts:
            text.set_fontsize(6)
        for autotext in autotexts:
            autotext.set_fontsize(6)
        ax_pie.axis("equal")
        fig_pie.tight_layout()
        st.pyplot(fig_pie)

    st.markdown("<h2 style='color:black;'>📊 Visual Insights</h2>", unsafe_allow_html=True)
    col3, col4 = st.columns(2)

    with col3:
        st.markdown("#### 🏢 Projects by Division")
        div_counts = filtered_df['Division'].value_counts()
        fig_div, ax_div = plt.subplots(figsize=(4, 2.5))
        ax_div.bar(div_counts.index, div_counts.values, color=plt.cm.Paired.colors)
        ax_div.set_ylabel("Projects", fontsize=9)
        ax_div.set_xlabel("Division", fontsize=9)
        ax_div.set_title("Projects by Division", fontsize=9)
        ax_div.tick_params(axis='x', labelrotation=45, labelsize=8)
        ax_div.tick_params(axis='y', labelsize=8)
        fig_div.tight_layout()
        st.pyplot(fig_div)

    with col4:
        st.markdown("#### 💊 Health Status Overview")
        expected_health = ['Green', 'Yellow', 'Red']
        health_counts = filtered_df['Health'].value_counts().reindex(expected_health, fill_value=0)
        color_map = {"Green": "#81C784", "Yellow": "#FFF176", "Red": "#E57373"}
        bar_colors = [color_map.get(status, "#90CAF9") for status in health_counts.index]
        fig_health, ax_health = plt.subplots(figsize=(4, 2.5))
        ax_health.bar(health_counts.index, health_counts.values, color=bar_colors)
        ax_health.set_ylabel("Projects", fontsize=9)
        ax_health.set_xlabel("Health Status", fontsize=9)
        ax_health.set_title("Health Status Distribution", fontsize=9)
        ax_health.tick_params(axis='x', labelsize=8)
        ax_health.tick_params(axis='y', labelsize=8)
        fig_health.tight_layout()
        st.pyplot(fig_health)

    spacer1, col5, spacer2 = st.columns([1, 6, 1])

    with col5:
        st.markdown("#### 📆 Task Trend Over Time")
        df_trend = filtered_df.copy()
        df_trend['Start Date'] = pd.to_datetime(df_trend['Start Date'], errors='coerce')
        df_trend = df_trend.dropna(subset=['Start Date', 'Tasks'])
        df_trend_grouped = df_trend.groupby(df_trend['Start Date'].dt.strftime('%d %b %Y'))['Tasks'].sum()

        fig_trend, ax_trend = plt.subplots(figsize=(6.5, 3))
        ax_trend.plot(
            df_trend_grouped.index,
            df_trend_grouped.values,
            marker='o',
            markersize=5,
            linewidth=2,
            color="#5C6BC0",
        )
        ax_trend.set_ylabel("Total Tasks", fontsize=9)
        ax_trend.set_xlabel("Date", fontsize=9)
        ax_trend.set_title("Daily Task Volume", fontsize=9)
        ax_trend.tick_params(axis='x', labelsize=8)
        ax_trend.tick_params(axis='y', labelsize=8)
        plt.xticks(rotation=30, ha='right')
        fig_trend.tight_layout()
        st.pyplot(fig_trend)

# working link with downloaded excel file
