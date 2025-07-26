import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os

st.set_page_config(page_title="VLU 4G Report App", layout="wide")
st.title("📡 VLU 4G Report Generator")

uploaded_file = st.file_uploader("📥 Upload your Excel file (.xlsx)", type="xlsx")

if uploaded_file:
    # Read sheets
    sheet1 = pd.read_excel(uploaded_file, sheet_name="Sheet1")
    sitelist = pd.read_excel(uploaded_file, sheet_name="SITELIST")

    # Build helper DataFrame
    unique_locations = sheet1['G'].dropna().unique()
    helper = pd.DataFrame({'Location': unique_locations})
    helper['Total_GB'] = helper['Location'].apply(
        lambda loc: sheet1[sheet1['G'] == loc]['L'].sum()
    )
    helper['SDCA'] = helper['Location'].apply(
        lambda loc: sheet1[sheet1['G'] == loc]['S'].iloc[0] if not sheet1[sheet1['G'] == loc].empty else ''
    )
    helper['VOLTE_ERL'] = helper['Location'].apply(
        lambda loc: sheet1[sheet1['G'] == loc]['Q'].sum()
    )

    selected_option = st.selectbox("📊 Select Report Type", [
        "Location-wise Data Usage",
        "Location-wise VOLTE Call Report",
        "Location Detail View",
        "Less Traffic Sites",
        "Generate PDF Report"
    ])

    def show_data_usage(sdca="Villupuram"):
        df = helper[helper['SDCA'].str.strip().str.lower() == sdca.lower()]
        df_sorted = df.sort_values(by="Total_GB", ascending=False).reset_index(drop=True)
        df_sorted.index += 1
        st.subheader(f"📍 Data Usage - SDCA: {sdca}")
        st.dataframe(df_sorted[['Location', 'Total_GB', 'VOLTE_ERL']])
        return df_sorted

    def show_volte_report(sdca="Villupuram"):
        df = helper[helper['SDCA'].str.strip().str.lower() == sdca.lower()]
        df_sorted = df.sort_values(by="VOLTE_ERL", ascending=False).reset_index(drop=True)
        df_sorted.index += 1
        st.subheader(f"📞 VOLTE Report - SDCA: {sdca}")
        st.dataframe(df_sorted[['Location', 'VOLTE_ERL', 'Total_GB']])
        return df_sorted

    def show_location_details():
        location = st.selectbox("📍 Select Location", sorted(helper['Location'].unique()))
        df_filtered = sheet1[sheet1['G'].str.strip().str.lower() == location.strip().lower()]
        df_out = df_filtered[['G', 'I', 'L', 'Q']]
        df_out.columns = ['Location', '4G Cell Name', 'Total (GB)', 'VOLTE Traffic']
        st.dataframe(df_out)
        return location, df_out

    def show_less_traffic(sdca="Villupuram"):
        df = helper[helper['SDCA'].str.strip().str.lower() == sdca.lower()]
        st.subheader("📉 Locations with < 10 GB Usage")
        st.dataframe(df[df['Total_GB'] < 10][['Location', 'Total_GB', 'SDCA']])
        st.subheader("📉 Zero Data Sectors")
        df_zero = sheet1[
            (sheet1['S'].str.strip().str.lower() == sdca.lower()) &
            (sheet1['L'] == 0)
        ][['G', 'I', 'Q', 'L']]
        df_zero.columns = ['Location', '4G Cell Name', 'VOLTE Traffic', 'Total (GB)']
        st.dataframe(df_zero)
        return df, df_zero

    def generate_pdf_report():
        date_cell = sheet1.at[1, 'R']
        report_date = pd.to_datetime(date_cell).strftime("%d-%m-%Y") if pd.notna(date_cell) else datetime.today().strftime("%d-%m-%Y")

        usage_df = show_data_usage()
        volte_df = show_volte_report()
        _, location_details_df = show_location_details()
        _, zero_data_df = show_less_traffic()

        # Render HTML
        env = Environment(loader=FileSystemLoader("templates"))
        template = env.get_template("report_template.html")
        html_content = template.render(
            date=report_date,
            data_usage=usage_df.to_html(index=False),
            volte_data=volte_df.to_html(index=False),
            detail_view=location_details_df.to_html(index=False),
            zero_data=zero_data_df.to_html(index=False)
        )

        pdf_bytes = HTML(string=html_content).write_pdf()
        st.download_button(
            label="📄 Download PDF Report",
            data=pdf_bytes,
            file_name=f"VLU_4G_Report_{report_date}.pdf",
            mime="application/pdf"
        )

    if selected_option == "Location-wise Data Usage":
        show_data_usage()

    elif selected_option == "Location-wise VOLTE Call Report":
        show_volte_report()

    elif selected_option == "Location Detail View":
        show_location_details()

    elif selected_option == "Less Traffic Sites":
        show_less_traffic()

    elif selected_option == "Generate PDF Report":
        generate_pdf_report()
