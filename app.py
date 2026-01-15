import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.figure_factory as ff

st.set_page_config(page_title="CRM Dashboard", layout="wide")
st.title("📊 CRM Excel Analysis Dashboard")
st.markdown("Upload your Excel file to visualize and analyze data interactively.")

# 1️⃣ Upload Excel
uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=0)

    # Strip invisible chars from columns
    df.columns = df.columns.str.strip().str.replace('\xa0','').str.replace('\n','').str.replace('\r','')

    # Expected columns
    expected_cols = ['Stage', 'Source', 'Date of creation', 'Responsible', 'Date modified', 'Company name']
    missing_cols = [col for col in expected_cols if col not in df.columns]

    if missing_cols:
        st.error(f"❌ Missing columns: {missing_cols}")
        st.write("Detected columns:", df.columns.tolist())
    else:
        st.success("✅ Columns detected!")

        # Convert dates
        df['Date of creation'] = pd.to_datetime(df['Date of creation'], errors='coerce', dayfirst=True)
        df['Date modified'] = pd.to_datetime(df['Date modified'], errors='coerce', dayfirst=True)

        # Sidebar filters
        st.sidebar.header("Filters")
        selected_stage = st.sidebar.multiselect("Stage", df['Stage'].dropna().unique(), default=df['Stage'].dropna().unique())
        selected_responsible = st.sidebar.multiselect("Responsible", df['Responsible'].dropna().unique(), default=df['Responsible'].dropna().unique())
        selected_source = st.sidebar.multiselect("Source", df['Source'].dropna().unique(), default=df['Source'].dropna().unique())
        start_date, end_date = st.sidebar.date_input(
            "Date of creation range",
            [df['Date of creation'].min().date(), df['Date of creation'].max().date()]
        )

        # Filter data
        df_filtered = df[
            (df['Stage'].isin(selected_stage)) &
            (df['Responsible'].isin(selected_responsible)) &
            (df['Source'].isin(selected_source)) &
            (df['Date of creation'].dt.date >= start_date) &
            (df['Date of creation'].dt.date <= end_date)
        ]

        st.subheader("Filtered Data")
        st.dataframe(df_filtered)

        # Summary metrics
        st.subheader("Summary Metrics")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Companies", df_filtered['Company name'].nunique())
        col2.metric("Total Records", df_filtered.shape[0])
        col3.metric("Total Responsible", df_filtered['Responsible'].nunique())
        col4.metric("Total Sources", df_filtered['Source'].nunique())

        # Pie charts
        st.subheader("Stage Distribution")
        st.plotly_chart(px.pie(df_filtered, names='Stage', title='Stage Distribution'), use_container_width=True)

        st.subheader("Responsible Distribution")
        st.plotly_chart(px.pie(df_filtered, names='Responsible', title='Responsible Distribution'), use_container_width=True)

        st.subheader("Source Distribution")
        st.plotly_chart(px.pie(df_filtered, names='Source', title='Source Distribution'), use_container_width=True)

        # Time trend charts
        st.subheader("Records Over Time (Date of creation)")
        df_filtered['Creation Date'] = df_filtered['Date of creation'].dt.date
        time_data = df_filtered.groupby('Creation Date').size().reset_index(name='Count')
        st.plotly_chart(px.line(time_data, x='Creation Date', y='Count', title='Records Created Over Time'), use_container_width=True)

        st.subheader("Records Over Time (Date modified)")
        df_filtered['Modified Date'] = df_filtered['Date modified'].dt.date
        mod_time_data = df_filtered.groupby('Modified Date').size().reset_index(name='Count')
        st.plotly_chart(px.line(mod_time_data, x='Modified Date', y='Count', title='Records Modified Over Time'), use_container_width=True)

        # Stage × Responsible heatmap
        st.subheader("Stage × Responsible Heatmap")
        heatmap_data = df_filtered.groupby(['Stage','Responsible']).size().reset_index(name='Count')
        heatmap_pivot = heatmap_data.pivot(index='Stage', columns='Responsible', values='Count').fillna(0)
        fig_heatmap = ff.create_annotated_heatmap(
            z=heatmap_pivot.values,
            x=list(heatmap_pivot.columns),
            y=list(heatmap_pivot.index),
            colorscale='Viridis',
            showscale=True
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

        # Top 10 companies
        st.subheader("Top 10 Companies by Records")
        top_companies = df_filtered.groupby('Company name').size().reset_index(name='Records').sort_values(by='Records', ascending=False).head(10)
        st.plotly_chart(px.bar(top_companies, x='Company name', y='Records', text='Records', title='Top 10 Companies'), use_container_width=True)

        # CSV download
        csv = df_filtered.to_csv(index=False).encode('utf-8')
        st.download_button("Download Filtered CSV", csv, "filtered_data.csv", "text/csv")

else:
    st.info("Upload your Excel file to start analysis.")
