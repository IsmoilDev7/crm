import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.figure_factory as ff

st.set_page_config(page_title="CRM Analysis", layout="wide")
st.title("📊 CRM Excel Analysis Dashboard")
st.markdown("Upload your Excel file to visualize and analyze the data.")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 1️⃣ Strip all column names and lowercase to avoid invisible chars
    df.columns = df.columns.str.strip().str.replace('\xa0','').str.replace('\n','').str.replace('\r','').str.lower()

    # 2️⃣ Expected columns (all lowercase for soft match)
    expected_cols = ['stage','source','date of creation'.lower(),'responsible','date modified','company']

    # 3️⃣ Check missing columns
    missing_cols = [col for col in expected_cols if col not in df.columns]

    if len(missing_cols) == 0:
        st.success("✅ Columns detected successfully!")

        # Map actual column names back to readable
        rename_map = {col: col.title() for col in df.columns}
        df.rename(columns=rename_map, inplace=True)

        # Convert dates
        df['Date Of Creation'] = pd.to_datetime(df['Date Of Creation'], errors='coerce', dayfirst=True)
        df['Date Modified'] = pd.to_datetime(df['Date Modified'], errors='coerce', dayfirst=True)

        # Sidebar filters
        st.sidebar.header("Filters")
        stages = df['Stage'].dropna().unique()
        responsibles = df['Responsible'].dropna().unique()
        selected_stage = st.sidebar.multiselect("Select Stage", stages, default=stages)
        selected_responsible = st.sidebar.multiselect("Select Responsible", responsibles, default=responsibles)

        # Filtered data
        df_filtered = df[df['Stage'].isin(selected_stage) & df['Responsible'].isin(selected_responsible)]
        st.subheader("Filtered Data")
        st.dataframe(df_filtered)

        # Summary metrics
        st.subheader("Summary Metrics")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Companies", df_filtered['Company'].nunique())
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

        # Time trends
        st.subheader("Records Over Time")
        df_filtered['Creation Date'] = df_filtered['Date Of Creation'].dt.date
        fig_time = px.line(df_filtered.groupby('Creation Date').size().reset_index(name='Count'),
                           x='Creation Date', y='Count', title='Records Created Over Time')
        st.plotly_chart(fig_time, use_container_width=True)

        df_filtered['Modified Date'] = df_filtered['Date Modified'].dt.date
        fig_mod = px.line(df_filtered.groupby('Modified Date').size().reset_index(name='Count'),
                          x='Modified Date', y='Count', title='Records Modified Over Time')
        st.plotly_chart(fig_mod, use_container_width=True)

        # Stage × Responsible Heatmap
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

        # Top 10 active companies
        st.subheader("Top 10 Companies by Records")
        top_companies = df_filtered.groupby('Company').size().reset_index(name='Records').sort_values(by='Records', ascending=False).head(10)
        st.plotly_chart(px.bar(top_companies, x='Company', y='Records', text='Records', title='Top 10 Active Companies'), use_container_width=True)

        # Download filtered CSV
        st.subheader("Download Filtered Data")
        csv = df_filtered.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download CSV", csv, "filtered_data.csv", "text/csv")

    else:
        st.error(f"❌ Missing columns detected (check invisible chars): {missing_cols}")

else:
    st.info("Please upload your Excel file to start analysis.")
