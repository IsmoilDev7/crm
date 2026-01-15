import pandas as pd
import streamlit as st
import plotly.express as px

# Streamlit page setup
st.set_page_config(page_title="Company Data Dashboard", layout="wide")
st.title("📊 Company Data Dashboard")
st.markdown("Upload your Excel file and get instant analysis!")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 1️⃣ Strip all column names to remove invisible chars
    df.columns = df.columns.str.strip().str.replace('\xa0','').str.replace('\n','').str.replace('\r','')

    # 2️⃣ Expected columns (soft match)
    expected_cols = ['Stage', 'Source', 'Date of creation', 'Responsible', 'Date modified', 'Company name']

    # 3️⃣ Check if expected columns exist (case insensitive)
    df_cols_lower = [c.lower() for c in df.columns]
    missing_cols = [col for col in expected_cols if col.lower() not in df_cols_lower]

    if len(missing_cols) == 0:
        st.success("✅ Columns are OK. Performing analysis...")

        # Map actual column names to expected (ignore case)
        col_map = {col.lower(): col for col in df.columns}
        df = df.rename(columns={col.lower(): col_map[col.lower()] for col in expected_cols})

        # 4️⃣ Convert date columns
        df['Date of creation'] = pd.to_datetime(df['Date of creation'], format='%d.%m.%Y %H:%M:%S')
        df['Date modified'] = pd.to_datetime(df['Date modified'], format='%d.%m.%Y %H:%M:%S')

        # 5️⃣ Sidebar filters
        st.sidebar.header("Filters")
        selected_stage = st.sidebar.multiselect("Select Stage", df['Stage'].unique(), default=df['Stage'].unique())
        selected_responsible = st.sidebar.multiselect("Select Responsible", df['Responsible'].unique(), default=df['Responsible'].unique())

        # 6️⃣ Filtered data
        df_filtered = df[(df['Stage'].isin(selected_stage)) & (df['Responsible'].isin(selected_responsible))]

        st.subheader("Filtered Data")
        st.dataframe(df_filtered)

        # 7️⃣ Summary metrics
        st.subheader("Summary Statistics")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Companies", df_filtered['Company name'].nunique())
        col2.metric("Total Records", df_filtered.shape[0])
        col3.metric("Total Responsible", df_filtered['Responsible'].nunique())

        # 8️⃣ Pie charts
        st.subheader("Stage Distribution")
        st.plotly_chart(px.pie(df_filtered, names='Stage', title='Stage Distribution'), use_container_width=True)
        st.subheader("Responsible Distribution")
        st.plotly_chart(px.pie(df_filtered, names='Responsible', title='Responsible Distribution'), use_container_width=True)

        # 9️⃣ Time trend
        st.subheader("Records Over Time")
        df_filtered['Creation Date'] = df_filtered['Date of creation'].dt.date
        df_time = df_filtered.groupby('Creation Date').size().reset_index(name='Count')
        st.plotly_chart(px.line(df_time, x='Creation Date', y='Count', title='Records Created Over Time'), use_container_width=True)

        # 10️⃣ Companies per Responsible
        st.subheader("Companies Managed by Responsible")
        df_resp_count = df_filtered.groupby('Responsible')['Company name'].nunique().reset_index().sort_values(by='Company name', ascending=False)
        st.plotly_chart(px.bar(df_resp_count, x='Responsible', y='Company name', text='Company name', title='Companies per Responsible'), use_container_width=True)

        # 11️⃣ Download filtered CSV
        st.subheader("Download Filtered Data")
        csv = df_filtered.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download CSV", csv, "filtered_data.csv", "text/csv")

    else:
        st.error(f"❌ Missing columns in Excel: {missing_cols}")
else:
    st.info("Please upload an Excel file to start analysis.")
