# 1. Kutubxonalarni import qilamiz
import pandas as pd
import streamlit as st
import plotly.express as px

# 2. Data yuklash
# Sizning Excel fayl nomi: 'company_data.xlsx'
df = pd.read_excel('company_data.xlsx')

# 3. Sana ustunlarini datetime formatga o'tkazish
df['Date of creation'] = pd.to_datetime(df['Date of creation'], format='%d.%m.%Y %H:%M:%S')
df['Date modified'] = pd.to_datetime(df['Date modified'], format='%d.%m.%Y %H:%M:%S')

# 4. Streamlit dashboard
st.set_page_config(page_title="Company Data Analysis", layout="wide")

st.title("📊 Company Data Dashboard")
st.markdown("**Analysis of Stage, Responsible, Dates, and Company Names**")

# 5. Filterlar
st.sidebar.header("Filters")
unique_stage = df['Stage'].unique()
selected_stage = st.sidebar.multiselect("Select Stage", unique_stage, default=unique_stage)

unique_responsible = df['Responsible'].unique()
selected_responsible = st.sidebar.multiselect("Select Responsible", unique_responsible, default=unique_responsible)

# Filterlangan data
df_filtered = df[(df['Stage'].isin(selected_stage)) & (df['Responsible'].isin(selected_responsible))]

st.subheader("Filtered Data")
st.dataframe(df_filtered)

# 6. Summary statistikalar
st.subheader("Summary Statistics")
col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Total Companies", df_filtered['Company name'].nunique())
with col2:
    st.metric("Total Records", df_filtered.shape[0])
with col3:
    st.metric("Total Responsible", df_filtered['Responsible'].nunique())

# 7. Dumaloq diagramlar (Stage, Responsible)
st.subheader("Stage Distribution")
fig_stage = px.pie(df_filtered, names='Stage', title='Stage Distribution')
st.plotly_chart(fig_stage, use_container_width=True)

st.subheader("Responsible Distribution")
fig_resp = px.pie(df_filtered, names='Responsible', title='Responsible Distribution')
st.plotly_chart(fig_resp, use_container_width=True)

# 8. Vaqt bo'yicha trend (Date of creation)
st.subheader("Records Over Time")
df_filtered['Creation Date'] = df_filtered['Date of creation'].dt.date
df_time = df_filtered.groupby('Creation Date').size().reset_index(name='Count')
fig_time = px.line(df_time, x='Creation Date', y='Count', title='Records Created Over Time')
st.plotly_chart(fig_time, use_container_width=True)

# 9. Responsible bo'yicha analiz (Company soni)
st.subheader("Companies Managed by Responsible")
df_resp_count = df_filtered.groupby('Responsible')['Company name'].nunique().reset_index().sort_values(by='Company name', ascending=False)
fig_resp_count = px.bar(df_resp_count, x='Responsible', y='Company name', text='Company name', title='Companies per Responsible')
st.plotly_chart(fig_resp_count, use_container_width=True)
