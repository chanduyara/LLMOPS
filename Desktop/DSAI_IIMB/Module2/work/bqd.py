import streamlit as st
from google.cloud import bigquery
import pandas as pd
import altair as alt

# --- Page setup ---
st.set_page_config(page_title="Stack Overflow Dashboard", layout="wide")
st.title("📊 Top 10 Most Viewed Stack Overflow Questions (2008–2021)")

# --- BigQuery client ---
client = bigquery.Client()

# --- Query ---
query = """
SELECT
  q.id AS post_id,
  q.creation_date AS created,
  q.title AS post_title,
  q.view_count,
  u.id AS user_id,
  u.display_name AS owner_name,
  u.reputation,
  u.location
FROM `bigquery-public-data.stackoverflow.posts_questions` AS q
LEFT JOIN `bigquery-public-data.stackoverflow.users` AS u
  ON q.owner_user_id = u.id
WHERE q.creation_date BETWEEN '2008-01-01' AND '2021-12-31'
ORDER BY q.view_count DESC
LIMIT 10;
"""

# --- Run query and load DataFrame ---
rows = client.query(query).result()
df = pd.DataFrame([dict(row) for row in rows])
# df = client.query(query).to_dataframe()
df['created'] = pd.to_datetime(df['created'])
df['year'] = df['created'].dt.year

# --- Sidebar filters ---
st.sidebar.header("Filters")
year_range = st.sidebar.slider("Select Year Range", 2008, 2021, (2008, 2021))
min_reputation = st.sidebar.number_input("Minimum Reputation", value=0)

# Apply filters
filtered_df = df[(df['year'] >= year_range[0]) & (df['year'] <= year_range[1])]
filtered_df = filtered_df[filtered_df['reputation'] >= min_reputation]

# --- KPI Metrics ---
col1, col2, col3 = st.columns(3)
col1.metric("Total Views", f"{filtered_df['view_count'].sum():,}")
col2.metric("Avg Reputation", f"{filtered_df['reputation'].mean():.0f}")
col3.metric("Unique Years", filtered_df['year'].nunique())

# --- Tabs for organization ---
tab1, tab2, tab3, tab4 = st.tabs(["Yearly Trends", "Reputation Analysis", "Raw Data", "Download"])

with tab1:
    # Chart 1: Posts per Year
    year_counts = filtered_df['year'].value_counts().reset_index()
    year_counts.columns = ['year', 'count']
    chart1 = alt.Chart(year_counts).mark_bar(color='skyblue').encode(
        x=alt.X('year:O', title='Year'),
        y=alt.Y('count:Q', title='Number of Posts'),
        tooltip=['year', 'count']
    ).properties(title="Number of Top 10 Posts per Year")
    st.altair_chart(chart1, width="stretch")

    # Chart 2: Total Views per Year
    views_per_year = filtered_df.groupby('year')['view_count'].sum().reset_index()
    chart2 = alt.Chart(views_per_year).mark_bar(color='orange').encode(
        x=alt.X('year:O', title='Year'),
        y=alt.Y('view_count:Q', title='Total Views'),
        tooltip=['year', 'view_count']
    ).properties(title="Total Views of Top 10 Posts by Year")
    st.altair_chart(chart2, width="stretch")

with tab2:
    # Chart 3: Reputation vs Views
    chart3 = alt.Chart(filtered_df).mark_circle(size=200, color='green').encode(
        x=alt.X('reputation:Q', title='User Reputation'),
        y=alt.Y('view_count:Q', title='View Count'),
        tooltip=['post_title', 'owner_name', 'reputation', 'view_count']
    ).properties(title="Reputation vs View Count")
    st.altair_chart(chart3, width="stretch")

with tab3:
    st.subheader("Raw Query Results")
    st.dataframe(filtered_df)

with tab4:
    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV", csv, "stackoverflow_top10.csv", "text/csv")
