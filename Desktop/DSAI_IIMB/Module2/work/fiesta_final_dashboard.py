"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         FIESTA GIFTS — STREAMLIT DASHBOARD                                 ║
║         IMB 761 | IIMB Case Study | Data-Driven Profitability              ║
║                                                                             ║
║  HOW TO RUN:                                                                ║
║    pip install streamlit plotly pandas openpyxl                             ║
║    streamlit run fiesta_dashboard_streamlit.py                              ║
║                                                                             ║
║  Place FiestaSales.xlsx in the same directory.                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import builtins
import math
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Fiesta Gifts Dashboard",
    page_icon="🎁",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY  = "#1B2B4B"
GOLD  = "#D4A843"
RED   = "#C0392B"
GREEN = "#1A6B3A"
TEAL  = "#0D7377"
PURP  = "#5B2D8E"
AMB   = "#B45309"
MID   = "#7A8BA8"
BG    = "#F4F6FB"

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #F4F6FB; }
    [data-testid="stSidebar"] { background-color: #1B2B4B; }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stRadio label { color: #D4A843 !important; font-weight: 600; }
    [data-testid="stSidebar"] .stMarkdown h2 { color: #D4A843 !important; }
    .kpi-card {
        background: white; border-radius: 10px; padding: 18px 20px;
        border-left: 5px solid #1B2B4B; box-shadow: 0 1px 6px rgba(0,0,0,0.07);
        margin-bottom: 8px;
    }
    .kpi-val  { font-size: 28px; font-weight: 700; line-height: 1.1; }
    .kpi-lbl  { font-size: 12px; color: #7A8BA8; margin-top: 4px; }
    .section-title {
        font-size: 18px; font-weight: 700; color: #1B2B4B;
        border-bottom: 2px solid #D4A843; padding-bottom: 4px; margin: 20px 0 14px;
    }
    .insight-box {
        background: #EAF1FA; border-radius: 8px; padding: 12px 16px;
        font-size: 13px; color: #1B2B4B; border-left: 4px solid #1B2B4B;
        margin-bottom: 10px;
    }
    .warn-box {
        background: #FEF3C7; border-radius: 8px; padding: 12px 16px;
        font-size: 13px; color: #633806; border-left: 4px solid #D4A843;
        margin-bottom: 10px;
    }
    .danger-box {
        background: #FEECEC; border-radius: 8px; padding: 12px 16px;
        font-size: 13px; color: #791F1F; border-left: 4px solid #C0392B;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ── Haversine (pure Python — safe from PySpark name collision) ────────────────
def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    la1,lo1,la2,lo2 = map(math.radians, [lat1,lon1,lat2,lon2])
    a = (math.sin((la2-la1)/2)**2
         + math.cos(la1)*math.cos(la2)*math.sin((lo2-lo1)/2)**2)
    return builtins.round(2*R*math.asin(math.sqrt(a)), 2)

DC_COORDS = {
    'Harrisburg':(40.27,-76.88),'Charlotte':(35.23,-80.84),
    'Chattanooga':(35.05,-85.31),'Fort Worth':(32.75,-97.33),
    'Jacksonville':(30.33,-81.66),'Kansas City':(39.10,-94.58),
    'Columbus':(39.96,-82.99),'Kenosha':(42.58,-87.82),
    'San Bernardino':(34.10,-117.29),'Mobile':(30.69,-88.04),
    'Sacramento':(38.58,-121.49),'Baltimore':(39.29,-76.61),
    'Fall River':(41.70,-71.15),'Nashua':(42.77,-71.46),
    'Phoenix':(33.45,-112.07),'Aurora':(39.73,-104.83),
    'Salt Lake City':(40.76,-111.89),'Bellevue':(47.61,-122.20),
    'Reno':(39.53,-119.81),'Hillsboro':(45.52,-122.99),
}

STATE_CENTROIDS = {
    'AL':(32.8,-86.8),'AZ':(34.3,-111.1),'CA':(36.8,-119.4),'CO':(39.1,-105.4),
    'FL':(27.8,-81.6),'GA':(32.7,-83.2),'IL':(40.0,-89.2),'IN':(39.8,-86.1),
    'KS':(38.5,-98.4),'KY':(37.5,-85.3),'LA':(31.1,-91.9),'ME':(45.4,-69.4),
    'MD':(39.0,-76.8),'MA':(42.2,-71.5),'MI':(44.3,-85.4),'MN':(46.4,-93.1),
    'MS':(32.7,-89.7),'MO':(38.5,-92.5),'NE':(41.5,-99.8),'NV':(39.3,-117.1),
    'NH':(43.7,-71.6),'NJ':(40.1,-74.6),'NY':(42.9,-75.7),'NC':(35.5,-79.5),
    'OH':(40.4,-82.8),'OK':(35.6,-96.9),'OR':(44.6,-122.1),'PA':(40.9,-77.8),
    'SC':(33.8,-80.9),'TN':(35.9,-86.7),'TX':(31.1,-97.6),'UT':(39.4,-111.1),
    'VA':(37.5,-78.5),'WA':(47.4,-121.5),'WI':(44.3,-89.7),
}

# ── Data loader ───────────────────────────────────────────────────────────────
@st.cache_data

def load_data():
    # Ensure the file path is correct relative to this script
    base_path = os.path.dirname(__file__)
    file_path = os.path.join(base_path, "FiestaSales.xlsx")

    # Check if file exists before reading
    if not os.path.exists(file_path):
        st.error(f"❌ File not found: {file_path}")
        st.stop()

    # Load and clean data
    df = pd.read_excel(file_path, sheet_name="OriginalData")
    df['SuppDCDist'] = pd.to_numeric(df['SuppDCDist'], errors='coerce')
    df['DCCustDist'] = pd.to_numeric(df['DCCustDist'], errors='coerce')
    df['PurchaseDate'] = pd.to_datetime(df['PurchaseDate'], errors='coerce')
    df['Month'] = df['PurchaseDate'].dt.to_period('M').astype(str)
# def load_data():
#     df = pd.read_excel("FiestaSales.xlsx", sheet_name="OriginalData")
#     df['SuppDCDist'] = pd.to_numeric(df['SuppDCDist'], errors='coerce')
#     df['DCCustDist'] = pd.to_numeric(df['DCCustDist'], errors='coerce')
#     df['PurchaseDate'] = pd.to_datetime(df['PurchaseDate'])
#     df['Month'] = df['PurchaseDate'].dt.to_period('M').astype(str)

    # Impute NA SuppDCDist
    def impute(row):
        if pd.notna(row['SuppDCDist']): return row['SuppDCDist']
        if row['SuppState'] not in STATE_CENTROIDS: return np.nan
        if row['DCCity']    not in DC_COORDS:       return np.nan
        sc  = STATE_CENTROIDS[row['SuppState']]
        dcc = DC_COORDS[row['DCCity']]
        return haversine(sc[0],sc[1],dcc[0],dcc[1])
    df['SuppDCDist'] = df.apply(impute, axis=1)

    total = df['TotalPrice'].sum()

    # SKU ABC
    sku = (df.groupby('SKUCode')
             .agg(spend=('TotalPrice','sum'), qty=('Quantity','sum'),
                  orders=('InvoiceNo','nunique'), suppliers=('SupplierID','nunique'),
                  item_type=('ItemType','first'), description=('Description','first'))
             .reset_index().sort_values('spend',ascending=False))
    sku['cum_pct'] = sku['spend'].cumsum()/total*100
    sku['ABC']     = sku['cum_pct'].apply(lambda c:'A' if c<=80 else('B' if c<=95 else 'C'))

    # Supplier ABC
    supp = (df.groupby('SupplierID')
              .agg(spend=('TotalPrice','sum'), orders=('InvoiceNo','nunique'),
                   skus=('SKUCode','nunique'))
              .reset_index().sort_values('spend',ascending=False))
    supp['cum_pct']  = supp['spend'].cumsum()/total*100
    supp['ABC']      = supp['cum_pct'].apply(lambda c:'A' if c<=80 else('B' if c<=95 else 'C'))
    supp['d3_action']= supp['spend'].apply(
        lambda s: 'EXIT' if s<500 else ('CONSOLIDATE' if s<2000 else 'KEEP'))

    # DC performance
    dc = (df.groupby(['DCCity','DCState'])
            .agg(spend=('TotalPrice','sum'), orders=('InvoiceNo','nunique'),
                 avg_supp_km=('SuppDCDist','mean'), avg_cust_km=('DCCustDist','mean'))
            .reset_index().sort_values('spend',ascending=False))
    dc['net']    = dc['spend'] - 100_000
    dc['status'] = dc['net'].apply(lambda n:'RETIRE' if n<0 else('MONITOR' if n<50000 else 'KEEP'))

    # Price variance
    pv = (df.groupby('SKUCode')
            .agg(qty=('Quantity','sum'), total=('TotalPrice','sum'),
                 min_p=('UnitPrice','min'), max_p=('UnitPrice','max'), avg_p=('UnitPrice','mean'))
            .reset_index().merge(sku[['SKUCode','ABC','description','item_type']],on='SKUCode'))
    pv['d8_saving']    = (pv['avg_p']-pv['min_p'])*pv['qty']
    pv['variance_pct'] = (pv['max_p']-pv['min_p'])/pv['min_p'].replace(0,np.nan)*100

    # Long haul
    lh = (df[df['SuppDCDist']>2000]
          .groupby(['DCCity','DCState'])
          .agg(rows=('TotalPrice','count'), spend=('TotalPrice','sum'),
               avg_km=('SuppDCDist','mean'))
          .reset_index().sort_values('spend',ascending=False))

    # Monthly
    monthly = df.groupby('Month')['TotalPrice'].sum().reset_index().sort_values('Month')

    # Clock problem
    clocks = df[df['ItemType'].str.lower()=='clock'].copy()
    clocks['clock_tier'] = clocks['UnitPrice'].apply(
        lambda p: 'Budget (<$6)' if p<6 else ('Mid ($6–15)' if p<15 else 'Premium (>$15)'))

    # Redirect map
    retire_list = dc[dc['status']=='RETIRE']['DCCity'].tolist()
    keep_list   = dc[dc['status']!='RETIRE']['DCCity'].tolist()
    redirect_map = {}
    redirect_rows = []
    for city in retire_list:
        if city not in DC_COORDS: continue
        coord = DC_COORDS[city]; best,bd = None,1e9
        for d in keep_list:
            if d not in DC_COORDS: continue
            dist = haversine(*coord,*DC_COORDS[d])
            if dist<bd: bd,best=dist,d
        redirect_map[city] = best
        r = dc[dc['DCCity']==city].iloc[0]
        redirect_rows.append({'Retired DC':city,'State':r['DCState'],
            'Annual Spend':builtins.round(r['spend']),'Net vs Lease':builtins.round(r['net']),
            'Redirect To':best,'Redirect km':builtins.round(bd)})
    redirect_df = pd.DataFrame(redirect_rows).sort_values('Net vs Lease')

    return df, total, sku, supp, dc, pv, lh, monthly, clocks, redirect_df, retire_list

# ── Load ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading Fiesta Gifts data…"):
    df, total, sku, supp, dc, pv, lh, monthly, clocks, redirect_df, retire_list = load_data()

n_retire  = len(retire_list)
NET_DC    = n_retire*100_000 + n_retire*15_000 - 8_000
D3_SAVING = (len(supp[supp['d3_action']=='EXIT'])*500
             + int(len(supp[supp['d3_action']=='CONSOLIDATE'])*500*0.67))
D8_SAVING = float(pv[pv['ABC']=='A']['d8_saving'].sum())
GRAND     = NET_DC + D3_SAVING + D8_SAVING

# ── Sidebar navigation ────────────────────────────────────────────────────────
st.sidebar.markdown("## 🎁 Fiesta Gifts")
st.sidebar.markdown("**IMB 761 · IIMB Case Study**")
st.sidebar.markdown("---")
page = st.sidebar.radio("Navigate", [
    "🏠 Overview",
    "📊 ABC Analysis",
    "🏭 Supplier Rationalisation",
    "🏢 DC Network",
    "💰 Price Optimisation",
    "⏱ Clock Problem",
    "📈 P&L Projection",
])
st.sidebar.markdown("---")
st.sidebar.markdown(f"**Total spend:** ${total:,.0f}")
st.sidebar.markdown(f"**Net saving:** ${GRAND:,.0f}")
st.sidebar.markdown(f"**P&L swing:** $1.41M")

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Overview":
    st.title("🎁 Fiesta Gifts — Data-Driven Road to Profitability")
    st.markdown("*IMB 761 · IIMB · Jayaram & Venkatagiri · 2019 · Spend Analysis Dashboard*")

    # KPI row 1
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    for col,(val,lbl,color) in zip([c1,c2,c3,c4,c5,c6],[
        ("$6.00M",  "Net sales FY2011",      RED),
        ("-$1.00M", "Net loss FY2011",        RED),
        ("50%",     "Gross margin",           GREEN),
        ("3,137",   "Active SKUs",            PURP),
        ("573",     "Active suppliers",       AMB),
        ("20",      "Distribution centres",   TEAL),
    ]):
        col.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color}">
          <div class="kpi-val" style="color:{color}">{val}</div>
          <div class="kpi-lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="section-title">💡 The Problem</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="danger-box">
    Fiesta Gifts generated <strong>$6M in sales</strong> in FY2011 but posted a
    <strong>$1M net loss</strong>. A 50% gross margin was entirely consumed by
    $4M in fixed indirect costs — led by a <strong>$2M DC lease bill</strong>
    across 20 flat-fee distribution centres. Investor Mahoney has given CEO
    Sanchez <strong>12 months</strong> to turn profitable — or close.
    </div>""", unsafe_allow_html=True)

    # Savings summary
    st.markdown('<div class="section-title">💰 Combined Savings from 8 Decisions</div>',
                unsafe_allow_html=True)

    wf_labels = ['Current loss','D7 DC closure','D3 Supplier exit',
                 'D3 Consolidation','D8 Price opt.','D1-D4 SKU OH','Net result']
    wf_values = [-1000, NET_DC/1000,
                 len(supp[supp['d3_action']=='EXIT'])*500/1000,
                 int(len(supp[supp['d3_action']=='CONSOLIDATE'])*500*0.67)/1000,
                 D8_SAVING/1000, len(sku[(sku['ABC']=='C')&(sku['spend']<140)])*15/1000, 0]
    wf_values[-1] = builtins.round(sum(wf_values[:-1]), 1)

    measures = ['absolute','relative','relative','relative','relative','relative','total']
    colors   = [RED, GREEN, GREEN, GREEN, GREEN, GREEN, TEAL]

    fig_wf = go.Figure(go.Waterfall(
        name="", measure=measures, x=wf_labels, y=wf_values,
        connector={"line":{"color":"#ccc","width":0.8}},
        decreasing={"marker":{"color":RED}},
        increasing={"marker":{"color":GREEN}},
        totals={"marker":{"color":TEAL}},
        text=[f"${abs(v):.0f}K" for v in wf_values],
        textposition="outside",
    ))
    fig_wf.update_layout(
        title="Savings waterfall — $K  (all 8 decisions)",
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial",size=12), height=380,
        yaxis_title="$K", showlegend=False,
        margin=dict(t=50,b=30,l=50,r=30),
    )
    st.plotly_chart(fig_wf, use_container_width=True)

    # KPI row 2 — savings
    c1,c2,c3,c4 = st.columns(4)
    for col,(val,lbl,color) in zip([c1,c2,c3,c4],[
        (f"${NET_DC:,.0f}",  "D7 DC closure saving",      GREEN),
        (f"${D3_SAVING:,}",  "D3 supplier rationalisation",GREEN),
        (f"${D8_SAVING:,.0f}","D8 price optimisation",    GREEN),
        (f"${GRAND:,.0f}",   "Total projected saving",     TEAL),
    ]):
        col.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color}">
          <div class="kpi-val" style="color:{color}">{val}</div>
          <div class="kpi-lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: ABC ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "📊 ABC Analysis":
    st.title("📊 ABC Analysis — SKU Portfolio")

    c1,c2,c3 = st.columns(3)
    for col,(cat,color,bg) in zip([c1,c2,c3],[
        ('A',GREEN,'#E8F5EE'),('B',AMB,'#FEF3C7'),('C',RED,'#FEECEC')]):
        s = sku[sku['ABC']==cat]
        col.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color};background:{bg}">
          <div class="kpi-val" style="color:{color}">Cat {cat}: {len(s):,} SKUs</div>
          <div class="kpi-lbl">${s['spend'].sum():,.0f} · {s['spend'].sum()/total*100:.0f}% of spend</div>
        </div>""", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    # Pareto bar
    with col1:
        fig = go.Figure()
        for cat,color in [('A',GREEN),('B',AMB),('C',RED)]:
            s = sku[sku['ABC']==cat]
            fig.add_trace(go.Bar(
                name=f"Cat {cat}", x=[cat],
                y=[s['spend'].sum()/1000],
                marker_color=color, text=[f"${s['spend'].sum()/1000:.0f}K"],
                textposition='auto',
            ))
        fig.update_layout(title="ABC spend distribution ($K)",
            plot_bgcolor="white",paper_bgcolor="white",
            font=dict(family="Arial",size=12),height=350,
            yaxis_title="Spend $K",showlegend=True,
            margin=dict(t=50,b=30))
        st.plotly_chart(fig, use_container_width=True)

    # Scatter: spend vs orders coloured by ABC
    with col2:
        fig2 = px.scatter(
            sku.sample(min(600,len(sku))), x='orders', y='spend',
            color='ABC', color_discrete_map={'A':GREEN,'B':AMB,'C':RED},
            hover_data=['SKUCode','description','item_type'],
            title="SKU spend vs orders (sample 600)",
            labels={'spend':'Annual spend ($)','orders':'Unique orders'},
        )
        fig2.update_layout(plot_bgcolor="white",paper_bgcolor="white",
            font=dict(family="Arial",size=12),height=350,
            margin=dict(t=50,b=30))
        st.plotly_chart(fig2, use_container_width=True)

    # Discard tiers
    st.markdown('<div class="section-title">D1 — Discard Tier Breakdown</div>',
                unsafe_allow_html=True)
    discard = sku[(sku['ABC']=='C')&(sku['spend']<140)].copy()
    discard['tier'] = discard['spend'].apply(
        lambda s: 'Tier 1: Dead (<$50)' if s<50 else
                  ('Tier 2: Low ($50–100)' if s<100 else 'Tier 3: Marginal ($100–140)'))
    tier_sum = discard.groupby('tier').agg(count=('SKUCode','count'),
                                            spend=('spend','sum')).reset_index()

    col1,col2 = st.columns(2)
    with col1:
        fig3 = px.bar(tier_sum, x='tier', y='count', color='tier',
            color_discrete_map={
                'Tier 1: Dead (<$50)':RED,
                'Tier 2: Low ($50–100)':AMB,
                'Tier 3: Marginal ($100–140)':GOLD,
            },
            text='count', title="Discard SKUs by tier (count)",
            labels={'count':'SKU count','tier':'Tier'})
        fig3.update_layout(plot_bgcolor="white",paper_bgcolor="white",
            showlegend=False,height=300,margin=dict(t=50,b=30))
        st.plotly_chart(fig3, use_container_width=True)
    with col2:
        st.dataframe(
            tier_sum.rename(columns={'tier':'Tier','count':'SKUs','spend':'Total Spend ($)'}),
            use_container_width=True, hide_index=True)
        st.markdown(f"""
        <div class="insight-box">
        <strong>{len(discard):,} SKUs</strong> to discard — {len(discard)/len(sku)*100:.0f}%
        of the catalogue generating only <strong>${discard['spend'].sum():,.0f}</strong>
        (1.93% of total spend). Overhead saving: <strong>${len(discard)*15:,}/yr</strong>.
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-title">Top 20 Category A SKUs by Spend</div>',
                unsafe_allow_html=True)
    top_a = sku[sku['ABC']=='A'].head(20).copy()
    fig4 = px.bar(top_a.sort_values('spend'), x='spend', y='SKUCode',
        orientation='h', color='spend',
        color_continuous_scale=[[0,GREEN],[1,NAVY]],
        hover_data=['description','item_type'],
        labels={'spend':'Annual spend ($)','SKUCode':'SKU Code'},
        title="Top 20 Category A SKUs")
    fig4.update_layout(plot_bgcolor="white",paper_bgcolor="white",
        font=dict(family="Arial",size=11),height=500,
        coloraxis_showscale=False,margin=dict(t=50,b=30))
    st.plotly_chart(fig4, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: SUPPLIER RATIONALISATION
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🏭 Supplier Rationalisation":
    st.title("🏭 Supplier Rationalisation — Decision 3")

    c1,c2,c3,c4 = st.columns(4)
    for col,(val,lbl,color) in zip([c1,c2,c3,c4],[
        ("573", "Total suppliers",           NAVY),
        (str(len(supp[supp['d3_action']=='EXIT'])),   "Exit (<$500 spend)",  RED),
        (str(len(supp[supp['d3_action']=='CONSOLIDATE'])),"Consolidate ($500–2K)",AMB),
        (f"${D3_SAVING:,}", "Total D3 saving/yr",   GREEN),
    ]):
        col.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color}">
          <div class="kpi-val" style="color:{color}">{val}</div>
          <div class="kpi-lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        d3_counts = supp['d3_action'].value_counts().reset_index()
        d3_counts.columns = ['Action','Count']
        fig = px.pie(d3_counts, names='Action', values='Count',
            color='Action',
            color_discrete_map={'KEEP':GREEN,'CONSOLIDATE':AMB,'EXIT':RED},
            title="D3 — Supplier action breakdown",
            hole=0.4)
        fig.update_layout(height=360,font=dict(family="Arial",size=12),
            plot_bgcolor="white",paper_bgcolor="white",margin=dict(t=60,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.histogram(supp[supp['spend']<5000], x='spend', nbins=40,
            color_discrete_sequence=[NAVY],
            title="Supplier spend distribution (< $5K shown)",
            labels={'spend':'Annual spend ($)','count':'No. of suppliers'})
        fig2.add_vline(x=500, line_dash="dash", line_color=RED,
                       annotation_text="Exit threshold $500",
                       annotation_position="top right")
        fig2.add_vline(x=2000, line_dash="dash", line_color=AMB,
                       annotation_text="Consolidate threshold $2K",
                       annotation_position="top right")
        fig2.update_layout(plot_bgcolor="white",paper_bgcolor="white",
            height=360,font=dict(family="Arial",size=12),margin=dict(t=60,b=30))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown(f"""
    <div class="insight-box">
    At <strong>$500 overhead per supplier per year</strong>, the 167 exit-tier suppliers
    each cost more to manage than they contribute in spend. Terminating them saves
    <strong>${len(supp[supp['d3_action']=='EXIT'])*500:,}/yr</strong> in procurement overhead
    with zero revenue impact — their SKUs redirect to existing Cat A/B suppliers in the same
    item type.
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-title">Top 20 Suppliers by Spend</div>',
                unsafe_allow_html=True)
    top_s = supp.head(20).copy()
    fig3 = px.bar(top_s.sort_values('spend'), x='spend', y='SupplierID',
        orientation='h', color='ABC',
        color_discrete_map={'A':GREEN,'B':AMB,'C':RED},
        hover_data=['orders','skus'],
        labels={'spend':'Annual spend ($)','SupplierID':'Supplier ID'},
        title="Top 20 suppliers (coloured by ABC)")
    fig3.update_layout(plot_bgcolor="white",paper_bgcolor="white",
        height=480,font=dict(family="Arial",size=11),margin=dict(t=50,b=30))
    st.plotly_chart(fig3, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: DC NETWORK
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🏢 DC Network":
    st.title("🏢 Distribution Centre Network — Decision 7")

    c1,c2,c3,c4 = st.columns(4)
    for col,(val,lbl,color) in zip([c1,c2,c3,c4],[
        ("20",        "Total DCs",           NAVY),
        (str(n_retire),"DCs to retire",      RED),
        (f"${NET_DC:,.0f}","Net D7 saving",  GREEN),
        ("91.2%",     "of $1M loss covered", TEAL),
    ]):
        col.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color}">
          <div class="kpi-val" style="color:{color}">{val}</div>
          <div class="kpi-lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-title">DC Spend vs $100K Lease Threshold</div>',
                unsafe_allow_html=True)
    dc_sorted = dc.sort_values('spend', ascending=True)
    bar_colors = [RED if s=='RETIRE' else (AMB if s=='MONITOR' else GREEN)
                  for s in dc_sorted['status']]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=dc_sorted['spend']/1000,
        y=dc_sorted['DCCity']+', '+dc_sorted['DCState'],
        orientation='h',
        marker_color=bar_colors,
        hovertemplate="<b>%{y}</b><br>Spend: $%{x:.0f}K<extra></extra>",
        text=[f"${v/1000:.0f}K" for v in dc_sorted['spend']],
        textposition='outside',
    ))
    fig.add_vline(x=100, line_dash="dash", line_color=RED, line_width=1.5,
                  annotation_text="$100K lease", annotation_position="top right")
    fig.update_layout(
        title="Annual DC spend vs $100K fixed lease (red=retire, amber=monitor, green=keep)",
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial",size=11), height=520,
        xaxis_title="Annual spend ($K)", showlegend=False,
        margin=dict(t=50,b=30,l=180,r=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    col1,col2 = st.columns([3,2])
    with col1:
        st.markdown('<div class="section-title">Redirect Policy</div>',
                    unsafe_allow_html=True)
        st.dataframe(redirect_df, use_container_width=True, hide_index=True)
        st.markdown("""
        <div class="warn-box">
        ⚠️ <strong>Sacramento absorbs 4 of 8 closures (+63% load)</strong> — 
        verify 3PL capacity contract before closing all eight simultaneously.
        Recommended: close in two waves.
        </div>""", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="section-title">Long-Haul Flagged Spend</div>',
                    unsafe_allow_html=True)
        lh_top = lh.head(10).sort_values('spend', ascending=True)
        fig2 = go.Figure(go.Bar(
            x=lh_top['spend']/1000,
            y=lh_top['DCCity']+', '+lh_top['DCState'],
            orientation='h',
            marker_color=RED,
            text=[f"${v/1000:.0f}K" for v in lh_top['spend']],
            textposition='outside',
        ))
        fig2.update_layout(
            title="Long-haul spend by DC\n(supplier→DC > 2,000 km)",
            plot_bgcolor="white",paper_bgcolor="white",
            font=dict(family="Arial",size=11),height=360,
            xaxis_title="Spend $K",showlegend=False,
            margin=dict(t=60,b=30,l=160,r=60),
        )
        st.plotly_chart(fig2, use_container_width=True)

    # DC map
    st.markdown('<div class="section-title">DC Location Map</div>',
                unsafe_allow_html=True)
    dc_map = dc.copy()
    dc_map['lat'] = dc_map['DCCity'].map(
        {k:v[0] for k,v in DC_COORDS.items()})
    dc_map['lon'] = dc_map['DCCity'].map(
        {k:v[1] for k,v in DC_COORDS.items()})
    dc_map = dc_map.dropna(subset=['lat','lon'])
    dc_map['color_label'] = dc_map['status'].map(
        {'RETIRE':'Retire','MONITOR':'Monitor','KEEP':'Keep'})
    fig3 = px.scatter_geo(dc_map, lat='lat', lon='lon',
        color='color_label',
        color_discrete_map={'Retire':RED,'Monitor':AMB,'Keep':GREEN},
        size='spend', size_max=30,
        hover_name='DCCity',
        hover_data={'lat':False,'lon':False,'spend':True,'net':True,'status':True},
        scope='usa',
        title="DC locations (size = annual spend, colour = decision)",
    )
    fig3.update_layout(height=420, font=dict(family="Arial",size=12),
                       margin=dict(t=60,b=20))
    st.plotly_chart(fig3, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: PRICE OPTIMISATION
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "💰 Price Optimisation":
    st.title("💰 Price Optimisation — Decision 8")

    c1,c2,c3 = st.columns(3)
    for col,(val,lbl,color) in zip([c1,c2,c3],[
        (f"${D8_SAVING:,.0f}", "Cat A saving at min price",     GREEN),
        ("9%",                  "Avg overpayment vs min price",   RED),
        ("696",                 "Category A SKUs affected",       NAVY),
    ]):
        col.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color}">
          <div class="kpi-val" style="color:{color}">{val}</div>
          <div class="kpi-lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    pv_a = pv[pv['ABC']=='A'].sort_values('d8_saving', ascending=False)

    col1,col2 = st.columns(2)
    with col1:
        top12 = pv_a.head(12).sort_values('d8_saving', ascending=True)
        fig = go.Figure(go.Bar(
            x=top12['d8_saving'], y=top12['SKUCode'],
            orientation='h', marker_color=GREEN,
            text=[f"${v:,.0f}" for v in top12['d8_saving']],
            textposition='outside',
            hovertemplate="<b>%{y}</b><br>%{customdata}<br>Saving: $%{x:,.0f}<extra></extra>",
            customdata=top12['description'],
        ))
        fig.update_layout(title="Top 12 Cat A SKUs — D8 saving",
            plot_bgcolor="white",paper_bgcolor="white",
            font=dict(family="Arial",size=11),height=420,
            xaxis_title="Annual saving ($)",
            margin=dict(t=50,b=30,l=80,r=80))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.scatter(pv_a.head(80), x='variance_pct', y='d8_saving',
            size='total', color='item_type',
            hover_data=['SKUCode','description','min_p','avg_p'],
            title="Price variance % vs saving (Cat A, top 80)",
            labels={'variance_pct':'Price range %','d8_saving':'Potential saving ($)',
                    'total':'Total spend'},
        )
        fig2.update_layout(plot_bgcolor="white",paper_bgcolor="white",
            font=dict(family="Arial",size=11),height=420,
            margin=dict(t=50,b=30),showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    # SKU 22423 monthly demand
    st.markdown('<div class="section-title">SKU 22423 — Regency Cakestand Monthly Demand</div>',
                unsafe_allow_html=True)
    sku22 = df[df['SKUCode']=='22423'].groupby('Month').agg(
        qty=('Quantity','sum'), min_p=('UnitPrice','min'),
        max_p=('UnitPrice','max'), avg_p=('UnitPrice','mean')).reset_index()
    sku22 = sku22.sort_values('Month')

    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
    fig3.add_trace(go.Bar(x=sku22['Month'], y=sku22['qty'],
        name="Qty ordered", marker_color=NAVY, opacity=0.75), secondary_y=False)
    fig3.add_trace(go.Scatter(x=sku22['Month'], y=sku22['avg_p'],
        name="Avg price $", mode='lines+markers',
        line=dict(color=RED,width=2), marker=dict(size=6)), secondary_y=True)
    fig3.add_trace(go.Scatter(x=sku22['Month'], y=sku22['min_p'],
        name="Min price $", mode='lines+markers',
        line=dict(color=GREEN,width=2,dash='dash'),
        marker=dict(size=5,symbol='diamond')), secondary_y=True)
    fig3.update_layout(
        title="SKU 22423 — December spike (917 units) vs avg 213/month · Pre-buy at min price $15.33",
        plot_bgcolor="white",paper_bgcolor="white",
        font=dict(family="Arial",size=11),height=400,
        legend=dict(x=0.01,y=0.99),margin=dict(t=60,b=40))
    fig3.update_yaxes(title_text="Quantity", secondary_y=False)
    fig3.update_yaxes(title_text="Unit price ($)", secondary_y=True)
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("""
    <div class="insight-box">
    <strong>Decision 8:</strong> Pre-purchase December inventory in September/October at the
    minimum observed price ($15.33). December demand for SKU 22423 is <strong>917 units</strong>
    vs a monthly average of 213 — a predictable structural spike. Buying at the December peak
    price of $17.85 adds $2.52 per unit in avoidable cost. Scaling to all 696 Cat A SKUs
    at their minimum prices: <strong>$353,000/yr saved permanently</strong>.
    </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: CLOCK PROBLEM
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "⏱ Clock Problem":
    st.title("⏱ Clock Item Type Problem — Part a")

    col1,col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class="danger-box">
        <strong>Problem 1 — Price heterogeneity:</strong><br>
        The 'clock' item type contains SKUs priced from
        <strong>$0.27 to $35.00</strong> — a 12,900% price range.
        Budget keyrings and premium alarm clocks share one category,
        making ABC analysis and price averages meaningless.
        </div>
        <div class="danger-box" style="margin-top:10px">
        <strong>Problem 2 — Cross-country misrouting:</strong><br>
        Maine suppliers shipped clocks to California DCs
        (avg <strong>6,109 km</strong>) when Kansas suppliers were
        available 400 km away from the Kansas City DC.
        </div>
        <div class="insight-box" style="margin-top:10px">
        <strong>Fix:</strong> Sub-tier into Budget (&lt;$6),
        Mid ($6–15), Premium (&gt;$15). Apply ABC separately per tier.
        Audit all clock routes exceeding 1,500 km and reassign.
        </div>""", unsafe_allow_html=True)

    with col2:
        tier_sum = clocks.groupby('clock_tier').agg(
            skus=('SKUCode','nunique'), spend=('TotalPrice','sum'),
            min_p=('UnitPrice','min'), max_p=('UnitPrice','max')
        ).reset_index()
        fig = px.bar(tier_sum, x='clock_tier', y='spend',
            color='clock_tier',
            color_discrete_map={
                'Budget (<$6)':GOLD,'Mid ($6–15)':AMB,'Premium (>$15)':RED},
            text=tier_sum['spend'].apply(lambda v:f"${v:,.0f}"),
            title="Clock spend by sub-tier (after fix)",
            labels={'spend':'Total spend ($)','clock_tier':'Sub-tier'})
        fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",
            height=300,showlegend=False,font=dict(family="Arial",size=12),
            margin=dict(t=50,b=30))
        st.plotly_chart(fig, use_container_width=True)

    # Cross-country route heatmap
    st.markdown('<div class="section-title">Clock Shipment Routes — Supplier State → DC State</div>',
                unsafe_allow_html=True)
    cross = (clocks.groupby(['SuppState','DCState'])
                   .agg(rows=('TotalPrice','count'), avg_km=('SuppDCDist','mean'))
                   .reset_index().dropna(subset=['avg_km']))
    pivot = cross.pivot(index='SuppState', columns='DCState', values='rows').fillna(0)
    fig2 = px.imshow(pivot, color_continuous_scale='RdYlGn_r',
        title="Clock shipment count: supplier state (rows) → DC state (cols)",
        labels=dict(color="Shipments"),
        text_auto=True)
    fig2.update_layout(height=420,font=dict(family="Arial",size=10),
                       margin=dict(t=60,b=30))
    st.plotly_chart(fig2, use_container_width=True)

    # Worst routes
    worst = (cross.sort_values('avg_km', ascending=False).head(10)
             .rename(columns={'SuppState':'Supplier State','DCState':'DC State',
                               'rows':'Shipments','avg_km':'Avg km'}))
    st.markdown('<div class="section-title">Top 10 Longest Clock Routes</div>',
                unsafe_allow_html=True)
    st.dataframe(worst, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: P&L PROJECTION
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "📈 P&L Projection":
    st.title("📈 FY2012 Optimised P&L Projection")

    st.markdown('<div class="section-title">Monthly Spend Trend FY2011</div>',
                unsafe_allow_html=True)
    fig_mo = go.Figure()
    fig_mo.add_trace(go.Scatter(
        x=monthly['Month'], y=monthly['TotalPrice']/1000,
        fill='tozeroy', fillcolor='rgba(13,115,119,0.12)',
        line=dict(color=TEAL,width=2.5),
        mode='lines+markers', marker=dict(size=6),
        name="Monthly spend $K"))
    fig_mo.update_layout(
        title="Monthly spend FY2011 ($K) — December holiday spike visible",
        plot_bgcolor="white",paper_bgcolor="white",
        font=dict(family="Arial",size=12),height=300,
        yaxis_title="Spend $K",margin=dict(t=50,b=40))
    st.plotly_chart(fig_mo, use_container_width=True)

    # P&L table
    st.markdown('<div class="section-title">FY2011 Actual vs FY2012 Target ($K)</div>',
                unsafe_allow_html=True)
    
    proj_net = builtins.round((GRAND - 1_000_000)/1000, 0)
    pl_df = pd.DataFrame([
        {"Line Item":"Net Sales",       "FY2011":"$6,000K","FY2012":"$6,300K","Change":"+$300K"},
        {"Line Item":"COGS",            "FY2011":"-$3,000K","FY2012":"-$2,781K","Change":"-$219K ✓ D8"},
        {"Line Item":"Gross Income",    "FY2011":"$3,000K","FY2012":"$3,519K","Change":"+$519K"},
        {"Line Item":"Salaries",        "FY2011":"-$1,200K","FY2012":"-$1,200K","Change":"—"},
        {"Line Item":"SG&A",            "FY2011":"-$600K","FY2012":"-$570K","Change":"-$30K"},
        {"Line Item":"DC Lease Fees",   "FY2011":"-$2,000K","FY2012":"-$1,188K","Change":f"-$912K ✓ D7"},
        {"Line Item":"Supplier Overhead","FY2011":"—","FY2012":f"+${D3_SAVING//1000}K","Change":f"-${D3_SAVING//1000}K ✓ D3"},
        {"Line Item":"Interest",        "FY2011":"-$200K","FY2012":"-$200K","Change":"0K"},
        {"Line Item":"Total Indirect",  "FY2011":"-$4,000K","FY2012":"-$3,266K","Change":"-$734K"},
        {"Line Item":"NET INCOME",      "FY2011":"-$1,000K",
         "FY2012":f"+${proj_net:.0f}K","Change":f"${GRAND/1000:+.0f}K swing"},
    ])
    st.dataframe(pl_df, use_container_width=True, hide_index=True)

    # Savings breakdown donut
    col1,col2 = st.columns(2)
    with col1:
        labels_d = ['D7 DC closure','D3 Supplier exit/consolidation','D8 Price opt.','D1-D4 SKU OH']
        d1_d4 = len(sku[(sku['ABC']=='C')&(sku['spend']<140)])*15
        vals_d = [NET_DC, D3_SAVING, D8_SAVING, d1_d4]
        fig_d = go.Figure(go.Pie(
            labels=labels_d, values=vals_d,
            hole=0.45,
            marker_colors=[RED,AMB,GREEN,TEAL],
            textinfo='label+percent',
            hovertemplate="<b>%{label}</b><br>$%{value:,.0f}/yr<extra></extra>",
        ))
        fig_d.update_layout(
            title=f"Saving breakdown (total ${GRAND:,.0f}/yr)",
            font=dict(family="Arial",size=11),height=380,
            margin=dict(t=60,b=10))
        st.plotly_chart(fig_d, use_container_width=True)

    with col2:
        st.markdown("""
        <div class="section-title">12-Month Implementation Roadmap</div>""",
        unsafe_allow_html=True)
        roadmap = pd.DataFrame([
            {"Quarter":"Q1 (Jan–Mar)","Decision":"D7","Action":"Close 6–8 DCs — issue termination notices","Saving":"$912K/yr"},
            {"Quarter":"Q1 (Jan–Mar)","Decision":"D3","Action":"Terminate 167 exit-tier suppliers","Saving":"$83.5K/yr"},
            {"Quarter":"Q2 (Apr–Jun)","Decision":"D1-D4","Action":"Delist 1,097 discard SKUs","Saving":"$16.5K/yr"},
            {"Quarter":"Q2 (Apr–Jun)","Decision":"D3","Action":"Consolidate 128 mid-tier suppliers","Saving":"$42.9K/yr"},
            {"Quarter":"Q3 (Jul–Sep)","Decision":"D5-D6","Action":"Audit long-haul routes; onboard closer suppliers","Saving":"Network"},
            {"Quarter":"Q3 (Jul–Sep)","Decision":"D8","Action":"Negotiate Cat A min-price evergreen contracts","Saving":"$219K+/yr"},
            {"Quarter":"Q4 (Oct–Dec)","Decision":"D8","Action":"Pre-buy December inventory at min price","Saving":"First holiday saving"},
        ])
        st.dataframe(roadmap, use_container_width=True, hide_index=True)

        proj = builtins.round((GRAND - 1_000_000))
        st.markdown(f"""
        <div class="insight-box" style="margin-top:12px">
        <strong>Projected FY2012 net income:</strong>
        <span style="color:{GREEN};font-weight:700;font-size:20px">
        ${proj:+,}</span><br>
        A swing of <strong>${GRAND:,.0f}</strong> from the FY2011
        <span style="color:{RED}">-$1,000,000</span> loss.
        Mahoney's 12-month deadline is achievable.
        </div>""", unsafe_allow_html=True)
