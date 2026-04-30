import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Sales Performance Dashboard", layout="wide")

st.markdown("""
<style>
.main { background-color: #0e1117; color: #e5e7eb; }
.block-container { padding-top: 1.5rem; }
.card {
    background: #161b22;
    padding: 22px;
    border-radius: 16px;
    border: 1px solid #30363d;
    box-shadow: 0 2px 12px rgba(0,0,0,0.6);
}
.metric-title { font-size: 13px; color: #9ca3af; text-transform: uppercase; }
.metric-value { font-size: 30px; font-weight: 600; color: #f9fafb; }
.section-title { font-size: 22px; font-weight: 700; margin-top: 25px; color: #f3f4f6; }
.section-subtitle { color: #9ca3af; margin-bottom: 15px; }
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_excel(file):
    df = pd.read_excel(file, sheet_name="RawData", engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    df.columns = df.columns.astype(str).str.strip()

    df["วันที่ออก"] = pd.to_datetime(df["วันที่ออก"], errors="coerce")
    df["มูลค่าก่อนภาษี"] = pd.to_numeric(df["มูลค่าก่อนภาษี"], errors="coerce").fillna(0)
    df = df.dropna(subset=["วันที่ออก"])
    df["Month"] = df["วันที่ออก"].dt.strftime("%Y-%m")

    for col in ["Category", "Sub-Category", "ชื่อลูกค้า", "เลขที่เอกสาร"]:
        df[col] = df[col].fillna("Unknown").astype(str).str.strip()

    return df


def money_short(value):
    if value >= 1_000_000:
        return f"฿{value/1_000_000:,.2f}M"
    if value >= 1_000:
        return f"฿{value/1_000:,.2f}K"
    return f"฿{value:,.2f}"


def style_chart(fig, height=420):
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=55, b=30),
        paper_bgcolor="#0e1117",
        plot_bgcolor="#0e1117",
        font=dict(size=12, color="#e5e7eb"),
        title=dict(font=dict(size=15, color="#f9fafb")),
        legend=dict(orientation="h", y=-0.25, font=dict(color="#d1d5db")),
    )
    fig.update_xaxes(showgrid=False, color="#9ca3af")
    fig.update_yaxes(gridcolor="#2a2f3a", color="#9ca3af")
    return fig


st.title("Sales Performance Dashboard")
st.caption("Upload Excel file and automatically generate dashboard")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Upload Excel to start")
    st.stop()

df = load_excel(uploaded_file)

col1, col2 = st.columns([1, 4])

with col1:
    category = st.selectbox("Category", ["All"] + sorted(df["Category"].unique()))

with col2:
    months = sorted(df["Month"].unique())
    selected_months = st.multiselect("Date Range", months, default=months)

filtered = df.copy()

if category != "All":
    filtered = filtered[filtered["Category"] == category]

if selected_months:
    filtered = filtered[filtered["Month"].isin(selected_months)]

total_sales = filtered["มูลค่าก่อนภาษี"].sum()
total_invoices = filtered["เลขที่เอกสาร"].replace("", None).dropna().nunique()
total_lines = len(filtered)
unique_customers = filtered["ชื่อลูกค้า"].nunique()
avg_transaction = total_sales / total_invoices if total_invoices else 0

m1, m2, m3, m4 = st.columns(4)

metrics = [
    ("Total Sales", money_short(total_sales)),
    ("Invoices", f"{total_invoices:,}"),
    ("Line Items", f"{total_lines:,}"),
    ("Customers", f"{unique_customers:,}")
]

for col, (title, value) in zip([m1, m2, m3, m4], metrics):
    with col:
        st.markdown(f"""
        <div class="card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
        </div>
        """, unsafe_allow_html=True)


st.markdown('<div class="section-title">Sales Overview</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">Monthly trends and category breakdown</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)

monthly = (
    filtered.groupby("Month", as_index=False)["มูลค่าก่อนภาษี"]
    .sum()
    .sort_values("Month")
)

fig1 = px.bar(
    monthly,
    x="Month",
    y="มูลค่าก่อนภาษี",
    title="Monthly Sales Trend",
    labels={"มูลค่าก่อนภาษี": "Sales", "Month": ""}
)
fig1 = style_chart(fig1)
c1.plotly_chart(fig1, use_container_width=True)

cat = (
    filtered.groupby("Category", as_index=False)["มูลค่าก่อนภาษี"]
    .sum()
    .sort_values("มูลค่าก่อนภาษี", ascending=False)
)

fig2 = px.pie(
    cat,
    names="Category",
    values="มูลค่าก่อนภาษี",
    title="Sales by Category"
)
fig2.update_traces(textinfo="percent+label")
fig2 = style_chart(fig2)
c2.plotly_chart(fig2, use_container_width=True)


st.markdown('<div class="section-title">Customer Analysis</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">Top customers and transaction details</div>', unsafe_allow_html=True)

top_cus = (
    filtered.groupby("ชื่อลูกค้า", as_index=False)["มูลค่าก่อนภาษี"]
    .sum()
    .sort_values("มูลค่าก่อนภาษี", ascending=False)
    .head(10)
)

fig3 = px.bar(
    top_cus,
    x="มูลค่าก่อนภาษี",
    y="ชื่อลูกค้า",
    orientation="h",
    title="Top 10 Customers by Sales",
    labels={"มูลค่าก่อนภาษี": "Sales", "ชื่อลูกค้า": ""}
)
fig3.update_layout(yaxis=dict(autorange="reversed"))
fig3 = style_chart(fig3, height=500)
st.plotly_chart(fig3, use_container_width=True)


customer_detail = (
    filtered.groupby("ชื่อลูกค้า", as_index=False)
    .agg(
        Transactions=("เลขที่เอกสาร", lambda x: x.replace("", None).dropna().nunique()),
        Total_Sales=("มูลค่าก่อนภาษี", "sum")
    )
)

customer_detail["Avg_Value"] = (
    customer_detail["Total_Sales"] / customer_detail["Transactions"]
).fillna(0)

customer_detail = customer_detail.sort_values("Total_Sales", ascending=False)

customer_detail = customer_detail.rename(columns={
    "ชื่อลูกค้า": "Customer Name",
    "Total_Sales": "Total Sales",
    "Avg_Value": "Avg Value"
})

st.markdown("### Customer Transaction Details")

st.dataframe(
    customer_detail,
    use_container_width=True,
    height=420,
    hide_index=True,
    column_config={
        "Customer Name": st.column_config.TextColumn("Customer Name"),
        "Transactions": st.column_config.NumberColumn("Transactions", format="%d"),
        "Total Sales": st.column_config.NumberColumn("Total Sales", format="฿%.2f"),
        "Avg Value": st.column_config.NumberColumn("Avg Value", format="฿%.2f"),
    }
)


st.markdown('<div class="section-title">Product Breakdown</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">Sub-category performance and monthly composition</div>', unsafe_allow_html=True)

top_sub = (
    filtered.groupby("Sub-Category", as_index=False)["มูลค่าก่อนภาษี"]
    .sum()
    .sort_values("มูลค่าก่อนภาษี", ascending=False)
    .head(10)
)

fig4 = px.bar(
    top_sub,
    x="Sub-Category",
    y="มูลค่าก่อนภาษี",
    title="Top 10 Sub-Categories by Sales",
    labels={"มูลค่าก่อนภาษี": "Sales", "Sub-Category": ""}
)
fig4.update_xaxes(tickangle=-30)
fig4 = style_chart(fig4, height=430)
st.plotly_chart(fig4, use_container_width=True)

monthly_cat = (
    filtered.groupby(["Month", "Category"], as_index=False)["มูลค่าก่อนภาษี"]
    .sum()
    .sort_values("Month")
)

fig5 = px.bar(
    monthly_cat,
    x="Month",
    y="มูลค่าก่อนภาษี",
    color="Category",
    title="Monthly Sales by Category (Stacked)",
    labels={"มูลค่าก่อนภาษี": "Sales", "Month": ""}
)
fig5 = style_chart(fig5, height=500)
st.plotly_chart(fig5, use_container_width=True)


with st.expander("View Raw Data"):
    st.dataframe(filtered, use_container_width=True, height=400)