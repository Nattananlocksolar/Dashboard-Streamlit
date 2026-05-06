import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Format Excel", layout="wide")

st.title("Format Excel from PEAK")
st.caption("Upload raw Excel → Clean / Format → Download cleaned Excel")

BASE_COLUMNS = [
    "เลขที่เอกสาร",
    "วันที่ออก",
    "สถานะ",
    "ชื่อลูกค้า",
    "รหัสสินค้า/บริการ",
    "ชื่อสินค้า/บริการ",
    "จำนวน",
    "ราคา",
    "มูลค่าก่อนภาษี",
]

FINAL_COLUMNS = BASE_COLUMNS + ["Category", "Sub-Category"]

code_map = {
    "101": ("Rooftop", "Nut-Screw"),
    "102": ("Rooftop", "Nut-Screw"),
    "103": ("Rooftop", "Hanger Bolt"),
    "104": ("Rooftop", "EPDM"),
    "105": ("Rooftop", "Nut-Screw"),
    "106": ("Rooftop", "Nut-Screw"),
    "107": ("Rooftop", "Clip Plate"),
    "108": ("Rooftop", "Cable Clip"),
    "109": ("Rooftop", "Cable Tile"),
    "110": ("Rooftop", "Tek bit"),
    "111": ("Rooftop", "Pan Tile Bracket"),
    "115": ("Rooftop", "Packing Box"),
    "116": ("Rooftop", "Round PIP"),
    "117": ("Rooftop", "Plate"),

    "151": ("Rooftop", "Nut-Screw"),
    "153": ("Rooftop", "Hanger Bolt"),
    "156": ("Rooftop", "Nut-Screw"),
    "157": ("Rooftop", "Clip Plate"),
    "158": ("Rooftop", "Cable Clip"),

    "201": ("Rooftop", "Mid Clamp"),
    "202": ("Rooftop", "End Clamp"),
    "203": ("Rooftop", "Kilplok"),
    "204": ("Rooftop", "Trapezoidal"),
    "205": ("Rooftop", "L-Bracket"),
    "206": ("Rooftop", "Tilt Kit"),
    "207": ("Rooftop", "Spacer"),
    "208": ("Rooftop", "Rail Nut"),
    "209": ("Rooftop", "Rail Nut"),
    "210": ("Rooftop", "Earthing Lug"),
    "211": ("Rooftop", "Pan-Tile Bracket"),
    "212": ("Solar Farm", "Terramount"),
    "213": ("Solar Farm", "Terramount"),
    "214": ("Rooftop", "Klip Lok 750"),
    "215": ("Rooftop", "Klip Lok V600"),
    "217": ("Rooftop", "Klip Lok 750"),
    "218": ("Rooftop", "Cable Ladder and Perforated Tray"),
    "219": ("Rooftop", "Guard Rail"),
    "220": ("Inverter", "Combiner Box Accessorie"),
    "233": ("Floating", "Floating"),
    "250": ("Rooftop", "Clamp"),
    "251": ("Rooftop", "Rail Nut"),
    "252": ("Rooftop", "End Clamp"),
    "266": ("Floating", "Floating"),

    "300": ("Rooftop", "Rail"),
    "301": ("Rooftop", "Mid Clamp"),
    "302": ("Rooftop", "End Clamp"),
    "303": ("Rooftop", "Kliplok Bracket"),
    "304": ("Rooftop", "Trapezoidal Bracket"),
    "305": ("Rooftop", "L-Bracket"),
    "306": ("Rooftop", "Tilt Kit"),
    "307": ("Rooftop", "Rail Joiner"),
    "308": ("Rooftop", "Rail to Bracket"),
    "309": ("Rooftop", "Earthing Lug"),
    "310": ("Rooftop", "Cross Brace bracket"),
    "312": ("Solar Farm", "TerraMount and DuoPost, TowerPost"),
    "313": ("Rooftop", "Hanger Bolt"),
    "314": ("Rooftop", "Pan Tile"),
    "315": ("Rooftop", "Klip lok V750"),
    "316": ("Rooftop", "Klip lok V600"),
    "318": ("Rooftop", "Guardrail"),
    "332": ("Floating", "Floating"),
    "333": ("Floating", "Floating"),
    "335": ("Inverter", "Inverter"),
    "336": ("Solar Panel", "Solar Panel"),
    "341": ("Floating", "Floating Testing Product"),
    "350": ("Rooftop", "Rail"),
    "351": ("Rooftop", "Mid Clamp"),
    "352": ("Rooftop", "End Clamp"),
    "353": ("Rooftop", "Kliplok Bracket"),
    "355": ("Rooftop", "L-Bracket"),
    "356": ("Rooftop", "Angle Tripod"),
    "357": ("Rooftop", "Rail Joiner Assy"),
    "358": ("Rooftop", "Rail to Bracket Connector"),
    "359": ("Rooftop", "Earthing Lug"),
    "360": ("Rooftop", "Tile hook"),
    "361": ("Rooftop", "Rail Clamp Kit"),
    "362": ("Floating", "Spreader Bar"),
    "363": ("Carport", "Carport"),
    "364": ("Rooftop", "Hanger Bolt"),
    "366": ("Floating", "Floating"),
    "367": ("Floating", "Floating"),
    "368": ("Solar Farm", "S-Plie"),
    "369": ("Rooftop", "Adjustable Foot"),

    "400": ("Rooftop", "Adj.Tilt Kit"),
    "401": ("Rooftop", "Mock Up"),
    "402": ("Rooftop", "LF Roof"),

    "500": ("Solar Farm", "TerraMount"),
    "501": ("Solar Carport", "Solar Carport"),
    "502": ("Walkway", "Walkway HDG"),
    "503": ("Floating", "Floating"),
    "509": ("Floating", "Floating"),
    "510": ("Rooftop", "Guardrail"),

    "601": ("Solar Panel", "Solar Panel"),
    "602": ("Solar Panel", "Solar Panel"),
    "603": ("Solar Panel", "Solar Panel"),
    "604": ("Solar Panel", "Solar Panel"),
    "605": ("Solar Panel", "Solar Panel"),
    "606": ("Solar Panel", "Solar Panel"),
    "607": ("Solar Panel", "Solar Panel"),
    "611": ("Floating", "Spreader Bar"),
    "621": ("Inverter", "Inverter"),
    "622": ("Inverter", "Smart Meter"),
    "623": ("Inverter", "Dongle / Smart Logger"),
    "624": ("Inverter", "CT Export"),
    "625": ("Inverter", "TOU"),
    "626": ("Inverter", "Micro Inverter Accessorie"),
    "627": ("Inverter", "Rapid Shutdown"),
    "628": ("Inverter", "Emergency Set"),
    "629": ("Inverter", "HV Junction Box"),
    "630": ("Inverter", "Accessorie"),
    "631": ("Inverter", "Battery"),
    "632": ("Inverter", "EV Charger"),
    "633": ("Inverter", "ArcBox"),
    "642": ("Boxset", "ชุดเซ็ท"),
    "643": ("Boxset", "ชุดเซ็ท"),
    "651": ("Inverter", "สายไฟ Solar"),
    "652": ("Inverter", "MC4"),
    "661": ("Inverter", "DC/AC Combiner Box"),

    "TS": ("Transport", "Transport"),
}


def find_header_row(raw_df):
    best_row = None
    best_score = 0

    for i in range(min(30, len(raw_df))):
        row_values = raw_df.iloc[i].astype(str).str.strip().tolist()
        score = sum(col in row_values for col in BASE_COLUMNS)

        if score > best_score:
            best_score = score
            best_row = i

    if best_score < 5:
        return None

    return best_row


def get_product_prefix(value):
    value = str(value).strip()

    if value.upper().startswith("TS"):
        return "TS"

    return value[:3]


def clean_peak_excel(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)

    header_row = find_header_row(raw_df)

    if header_row is None:
        st.error("หาแถว Header ไม่เจอ กรุณาเช็คไฟล์ Excel อีกครั้ง")
        st.stop()

    df = raw_df.iloc[header_row + 1:].copy()
    df.columns = raw_df.iloc[header_row].astype(str).str.strip()
    df = df.reset_index(drop=True)

    df = df.dropna(how="all")

    if len(df) > 0:
        df = df.iloc[:-1]

    df.columns = df.columns.astype(str).str.strip()

    missing_cols = [col for col in BASE_COLUMNS if col not in df.columns]
    if missing_cols:
        st.error(f"ไม่พบ column เหล่านี้ในไฟล์: {missing_cols}")
        st.write("Columns ที่เจอจริง:", list(df.columns))
        st.stop()

    df = df[BASE_COLUMNS]

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(["nan", "None", ""], pd.NA)

    for col in ["จำนวน", "ราคา", "มูลค่าก่อนภาษี"]:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("฿", "", regex=False)
            .str.replace(" ", "", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["วันที่ออก"] = pd.to_datetime(df["วันที่ออก"], errors="coerce")

    prefix = df["รหัสสินค้า/บริการ"].apply(get_product_prefix)
    mapped = prefix.map(code_map)

    df["Category"] = mapped.apply(lambda x: x[0] if isinstance(x, tuple) else "Unknown")
    df["Sub-Category"] = mapped.apply(lambda x: x[1] if isinstance(x, tuple) else "Unknown")

    df = df[FINAL_COLUMNS]

    return df, header_row


def to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="RawData")

    return output.getvalue()


uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

output_filename = st.text_input(
    "Output file name",
    value="รายงานยอดขายในระบบ Peak Format.xlsx"
)

if uploaded_file is None:
    st.info("Please upload Excel file.")
    st.stop()

cleaned_df, header_row = clean_peak_excel(uploaded_file)

st.success(f"Clean completed — detected header row: {header_row}")

c1, c2, c3 = st.columns(3)
c1.metric("Rows", f"{len(cleaned_df):,}")
c2.metric("Columns", f"{len(cleaned_df.columns):,}")
c3.metric("Unknown Category", f"{(cleaned_df['Category'] == 'Unknown').sum():,}")

st.subheader("Cleaned Data Preview")
st.dataframe(cleaned_df.head(100), use_container_width=True)

excel_file = to_excel(cleaned_df)

st.download_button(
    label="Download Cleaned Excel",
    data=excel_file,
    file_name=output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)