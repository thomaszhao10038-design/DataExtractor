import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl.utils

# Excel base date (Windows Excel treats 1900 as leap year → use 1899-12-30)
EXCEL_BASE = datetime(1899, 12, 30)

def excel_date_serial(dt):
    delta = dt - EXCEL_BASE
    return delta.days + (delta.seconds / 86400.0)

def time_to_fraction(t):
    if pd.isna(t) or not isinstance(t, str):
        return 0.0
    parts = str(t).strip().split(':')
    if len(parts) < 3:
        return 0.0
    try:
        h, m, s = map(int, parts[:3])
        return h/24 + m/1440 + s/86400
    except:
        return 0.0

def process_raw_csv(file):
    df = pd.read_csv(file, skiprows=2, low_memory=False)
    df['Datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'],
                                  format='%d/%m/%Y %H:%M:%S',
                                  errors='coerce',
                                  dayfirst=True)
    df = df.dropna(subset=['Datetime']).copy()
    df['Day_Serial'] = df['Datetime'].apply(lambda x: int(excel_date_serial(x)))
    df['Local_Time_Stamp'] = df['Time'].apply(time_to_fraction)
    df['Active_Power_W'] = pd.to_numeric(df['PSum'], errors='coerce').fillna(0)
    df['kW'] = -df['Active_Power_W'] / 1000
    return df

# =============================================
st.set_page_config(page_title="EnergyAnalyser", layout="centered")
st.title("EnergyAnalyser")
st.markdown("Upload the three raw CSV files (MSB 1, MSB 2, MSB 3) → get **final_document.xlsx**")

msb1 = st.file_uploader("Raw data MSB 1.csv", type="csv")
msb2 = st.file_uploader("Raw data MSB 2.csv", type="csv")
msb3 = st.file_uploader("Raw data MSB 3.csv", type="csv")

if st.button("Generate Excel") and msb1 and msb2 and msb3:
    with st.spinner("Processing… (10–30 seconds)"):
        df1 = process_raw_csv(msb1)
        df2 = process_raw_csv(msb2)
        df3 = process_raw_csv(msb3)

        daily1 = df1.groupby('Day_Serial')['kW'].mean().round(6)
        daily2 = df2.groupby('Day_Serial')['kW'].mean().round(6)
        daily3 = df3.groupby('Day_Serial')['kW'].mean().round(6)
        all_days = sorted({*daily1.index, *daily2.index, *daily3.index})

        output = BytesIO()
        wb = openpyxl.Workbook()

        def write_msb_sheet(name, df):
            ws = wb.create_sheet(title=name)
            days = sorted(df['Day_Serial'].unique())
            col = 2
            for i, day in enumerate(days):
                block = df[df['Day_Serial'] == day].sort_values('Local_Time_Stamp')
                ws.cell(1, col,   "Local Time Stamp")
                ws.cell(1, col+1, "Active Power (W)")
                ws.cell(1, col+2, "kW")
                ws.cell(1, col+3, "")
                if i > 0:
                    ws.cell(2, col, day)
                for r, row in enumerate(block.itertuples(), start=2):
                    ws.cell(r, col,   round(row.Local_Time_Stamp, 11))
                    ws.cell(r, col+1, row.Active_Power_W)
                    ws.cell(r, col+2, row.kW)
                    ws.cell(r, col+3, "")
                col += 4
            if days:
                ws['A1'] = "UTC Offset (minutes)"
                ws['A2'] = days[0]

        write_msb_sheet("MSB 1", df1)
        write_msb_sheet("MSB 2", df2)
        write_msb_sheet("MSB 3", df3)

        # Total MSB sheet
        ws_total = wb.create_sheet("Total MSB")
        for c, h in enumerate(["Date","MSB-1","MSB-2","MSB-3","Total Building Load"], 1):
            ws_total.cell(1, c, h)
        for i, day in enumerate(all_days, 2):
            k1 = daily1.get(day, 0)
            k2 = daily2.get(day, 0)
            k3 = daily3.get(day, 0)
            ws_total.cell(i,1, day)
            ws_total.cell(i,2, k1)
            ws_total.cell(i,3, k2)
            ws_total.cell(i,4, k3)
            ws_total.cell(i,5, k1+k2+k3)

        # Load Apportioning (Marelli)
        ws_load = wb.create_sheet("Load Apportioning (Marelli)")
        headers = ["Date","Day","MSB No. 2","MSB No. 3","MSB No. 4","MSB No. 5",
                   "MSB No. 6","MSB No. 7","EMSB","Maximum Demand, kW"]
        for c, h in enumerate(headers, 1):
            ws_load.cell(1, c, h)

        row = 3
        for day in all_days:
            dt = EXCEL_BASE + timedelta(days=day)
            k1 = daily1.get(day, 0)
            k2 = daily2.get(day, 0)
            k3 = daily3.get(day, 0)
            ws_load.cell(row,1, day)
            ws_load.cell(row,2, dt.strftime("%A"))
            ws_load.cell(row,3, k1)
            ws_load.cell(row,4, k2)
            ws_load.cell(row,5, k3)
            ws_load.cell(row,10, round(k1+k2+k3, 5))
            row += 1

        last = row - 1
        ws_load.cell(row,2,"Average")
        for c in range(3,11):
            letter = openpyxl.utils.get_column_letter(c)
            ws_load.cell(row,c,f"=AVERAGE({letter}3:{letter}{last})")
        row += 1
        ws_load.cell(row,2,"Total")
        for c in range(3,11):
            letter = openpyxl.utils.get_column_letter(c)
            ws_load.cell(row,c,f"=SUM({letter}3:{letter}{last})")

        wb.remove(wb["Sheet"])
        wb.save(output)
        output.seek(0)

    st.success("Done!")
    st.download_button(
        label="Download final_document.xlsx",
        data=output.getvalue(),
        file_name="final_document.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload all three raw CSV files → click **Generate Excel**")
