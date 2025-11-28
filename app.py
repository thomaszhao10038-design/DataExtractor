import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from io import BytesIO

# Excel base date for serial numbers
EXCEL_BASE = datetime(1899, 12, 30)

def excel_serial(dt):
    return (dt - EXCEL_BASE).days + (dt.hour*3600 + dt.minute*60 + dt.second)/86400

def time_fraction(t):
    h, m, s = map(int, t.split(':'))
    return h/24 + m/1440 + s/86400

def process_raw_csv(file):
    df = pd.read_csv(file, skiprows=2, low_memory=False)
    df['Datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'], dayfirst=True)
    df['Day_Serial'] = df['Datetime'].dt.date.apply(lambda x: excel_serial(datetime.combine(x, datetime.min.time())))
    df['Local_Time'] = df['Time'].apply(time_fraction)
    df['kW'] = -df['PSum'].astype(float) / 1000
    df['Power_W'] = df['PSum'].astype(float)
    return df

# -------------------------------------------------
st.set_page_config(page_title="EnergyAnalyser", layout="wide")
st.title("EnergyAnalyser – Raw MSB → Final Excel")
st.markdown("Upload the three raw CSV files (MSB 1, MSB 2, MSB 3) → get the exact **final document.xlsx** format")

msb1_file = st.file_uploader("Raw data MSB 1.csv", type="csv")
msb2_file = st.file_uploader("Raw data MSB 2.csv", type="csv")
msb3_file = st.file_uploader("Raw data MSB 3.csv", type="csv")

if msb1_file and msb2_file and msb3_file:
    with st.spinner("Processing files..."):
        df1 = process_raw_csv(msb1_file)
        df2 = process_raw_csv(msb2_file)
        df3 = process_raw_csv(msb3_file)

        # Daily summaries for Total & Load Apportioning sheets
        daily1 = df1.groupby('Day_Serial')['kW'].mean().reset_index()
        daily2 = df2.groupby('Day_Serial')['kW'].mean().reset_index()
        daily3 = df3.groupby('Day_Serial')['kW'].mean().reset_index()
        max_daily = (daily1['kW'] + daily2['kW'] + daily3['kW']).max()

        # Create Excel in memory
        output = BytesIO()
        wb = openpyxl.Workbook()

        # Helper to write one MSB sheet
        def write_msb_sheet(name, df):
            ws = wb.create_sheet(title=name)
            days = sorted(df['Day_Serial'].unique())
            col = 2
            for i, day in enumerate(days):
                day_data = df[df['Day_Serial'] == day].sort_values('Local_Time')
                # Header row
                ws.cell(1, col,   "Local Time Stamp")
                ws.cell(1, col+1, "Active Power (W)")
                ws.cell(1, col+2, "kW")
                ws.cell(1, col+3, "")
                # Day serial in first data row of this block
                ws.cell(2, col, day if i > 0 else "")
                # Data
                for r, row in enumerate(day_data.itertuples(), start=2):
                    ws.cell(r, col,   row.Local_Time)
                    ws.cell(r, col+1, row.Power_W)
                    ws.cell(r, col+2, row.kW)
                    ws.cell(r, col+3, "")
                col += 4
            # UTC Offset column
            if days:
                ws['A2'] = days[0]

        write_msb_sheet("MSB 1", df1)
        write_msb_sheet("MSB 2", df2)
        write_msb_sheet("MSB 3", df3)

        # Total MSB sheet
        ws_total = wb.create_sheet("Total MSB")
        ws_total['A1'] = "Date"
        ws_total['B1'] = "MSB-1"
        ws_total['C1'] = "MSB-2"
        ws_total['D1'] = "MSB-3"
        ws_total['E1'] = "Total Building Load"
        all_days = sorted(set(daily1['Day_Serial']) | set(daily2['Day_Serial']) | set(daily3['Day_Serial']))
        for i, day in enumerate(all_days, 2):
            k1 = daily1.loc[daily1['Day_Serial']==day, 'kW'].iloc[0] if day in daily1['Day_Serial'].values else 0
            k2 = daily2.loc[daily2['Day_Serial']==day, 'kW'].iloc[0] if day in daily2['Day_Serial'].values else 0
            k3 = daily3.loc[daily3['Day_Serial']==day, 'kW'].iloc[0] if day in daily3['Day_Serial'].values else 0
            ws_total.cell(i,1, day)
            ws_total.cell(i,2, round(k1, 5))
            ws_total.cell(i,3, round(k2, 5))
            ws_total.cell(i,4, round(k3, 5))
            ws_total.cell(i,5, round(k1+k2+k3, 5))

        # Load Apportioning sheet (matches your sample layout & formulas)
        ws_load = wb.create_sheet("Load Apportioning (Marelli)")
        headers = ["Date","Day","MSB No. 2","MSB No. 3","MSB No. 4","MSB No. 5","MSB No. 6","MSB No. 7","EMSB","Maximum Demand, kW"]
        for c, h in enumerate(headers, 1):
            ws_load.cell(1, c, h)

        row = 3
        for day in all_days:
            dt = datetime(1899,12,30) + timedelta(days=int(day))
            k1 = daily1.loc[daily1['Day_Serial']==day, 'kW'].iloc[0] if day in daily1['Day_Serial'].values else 0
            k2 = daily2.loc[daily2['Day_Serial']==day, 'kW'].iloc[0] if day in daily2['Day_Serial'].values else 0
            k3 = daily3.loc[daily3['Day_Serial']==day, 'kW'].iloc[0] if day in daily3['Day_Serial'].values else 0
            ws_load.cell(row,1, day)
            ws_load.cell(row,2, dt.strftime("%A"))
            ws_load.cell(row,3, round(k1, 5))   # MSB No. 2 = MSB1 average
            ws_load.cell(row,4, round(k2, 5))   # MSB No. 3 = MSB2
            ws_load.cell(row,5, round(k3, 5))   # MSB No. 4 = MSB3
            ws_load.cell(row,10, round(max_daily, 5))
            row += 1

        # Averages & Totals (Excel formulas)
        last_data_row = row - 1
        ws_load.cell(row, 2, "Average")
        for c in range(3, 10):
            ws_load.cell(row, c, f"=AVERAGE({openpyxl.utils.get_column_letter(c)}3:{openpyxl.utils.get_column_letter(c)}{last_data_row})")
        row += 1
        ws_load.cell(row, 2, "Total")
        for c in range(3, 10):
            ws_load.cell(row, c, f"=SUM({openpyxl.utils.get_column_letter(c)}3:{openpyxl.utils.get_column_letter(c)}{last_data_row})")

        # Remove default sheet
        wb.remove(wb["Sheet"])

        wb.save(output)
        output.seek(0)

    st.success("Processing complete!")
    st.download_button(
        label="Download final_document.xlsx",
        data=output,
        file_name="final_document.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
