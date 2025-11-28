import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl.utils

# Excel base date (Windows Excel bug: 1900 is leap year)
EXCEL_BASE = datetime(1899, 12, 30)

def excel_date_serial(dt):
    """Convert datetime → Excel serial number (including time fraction)"""
    delta = dt - EXCEL_BASE
    return delta.days + (delta.seconds / 86400)

def time_to_fraction(t):
    """Convert 'HH:MM:SS' string → fraction of day (safe)"""
    if pd.isna(t) or not isinstance(t, str):
        return 0.0
    parts = t.strip().split(':')
    if len(parts) != 3:
        return 0.0
    try:
        h, m, s = map(int, parts)
        return h/24 + m/1440 + s/86400
    except:
        return 0.0

def process_raw_csv(file):
    """Robust
