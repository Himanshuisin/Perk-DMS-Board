import os
import sys
import shutil
import glob
import time
import subprocess
import re
import tempfile
from datetime import datetime, timedelta

# Default target date (Yesterday)
TARGET_DATE = datetime.now().date() - timedelta(days=1)

# File Paths
FILE_PATH = r"C:\Users\AUB5367\OneDrive - MDLZ\DMS-2\PERK MIS.-2026.xlsx"
SHEET_DMS = "DMS-2"
MINOR_STOPS_DIR = r"C:\Users\AUB5367\Desktop\New folder\AM STEP-4\dESKTOP BACKUP\IL6S\Minor Stop-SWP\MTBF-PBI 2026\Minor Stops"
SHOPLOGIX_DIR = r"c:\Users\AUB5367\OneDrive - MDLZ\Desktop\New folder\AM STEP-4\dESKTOP BACKUP\IL6S\Minor Stop-SWP\MTBF-PBI 2026\DMS-Downtimes"
EFFICIENCY_DIR = r"C:\Users\AUB5367\OneDrive - MDLZ\Desktop\New folder\AM STEP-4\dESKTOP BACKUP\IL6S\Minor Stop-SWP\MTBF-PBI 2026\Machine Efficiency"
DOWNLOAD_DIR = r"C:\Users\AUB5367\Downloads"

SHOPLOGIX_MAKING_URL = "https://portal.shoplogix.com/app/main/dashboards/6a86963d09213c5b9401f708"
SHOPLOGIX_EFF_URL = "https://portal.shoplogix.com/app/main/dashboards/66af6f6e1f22b100361259da"

def safe_copy_file(src):
    temp_dir = tempfile.gettempdir()
    dst = os.path.join(temp_dir, f"temp_{os.path.basename(src)}")
    try:
        subprocess.run(
            f'copy /Y "{src}" "{dst}"',
            shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    except Exception:
        shutil.copy2(src, dst)
    return dst

def to_norm_tag(val):
    import pandas as pd
    if pd.isna(val) or str(val).strip() in ("", "nan", "NAN"):
        return ""
    val_s = str(val).strip()
    dt = pd.to_datetime(val_s, errors="coerce")
    if pd.notna(dt):
        return f"{dt.day}-{dt.strftime('%b')}".upper()
    m = re.match(r"^(\d{1,2})[-/ ]*([A-Za-z]{3})", val_s)
    if m:
        return f"{int(m.group(1))}-{m.group(2)[:3]}".upper()
    return val_s.upper()

def parse_date_to_timestamp(val_s):
    import pandas as pd
    if pd.isna(val_s) or not str(val_s).strip():
        return pd.Timestamp.min
    s = str(val_s).strip()
    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return dt
    try:
        return pd.to_datetime(s + "-2026", format="%d-%b-%Y")
    except Exception:
        return pd.Timestamp.min


# ==============================================================================
# SECTION A: STREAMLIT DASHBOARD (Runs by default everywhere)
# ==============================================================================
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(
    page_title="PERK DMS BOARD",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ... (keep all your dashboard layout, tables, and charts code here without any indentation wrapper) ...
