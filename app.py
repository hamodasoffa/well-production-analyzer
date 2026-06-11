# Save as app.py
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import StringIO

st.set_page_config(page_title="Well Production Analyzer", layout="wide")
st.title("🛢️ Well Production Test Analyzer")
st.write("Upload your Excel test sheet and analyze production data.")
