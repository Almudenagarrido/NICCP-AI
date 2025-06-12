import streamlit as st
import requests
import pandas as pd
from io import BytesIO
import layout as l
import functions as f


l.header()
l.start_content()
page, subsection = l.sidebar()

API_URL = "http://127.0.0.1:8000"

if page == "General Information" and subsection:
    f.GeneralInformation(subsection)()
    
if page == "Manage Techno-Economic Models":
    f.manage_technoeconomic_models()

l.end_content()
l.footer()
