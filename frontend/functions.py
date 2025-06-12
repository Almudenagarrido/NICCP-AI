import streamlit as st
import pandas as pd
import numpy as np
import requests
from io import BytesIO


class GeneralInformation:

    def __init__(self, subsection):
        self.subsection = subsection
        self.sheet_map = {
            "Financial Electricity Inputs": "Electricity",
            "Financial LPG Inputs": "LPG",
            "CO2 - Carbon Finance Inputs": "Carbon"
        }
        self.get_url = "http://localhost:8000/general-information"
        self.post_url = "http://localhost:8000/save-general-information"
        self.sheet_name = self.sheet_map.get(subsection)
        self.df = None
        self.edited_df = None

    def fetch_and_load(self):
        res = requests.get(self.get_url)
        if res.status_code != 200:
            st.error("Could not load 'general-information.xlsm'")
            return False

        if not self.sheet_name:
            st.warning("No sheet found in file for selected subsection.")
            return False

        file_data = BytesIO(res.content)
        self.df = pd.read_excel(file_data, sheet_name=self.sheet_name, engine="openpyxl", index_col=0)
        return True

    def show_excel_editor(self):
        st.subheader(self.subsection)
        self.edited_df = st.data_editor(self.df, num_rows="dynamic")

    def save_changes(self):
        self.edited_df = self.edited_df.replace([np.nan], "-")
        data_json = self.edited_df.reset_index().to_dict(orient="records")
        payload = {
            "sheet_name": self.sheet_name,
            "data": data_json
        }
        res = requests.post(self.post_url, json=payload)
        if res.status_code == 200:
            st.success("Changes saved correctly in 'general-information.xlsm'")
        else:
            st.error("Error saving the changes, try again later.")

    def __call__(self):
        if not self.fetch_and_load():
            return
        self.show_excel_editor()
        if st.button("Save changes"):
            self.save_changes()


def manage_technoeconomic_models():

    res = requests.get("http://localhost:8000/technoeconomic-models")
    if res.status_code == 200:
        files = res.json().get("files", [])
        model = st.selectbox("Manage techno-economic models", files)
        if model:
            file_res = requests.get(f"http://localhost:8000/technoeconomic-models/{model}")
            if file_res.status_code == 200:
                data = file_res.content
                df = pd.read_excel(BytesIO(data))
                st.write(df)
            else:
                st.error("Failed to download the file")
    else:
        st.error("Failed to fetch the list of case studies")