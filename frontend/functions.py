import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import pandas as pd
import numpy as np
import requests
from io import BytesIO
import datetime


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
        self.non_editable_columns = ["Inputs", "Units"]
        self.percentage_rows = [
                "Expected non-technical losses",
                "Tax rate",
                "% EBITDA Margin",
                "OPEX subsidies",
                "CO2 certificate",
                "Liquidity"
        ]

    def fetch_and_load(self):
        res = requests.get(self.get_url)
        if res.status_code != 200:
            st.error("Could not load 'general-information.xlsx'")
            return False

        if not self.sheet_name:
            st.warning("No sheet found in file for selected subsection.")
            return False

        file_data = BytesIO(res.content)
        self.df = pd.read_excel(file_data, sheet_name=self.sheet_name, engine="openpyxl")

        for col in self.df.columns:
            if col not in ["Inputs", "Units"]:
                self.df[col] = pd.to_numeric(self.df[col].replace("-", pd.NA), errors="coerce")
        
        return True

    def show_excel_editor(self):
        st.subheader(self.subsection)
        df = self.df.copy()
        gb = GridOptionsBuilder.from_dataframe(df)
        
        for col in df.columns:
            if col in self.non_editable_columns:
                gb.configure_column(col, editable=False)
            else:
                gb.configure_column(col, editable=True)

        grid_options = gb.build()
        grid_response = AgGrid(
            df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            allow_unsafe_jscode=True,
            enable_enterprise_modules=False
        )
        self.edited_df = pd.DataFrame(grid_response["data"])

    def cell_validations(self, df):
        invalid_cells = []

        for index, row in df.iterrows():
            row_label = str(row.get("Inputs", "")).lower()
            is_percentage = any(name.lower() in row_label for name in self.percentage_rows)

            for col in df.columns:
                if col in self.non_editable_columns:
                    continue

                value = row[col]
                try:
                    if value == "-" or pd.isna(value):
                        continue

                    num_value = float(value)

                    if is_percentage:
                        if not (0 <= num_value <= 100):
                            invalid_cells.append((index, row["Inputs"], col, value))
                    else:
                        if num_value < 0:
                            invalid_cells.append((index, row["Inputs"], col, value))

                except Exception:
                    invalid_cells.append((index, row["Inputs"], col, value))

        return invalid_cells
               
    def save_changes(self):
        df = self.edited_df.copy()
        df = df[self.df.columns]
        
        invalid_cells = self.cell_validations(df)
        if invalid_cells:
            for _, input_name, col, value in invalid_cells:
                st.warning(f"Invalid value '{value}' in row '{input_name}' (column '{col}')")
            st.error("Fix validation errors before saving.")
            return

        data_json = df.reset_index(drop=True).to_dict(orient="records")
        payload = {
            "sheet_name": self.sheet_name,
            "data": data_json
        }
        res = requests.post(self.post_url, json=payload)
        if res.status_code == 200:
            st.success("Changes saved correctly in 'general-information.xlsx'")
        else:
            st.error("Error saving the changes, try again later.")

    def reset_sheet(self, sheet_name: str):
        response = requests.post("http://localhost:8000/reset-general-information", json={"sheet_name": sheet_name, "data": []})
        if response.status_code == 200:
            st.success(response.json().get("message", "Sheet reset"))
        else:
            st.error(response.json().get("error", "Something went wrong"))

    def __call__(self):
        
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()
        
        if st.button("Save"):
            self.save_changes()

        if st.button("Reset"):
            self.reset_sheet(self.sheet_name)
            self.show_excel_editor()


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