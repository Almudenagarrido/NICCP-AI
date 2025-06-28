import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from st_aggrid.shared import JsCode
import pandas as pd
import numpy as np
import requests
import uuid
from io import BytesIO
import time


class FuelMarketInformation:

    def __init__(self, api_url, fuel_market):
        self.api_base = api_url
        self.get_url = f"{self.api_base}/fuel-market-information"
        self.save_url = f"{self.api_base}/save-fuel-market-information"
        self.reset_url = f"{self.api_base}/reset-fuel-market-information"
        self.add_url = f"{self.api_base}/add-fuel-market"
        self.delete_url = f"{self.api_base}/delete-fuel-market"
        self.fuel_market = fuel_market
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
        self.df_heights = {"E-Cooking": 170, "LPG": 350, "Carbon": 170}

    def get_fuel_markets(self):
        try:
            res = requests.get(self.get_url)
            if res.status_code != 200:
                st.error("The file 'fuel-market-information.xlsx' was not found.")
                return []
            file_data = BytesIO(res.content)
            xls = pd.ExcelFile(file_data)
            return xls.sheet_names
        except Exception as e:
            st.error(f"Error fetching financial markets: {e}")
            return []

    def fetch_and_load(self):
        res = requests.get(self.get_url)
        if res.status_code != 200:
            st.error("Could not load 'fuel-market-information.xlsx'")
            return False

        if not self.fuel_market:
            st.write(self.fuel_market)
            st.warning("No sheet found in file for selected technology market.")
            return False

        file_data = BytesIO(res.content)
        self.df = pd.read_excel(file_data, sheet_name=self.fuel_market, engine="openpyxl")

        for col in self.df.columns:
            if col not in ["Inputs", "Units"]:
                self.df[col] = pd.to_numeric(self.df[col].replace("-", pd.NA), errors="coerce")
        
        return True

    def show_excel_editor(self):

        st.subheader(f"{self.fuel_market} Financial Inputs")
        df = self.df.copy()
        gb = GridOptionsBuilder.from_dataframe(df)
        height = self.df_heights[self.fuel_market] if self.fuel_market in self.df_heights else self.df_heights["E-Cooking"]
        
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
            enable_enterprise_modules=False,
            height=height
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

        for i, row in df.iterrows():
            if str(row.get("Inputs", "")).strip().lower() == "number of years that you could sell those carbon credits":
                for col in df.columns:
                    if col not in ["Inputs", "Units", "Baseline"]:
                        df.at[i, col] = pd.NA
        
        for col in df.columns:
            if col not in self.non_editable_columns:
                df[col] = df[col].fillna("-")
        
        invalid_cells = self.cell_validations(df)
        if invalid_cells:
            for _, input_name, col, value in invalid_cells:
                st.warning(f"Invalid value '{value}' in row '{input_name}' (column '{col}')")
            st.error("Fix validation errors before saving.")
            return

        data_json = df.reset_index(drop=True).to_dict(orient="records")
        payload = {
            "model": "",
            "sheet_name": self.fuel_market,
            "data": data_json
        }
        res = requests.post(self.save_url, json=payload)
        if res.status_code == 200:
            st.success(f"Changes in '{self.fuel_market}' saved successfully.")
        else:
            st.error("Error saving the changes, try again later.")

    def reset_sheet(self):
        response = requests.post(self.reset_url, json={"model": "", "sheet_name": self.fuel_market, "data": []})
        if response.status_code == 200:
            st.success(f"File '{self.fuel_market}' reset to template version.")
        else:
            st.error(response.json().get("error", "Something went wrong"))

    def add_market(self):
        st.subheader("Add new fuel market")
        new_market = st.text_input("Enter name of the new technology: ")

        if st.button("Create market"):
            new_market = new_market.strip()

            if new_market == "":
                st.warning("Fuel market name cannot be empty.")
            else:
                try:
                    response = requests.post(
                        self.add_url, 
                        json={"name": new_market}
                    )
                    if response.status_code == 200:
                        st.success(f"Market '{new_market}' added successfully.")
                        st.session_state.fuel_market = f"{new_market}"
                        st.rerun()
                    else:
                        try:
                            error_msg = response.json().get("error", "Unknown error occurred.")
                            st.error(error_msg)
                        except:
                            st.error("Error contacting backend.")
                except Exception as e:
                    st.error(f"Failed to connect to backend: {e}")

    def delete_market(self, sheet_name):

        try:
            response = requests.post(self.delete_url, json={"name": sheet_name})
            if response.status_code == 200:
                return True
            else:
                st.error(f"Failed to delete '{sheet_name}'. Server responded with {response.status_code}.")
                return False
        except Exception as e:
            st.error(f"Error deleting market: {e}")
            return False

    def __call__(self):

        if self.fuel_market == "Add":
            self.add_market()
            return
        
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()
        
        if st.button("Save"):
            self.save_changes()
            time.sleep(1)
            st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            time.sleep(1)
            st.rerun()


class TechnoEconomicModels:
    
    def __init__(self, api_url, subsection, model, fuel_market):
        self.api_url = api_url
        self.subsection = subsection
        self.subsections = {
            "Manage": ManageModels(self.api_url),
            "Design Capital Structure": DesignCapitalStructure(api_url, subsection, model, fuel_market),
            "Techno-Economic Inputs": TechnoEconomicInputs(api_url, subsection, model, fuel_market),
            "Summary Financing": SummaryFinancing(self.api_url)
        }
        self.model = model
        self.fuel_market = fuel_market

    def __call__(self):
        if self.subsection in self.subsections:
            self.subsections[self.subsection].show()
        else:
            st.info("Please select a valid subsection.")


class ManageModels:

    def __init__(self, api_base):
        self.api_base = api_base
        self.list_url = f"{self.api_base}/technoeconomic-models"
        self.download_files_url = f"{self.api_base}/download-technoeconomic-model-files"
        self.create_url = f"{self.api_base}/create-technoeconomic-model"
        self.delete_url = f"{self.api_base}/delete-technoeconomic-model"
        self.upload_url = f"{self.api_base}/upload-technoeconomic-model"
        self.valid_extensions = ["xlsx", "xlsm", "xls", "xltx", "xltm"]

    def list_technoeconomic_models(self):
        try:
            res = requests.get(self.list_url)
            return res.json().get("models", []) if res.status_code == 200 else []
        except:
            st.error("Error loading models from server.")
            return []

    def remove_extension(self, filename):
        for ext in self.valid_extensions:
            if filename.lower().endswith(f".{ext}"):
                return filename[: -len(ext) - 1]
        return filename
    
    def download_technoeconomic_model_files(self, name):
        url = f"{self.download_files_url}/{name}"
        res = requests.get(url)
        return res.content if res.status_code == 200 else None

    def delete_technoeconomic_model(self, name):
        res = requests.delete(f"{self.delete_url}/{name}")
        if res.status_code == 200:
            return True, f"Model '{name}' deleted successfully."
        else:
            return False, "Failed to delete the model."

    def upload_technoeconomic_model(self, name, file):
        files = {"file": (file.name, file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        res = requests.post(f"{self.upload_url}/{name}", files=files)

        if res.status_code == 200:
            return True
        else:
            error_detail = res.json().get("detail", res.text)
            st.error(f"Upload failed: {error_detail}")
            return False

    def create_technoeconomic_model(self, model, start_year, end_year):
        try:
            url = f"{self.create_url}/{model}"
            response = requests.post(
                url, params={"start_year": start_year, "end_year": end_year}
            )
            data = response.json()
            if "success" in data and data["success"]:
                return True, data["message"]
            else:
                return False, data.get("error", "Unknown error occurred.")
        except Exception as e:
            return False, str(e)

    def model_creator(self):
        st.markdown("#### ➕ Create New Techno-Economic Model")
        with st.form("create_model_form"):
            name = st.text_input("Model Name")
            col1, col2 = st.columns(2)
            with col1:
                start_year = st.number_input("Start Year", step=1, format="%d")
            with col2:
                end_year = st.number_input("End Year", step=1, format="%d")
            create = st.form_submit_button("Create")

            if create:
                if not name.strip():
                    st.error("Model name cannot be empty.")
                elif start_year >= end_year:
                    st.error("Start year must be less than end year.")
                else:
                    success, msg = self.create_technoeconomic_model(name.strip(), int(start_year), int(end_year))
                    if success:
                        st.success(msg)
                        st.session_state["model_name_input"] = ""
                        st.session_state["start_year_input"] = 0
                        st.session_state["end_year_input"] = 0
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)

    def show_technoeconomic_models(self, models):
        for model in models:
            clean_name = self.remove_extension(model)
            col1, col2, col3, col4 = st.columns([0.7, 0.1, 0.1, 0.1])

            with col1:
                if st.button(f"📄 {clean_name}"):
                    st.session_state.page = "Techno-Economic Models"
                    st.session_state.subsection = "Design Capital Structure"
                    st.session_state.model = clean_name
                    st.rerun()

            with col2:
                content = self.download_technoeconomic_model_files(clean_name)
                if content:
                    st.download_button(
                        "⬇️",
                        data=content,
                        file_name=f"{clean_name}_files.zip",
                        mime="application/zip"
                    )

            with col3:
                if st.button("📤", key=f"trigger_upload_{model}"):
                    st.session_state[f"show_uploader_{model}"] = True

            with col4:
                if st.button("❌", key=f"delete_{clean_name}"):
                    success, message = self.delete_technoeconomic_model(clean_name)
                    if success:
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                        st.rerun()

            if st.session_state.get(f"show_uploader_{model}", False):
                file = st.file_uploader(
                    f"Upload file for {model}",
                    type=self.valid_extensions,
                    key=f"upload_{model}",
                    label_visibility="collapsed"
                )
                if file:
                    if any(file.name.lower().endswith(f".{ext}") for ext in self.valid_extensions):
                        upload_success = self.upload_technoeconomic_model(clean_name, file)
                        if upload_success:
                            st.success(f"Uploaded to model '{model}'")
                            st.session_state[f"show_uploader_{model}"] = False
                            time.sleep(1)
                            st.rerun()
                    else:
                        st.error(f"Only Excel files are allowed: {', '.join(self.valid_extensions)}")

    def show(self):
        st.subheader("Manage Techno-Economic Models")
        models = self.list_technoeconomic_models()
        if not models:
            st.info("No techno-economic models available.")
        else:
            self.show_technoeconomic_models(models)
        self.model_creator()


class DesignCapitalStructure:
    
    def __init__(self, api_base, subsection, model, fuel_market):
        self.api_base = api_base
        self.get_url = f"{self.api_base}/design-capital-structure"
        self.save_url = f"{self.api_base}/save-design-capital-structure"
        self.reset_url = f"{self.api_base}/reset-design-capital-structure"
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.df = None
        self.subtables = {"Finance": None, "Division": None, "Debts": None}
        self.subtable_titles = {
            "Finance": "Financiation",
            "Division": "Total",
            "Debts": "Type"
        }
        self.subtable_heights = {
            "Finance": 100,
            "Division": 230,
            "Debts": 150
        }
        self.editable_columns = {
            "Finance": ["Amount"],
            "Division": ["Amount Total", "Amount Division"],
            "Debts": ["Baseline"] + [str(year) for year in range(2020, 2051)]
        }

    def get_design_capital(self):
        try:
            url = f"{self.get_url}/{self.model}"
            res = requests.get(url)
            if res.status_code != 200:
                st.error("The file 'design-capital-structure.xlsx' was not found.")
                return None

            file_data = BytesIO(res.content)
            xls = pd.ExcelFile(file_data)
            return xls.sheet_names

        except Exception as e:
            st.error(f"Error fetching capital structure: {e}")
            return None

    def fetch_and_load(self):
        url = f"{self.get_url}/{self.model}"
        res = requests.get(url)
        if res.status_code != 200:
            st.error(f"Could not load 'design-capital-structure-{self.model}.xlsx'")
            return False

        if not self.subsection:
            st.warning("No sheet found in file for selected subsection.")
            return False
        
        file_data = BytesIO(res.content)
        self.df = pd.read_excel(file_data, sheet_name=self.
        fuel_market, engine="openpyxl", header=None, index_col=None)
        self.split_into_subtables()
        return True
    
    def split_into_subtables(self):
        df = self.df
        section_starts = {}
        
        for idx, row in df.iterrows():
            row_str = " ".join(str(v) for v in row if pd.notna(v))
            for key, keyword in self.subtable_titles.items():
                if key not in section_starts and keyword in row_str:
                    section_starts[key] = idx

        sorted_sections = sorted(section_starts.items(), key=lambda x: x[1])
        for i, (key, start_idx) in enumerate(sorted_sections):
            end_idx = sorted_sections[i + 1][1] if i + 1 < len(sorted_sections) else len(df)
            subdf = df.iloc[start_idx:end_idx].reset_index(drop=True)
            
            subdf.replace("-", np.nan, inplace=True)
            subdf.dropna(axis=0, how="all", inplace=True)
            subdf.dropna(axis=1, how="all", inplace=True)
            
            if not subdf.empty:
                header_row = subdf.iloc[0]
                subdf.columns = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in header_row]
                subdf = subdf[1:].reset_index(drop=True)
            
            self.subtables[key] = subdf

    def show_excel_editor(self):
        st.subheader(f"{self.fuel_market} Financial Plan")

        for key, df in self.subtables.items():
            if df is not None and not df.empty:
                df.columns = df.columns.astype(str)
                base_columns = {col: col.split('_')[0] for col in df.columns}
                editable_cols = self.editable_columns.get(key, [])
                
                gb = GridOptionsBuilder.from_dataframe(df)
                for col in df.columns:
                    is_editable = base_columns[col] in editable_cols
                
                    if pd.api.types.is_numeric_dtype(df[col]):
                        gb.configure_column(
                            col, 
                            editable=is_editable,
                            type=["numericColumn", "numberColumnFilter"]
                        )
                    else:
                        gb.configure_column(
                            col,
                            editable=is_editable
                        )
                
                grid_options = gb.build()
                grid_response = AgGrid(
                    df,
                    gridOptions=grid_options,
                    update_mode=GridUpdateMode.VALUE_CHANGED,
                    allow_unsafe_jscode=True,
                    enable_enterprise_modules=False,
                    height=self.subtable_heights[key],
                    reload_data=False
                )
                
                self.subtables[key] = pd.DataFrame(grid_response["data"])

    def save_changes(self):
        full_df = self.df.copy()

        for key, title in self.subtable_titles.items():
            subdf = self.subtables.get(key)
            if subdf is None or subdf.empty:
                continue

            section_start = None
            for idx, row in full_df.iterrows():
                if title in " ".join(str(v) for v in row if pd.notna(v)):
                    section_start = idx
                    break
            if section_start is None:
                continue

            headers = full_df.iloc[section_start].tolist()
            headers = [str(int(h)) if isinstance(h, float) and h.is_integer() else str(h) for h in headers]

            editable_cols = [col for col in subdf.columns if col in headers]

            for i, (_, sub_row) in enumerate(subdf.iterrows()):
                target_idx = section_start + 1 + i
                for col in editable_cols:
                    col_idx = headers.index(col)
                    full_df.iat[target_idx, col_idx] = sub_row[col]

        full_df.replace([np.nan, np.inf, -np.inf, None], "-", inplace=True)

        payload = {
            "model": self.model,
            "sheet_name": self.fuel_market,
            "data": full_df.to_dict(orient="records")
        }

        res = requests.post(self.save_url, json=payload)
        if res.status_code == 200:
            st.success(f"Changes in '{self.fuel_market}' saved successfully in '{self.model}'.")
        else:
            st.error(f"Save failed (HTTP {res.status_code}): {res.text}")

    def reset_sheet(self):
        payload = {
            "model": self.model,
            "sheet_name": self.fuel_market,
            "data": []
        }

        res = requests.post(self.reset_url, json=payload)
        if res.status_code == 200:
            st.success(f"File '{self.fuel_market}' of '{self.model}' was reset to template version.")
        else:
            st.error(res.json().get("error", "Something went wrong"))
    
    def show(self):
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()

        if st.button("Save"):
            self.save_changes()
            time.sleep(1)
            st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            time.sleep(1)
            st.rerun()


class TechnoEconomicInputs:
    
    def __init__(self, api_base, subsection, model, fuel_market):
        self.api_base = api_base
        self.fuel_market_get_url = f"{self.api_base}/fuel-market-information"
        self.get_url = f"{self.api_base}/technoeconomic-inputs"
        self.save_url = f"{self.api_base}/save-technoeconomic-inputs"
        self.reset_url = f"{self.api_base}/reset-technoeconomic-inputs"
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.df_fuel_market_inputs = None
        self.df_heights_fuel_market_inputs = {"E-Cooking": 170, "LPG": 260, "C02": 170}
        self.df = None
        self.df_heights = {"E-Cooking": 580, "LPG": 500, "Rest of subsidies or taxes": 410}
        self.edited_df = None
        self.non_editable_columns = ["Inputs", "Type", "Units"]
    
    def get_technoeconomic_inputs(self):
        try:
            url = f"{self.get_url}/{self.model}"
            res = requests.get(url)
            if res.status_code != 200:
                st.error("The file 'technoeconomic-inputs.xlsx' was not found.")
                return None

            file_data = BytesIO(res.content)
            xls = pd.ExcelFile(file_data)
            return xls.sheet_names

        except Exception as e:
            st.error(f"Error fetching techno-economic inputs: {e}")
            return None

    def fetch_and_load(self, fuel_market: bool):
        
        url = self.fuel_market_get_url if fuel_market else f"{self.get_url}/{self.model}"
        res = requests.get(url)
        if res.status_code != 200:
            if fuel_market:
                st.error("Could not load 'fuel-market-information.xlsx'")
            else:
                st.error("Could not load 'technoeconomic-inputs.xlsx'")
            return False
        
        if not self.fuel_market:
            st.warning("No sheet found in file for selected technology market.")
            return False
        
        file_data = BytesIO(res.content)
        if fuel_market:
            if not self.fuel_market == "C02":
                self.df_fuel_market_inputs = pd.read_excel(file_data, sheet_name=self.fuel_market, engine="openpyxl")
            else:
                self.df_fuel_market_inputs = pd.read_excel(file_data, sheet_name="Carbon", engine="openpyxl")
        else:
            self.df = pd.read_excel(file_data, sheet_name=self.fuel_market, engine="openpyxl")

        if fuel_market:
            for col in self.df_fuel_market_inputs.columns:
                if col not in ["Inputs", "Units"]:
                    self.df_fuel_market_inputs[col] = pd.to_numeric(self.df_fuel_market_inputs[col].replace("-", pd.NA), errors="coerce")
        else:
            for col in self.df.columns:
                if col not in ["Inputs", "Type", "Units"]:
                    self.df[col] = pd.to_numeric(self.df[col].replace("-", pd.NA), errors="coerce")

        return True

    def show_excel_editor(self, fuel_market: bool):
        df = self.df_fuel_market_inputs.copy() if fuel_market else self.df.copy()
        df.columns = df.columns.astype(str)

        if fuel_market and self.fuel_market == "LPG":
            df["Inputs"] = df["Inputs"].astype(str)
            df = df[~df["Inputs"].str.contains("EBITDA", na=False)]
            df["Inputs"] = df["Inputs"].str.replace("OPEX", "Tariff", regex=False)

        gb = GridOptionsBuilder.from_dataframe(df)

        if not fuel_market:
            for col in df.columns:
                if col not in self.non_editable_columns:
                    gb.configure_column(col, editable=True)
        else:
            gb.configure_columns(df.columns, editable=False)

        if fuel_market: 
            height = self.df_heights_fuel_market_inputs.get(self.fuel_market, self.df_heights_fuel_market_inputs["E-Cooking"])
        else:
            height = self.df_heights.get(self.fuel_market, self.df_heights["E-Cooking"])

        grid_options = gb.build()

        grid_response = AgGrid(
            df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.VALUE_CHANGED if not fuel_market else GridUpdateMode.NO_UPDATE,
            allow_unsafe_jscode=True,
            enable_enterprise_modules=False,
            height=height
        )

        if not fuel_market:
            self.edited_df = pd.DataFrame(grid_response["data"])

    def save_changes(self):
        df = self.edited_df.copy()
        df.columns = df.columns.map(str)
        self.df.columns = self.df.columns.map(str)
        df = df[self.df.columns]
        
        for col in df.columns:
            if col not in self.non_editable_columns:
                df[col] = df[col].fillna("-")

        data_json = df.reset_index(drop=True).to_dict(orient="records")
        payload = {
            "model": self.model,
            "sheet_name": self.fuel_market,
            "data": data_json
        }
        res = requests.post(self.save_url, json=payload)
        if res.status_code == 200:
            st.success(f"Changes in '{self.fuel_market}' saved successfully.")
        else:
            st.error("Error saving the changes, try again later.")

    def reset_sheet(self):
        payload = {
            "model": self.model,
            "sheet_name": self.fuel_market,
            "data": []
        }
        response = requests.post(self.reset_url, json=payload)
        if response.status_code == 200:
            st.success(f"File '{self.fuel_market}' reset to template version.")
        else:
            st.error(response.json().get("error", "Something went wrong"))

    def show(self):
        
        st.subheader(f"{self.fuel_market} Inputs")
        if self.fuel_market == "Rest of subsidies or taxes":
            if not self.fetch_and_load(fuel_market=False):
                return
            self.show_excel_editor(fuel_market=False)
            if st.button("Save"):
                self.save_changes()
                time.sleep(1)
                st.rerun()

            if st.button("Reset"):
                self.reset_sheet()
                time.sleep(1)
                st.rerun()
            return

        
        if self.fuel_market == "C02":
            if not self.fetch_and_load(fuel_market=True):
                return
            self.show_excel_editor(fuel_market=True)
            return
    
        if not self.fetch_and_load(fuel_market=True):
            return
        
        self.show_excel_editor(fuel_market=True)

        if not self.fetch_and_load(fuel_market=False):
            return
        
        self.show_excel_editor(fuel_market=False)

        if st.button("Save"):
            self.save_changes()
            time.sleep(1)
            st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            time.sleep(1)
            st.rerun()


class SummaryFinancing:
    def __init__(self, api_base):
        self.api_base = api_base

    def show(self):
        pass
