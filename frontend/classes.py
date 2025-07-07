import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import pandas as pd
import numpy as np
import requests
from io import BytesIO
import time


class CellValidator:
    
    def __init__(self):
        self.validators = {
            "%": self.validate_percentage,
            "days": self.validate_positive_integer,
            "years": self.validate_positive_integer,
            "$ / ton": self.validate_positive_integer,
            "m$": self.validate_positive_integer,
        }

    def validate_percentage(self, value):
        try:
            if pd.isna(value) or value == "-":
                return None
            val = float(value)
            if not (0 <= val <= 100):
                return "Value must be between 0 and 100 (percentage)"
        except:
            return "Not a valid number"
        return None

    def validate_positive_integer(self, value):
        try:
            if pd.isna(value) or value == "-":
                return None
            val = float(value)
            if val < 0 or not float(val).is_integer():
                return "Value must be a positive integer"
        except:
            return "Not a valid number"
        return None

    def validate_cell(self, value, unit):
        validator = self.validators.get(unit.strip().lower())
        if validator:
            return validator(value)
        return None
    
    def cell_validations(self, df, editable_columns):
        invalid_cells = []

        unit_col = next((col for col in df.columns if "units" in col.lower()), None)

        for index, row in df.iterrows():
            units = str(row[unit_col]).lower() if unit_col else ""
            input_name = row["Inputs"] if "Inputs" in df.columns else int(index) + 1

            for col in df.columns:
                if col not in editable_columns:
                    continue

                value = row[col]
                error = self.validate_cell(value, units)
                if error:
                    invalid_cells.append((index, input_name, col, value, error))

        return invalid_cells
    

class FuelMarketInformation:

    def __init__(self, api_url, fuel_market):
        self.api_base = api_url
        self.fuel_market = fuel_market
        self.cell_validator = CellValidator()
        self.get_url = f"{self.api_base}/fuel-market-information"
        self.save_url = f"{self.api_base}/save-fuel-market-information"
        self.reset_url = f"{self.api_base}/reset-fuel-market-information"
        self.add_url = f"{self.api_base}/add-fuel-market"
        self.delete_url = f"{self.api_base}/delete-fuel-market"
        self.df = None
        self.edited_df = None
        self.editable_columns = ["Baseline"] + [str(year) for year in range(2020, 2051)]
        self.df_heights = {"Electricity": 170, "LPG": 290, "Carbon": 170}

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
        height = self.df_heights[self.fuel_market] if self.fuel_market in self.df_heights else self.df_heights["Electricity"]
        
        for col in df.columns:
            if col in self.editable_columns:
                gb.configure_column(col, editable=True)
            else:
                gb.configure_column(col, editable=False)

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
           
    def save_changes(self):
        df = self.edited_df.copy()
        df = df[self.df.columns]

        for col in df.columns:
            if col in self.editable_columns:
                df[col] = df[col].fillna(0)

        for i, row in df.iterrows():
            if str(row.get("Inputs", "")).strip().lower() == "number of years that you could sell those carbon credits":
                for col in df.columns:
                    if col not in ["Inputs", "Units", "Baseline"]:
                        df.at[i, col] = pd.NA
        
        editable_cols_in_df = [col for col in self.editable_columns if col in df.columns]
        df[editable_cols_in_df] = df[editable_cols_in_df].astype(object)
        df.loc[df.index[-1], editable_cols_in_df] = df.loc[df.index[-1], editable_cols_in_df].fillna("-")

                
        invalid_cells = self.cell_validator.cell_validations(df, self.editable_columns)

        if invalid_cells:
            for _, input_name, col, value, error in invalid_cells:
                st.warning(f"Invalid value '{value}' in row '{input_name}' (column '{col}'): {error}")
            return False

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

        return error

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
            saved = self.save_changes()
            if saved:
                time.sleep(2)
                st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            time.sleep(2)
            st.rerun()


class TechnoEconomicModels:
    
    def __init__(self, api_url, subsection, model, fuel_market):
        self.api_url = api_url
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.cell_validator = CellValidator()
        self.subsections = {
            "Manage Techno-Economic Models": ManageModels(self.api_url),
            "Design Capital Structure": DesignCapitalStructure(api_url, subsection, model, fuel_market, self.cell_validator),
            "Techno-Economic Inputs": TechnoEconomicInputs(api_url, subsection, model, fuel_market, self.cell_validator),
            "Carbon Credits": CarbonCredits(api_url, subsection, model, self.cell_validator),
            "Financial Statements": FinancialStatements(api_url, subsection, model, fuel_market),
            "Capex Fuel Market": CapexFuelMarket(api_url, subsection, model, fuel_market),
            "Summary Financing": SummaryFinancing(self.api_url)
        }

    def __call__(self):
        if self.subsection in self.subsections:
            self.subsections[self.subsection]()
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
            try:
                error_detail = res.json().get("detail", res.text)
            except ValueError:
                error_detail = res.text
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

        models = self.list_technoeconomic_models()
        bau_exists = any(self.remove_extension(m).lower() == "bau" for m in models)

        if not bau_exists:
            st.info("First create the BAU (Business As Usual) model.")
            with st.form("create_bau_form"):
                start_year = st.number_input("Start Year (BAU)", step=1, format="%d")
                end_year = st.number_input("End Year (BAU)", step=1, format="%d")

                upload_file = st.file_uploader("Upload Excel file for BAU (optional)", type=self.valid_extensions)

                create = st.form_submit_button("Create BAU Model")

                if create:
                    if start_year >= end_year:
                        st.error("Start year must be less than end year.")
                    else:
                        self.create_technoeconomic_model("bau", start_year, end_year)
                        st.success(f"BAU model created succesfully.")
                        if upload_file:
                            upload_success = self.upload_technoeconomic_model("BAU", upload_file)
                            if upload_success:
                                st.success("Excel file successfully uploaded for BAU.")
                            else:
                                st.error("Failed to upload Excel file for BAU.")
                        time.sleep(1)
                        st.rerun()

        else:
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
                    elif name.strip().lower() == "bau":
                        st.error("The model name 'BAU' is reserved for the Business As Usual model.")
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
                    if clean_name == "BAU":
                        st.session_state.page = "Techno-Economic Models"
                        st.session_state.subsection = "Carbon Credits"
                    else:
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
                            st.success(f"Information uploaded to 'Techno-Economic Inputs' file for model '{model}'")
                            st.session_state[f"show_uploader_{model}"] = False
                            time.sleep(2)
                            st.rerun()
                    else:
                        st.error(f"Only Excel files are allowed: {', '.join(self.valid_extensions)}")

    def __call__(self):
        st.subheader("Manage Techno-Economic Models")
        models = self.list_technoeconomic_models()
        if not models:
            st.info("No techno-economic models available.")
        else:
            self.show_technoeconomic_models(models)
        self.model_creator()


class DesignCapitalStructure:
    
    def __init__(self, api_base, subsection, model, fuel_market, cell_validator):
        self.api_base = api_base
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.cell_validator = cell_validator
        self.get_url = f"{self.api_base}/design-capital-structure"
        self.save_url = f"{self.api_base}/save-design-capital-structure"
        self.reset_url = f"{self.api_base}/reset-design-capital-structure"
        self.df = None
        self.subtables = {"Finance": None, "Total": None, "Division": None, "Debts": None}
        self.subtable_titles = {
            "Finance": "Financiation",
            "Total": "Total",
            "Division": "Division",
            "Debts": "Type"
        }
        self.subtable_heights = {
            "Finance": 100,
            "Total": 130,
            "Division": 230,
            "Debts": 150
        }
        self.editable_columns = {
            "Finance": ["Amount"],
            "Total": ["Amount"],
            "Division": ["Amount"],
            "Debts": ["Baseline"] + [str(year) for year in range(2020, 2051)]
        }

    def get_design_capital(self):
        try:
            url = f"{self.get_url}/{self.model}"
            res = requests.get(url)
            if res.status_code != 200:
                st.error(f"The file 'design-capital-structure-{self.model}.xlsx' was not found.")
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
            first_col_val = str(row[0]) if pd.notna(row[0]) else ""
            for key, keyword in self.subtable_titles.items():
                if key not in section_starts and keyword.lower() in first_col_val.lower():
                    section_starts[key] = idx

        sorted_sections = sorted(section_starts.items(), key=lambda x: x[1])
        for i, (key, start_idx) in enumerate(sorted_sections):
            end_idx = sorted_sections[i + 1][1] if i + 1 < len(sorted_sections) else len(df)
            subdf = df.iloc[start_idx:end_idx].reset_index(drop=True)
            
            subdf = subdf.mask(subdf == "-", np.nan)
            subdf.dropna(axis=0, how="all", inplace=True)
            subdf.dropna(axis=1, how="all", inplace=True)
            
            if not subdf.empty:
                header_row = subdf.iloc[0]
                subdf.columns = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in header_row]
                subdf = subdf[1:].reset_index(drop=True)
                subdf = subdf.loc[:, subdf.columns.notna()]
                subdf = subdf.loc[:, subdf.columns.str.strip() != '']
                subdf = subdf.loc[:, ~subdf.columns.duplicated()]
                subdf = subdf.loc[:, ~subdf.columns.str.contains("^Unnamed", case=False)]
            
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

        for i, row in full_df.iterrows():
            if str(row[0]).strip().lower() == "3.- debt":
                for col in full_df.columns:
                    if col != 0:
                        full_df.at[i, col] = pd.NA

        for col in full_df.columns:
            full_df[col] = full_df[col].astype(object)
            full_df[col] = full_df[col].fillna('-')

        for key, title in self.subtable_titles.items():
            subdf = self.subtables.get(key)
            if subdf is None or subdf.empty:
                continue

            editable_cols = self.editable_columns.get(key, [])

            invalid_cells = self.cell_validator.cell_validations(subdf, editable_cols)
            if invalid_cells:
                for _, input_name, col, value, error in invalid_cells:
                    st.warning(f"[{key}] Invalid value '{value}' in row '{input_name}' (column '{col}'): {error}")
                return False

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

        return True

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
    
    def __call__(self):
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()

        if st.button("Save"):
            saved = self.save_changes()
            if saved:
                time.sleep(2)
                st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            time.sleep(2)
            st.rerun()


class TechnoEconomicInputs:
    
    def __init__(self, api_base, subsection, model, fuel_market, cell_validator):
        self.api_base = api_base
        self.fuel_market_get_url = f"{self.api_base}/fuel-market-information"
        self.get_url = f"{self.api_base}/technoeconomic-inputs"
        self.save_url = f"{self.api_base}/save-technoeconomic-inputs"
        self.reset_url = f"{self.api_base}/reset-technoeconomic-inputs"
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.df_fuel_market_inputs = None
        self.df_heights_fuel_market_inputs = {"Electricity": 170, "LPG": 230, "C02": 170}
        self.df = None
        self.df_heights = {"Electricity": 580, "LPG": 500, "Rest of subsidies or taxes": 410}
        self.edited_df = None
        self.editable_columns = ["Baseline"] + [str(year) for year in range(2020, 2051)]
        self.cell_validator = cell_validator
    
    def get_technoeconomic_inputs(self):
        try:
            url = f"{self.get_url}/{self.model}"
            res = requests.get(url)
            if res.status_code != 200:
                st.error("The file 'technoeconomic-inputs.xlsx' was not found.")
                return None

            file_data = BytesIO(res.content)
            with pd.ExcelFile(file_data) as xls:
                return xls.sheet_names

        except Exception as e:
            st.error(f"Error fetching techno-economic inputs: {e}")
            return None

    def fetch_and_load(self, fuel_market: bool):
        
        url = self.fuel_market_get_url if fuel_market else f"{self.get_url}/{self.model}"
        res = requests.get(url)
        if res.status_code != 200:
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

        if not fuel_market:
            for col in self.df.columns:
                if col not in ["Inputs", "Type", "Units"]:
                    self.df[col] = pd.to_numeric(self.df[col].replace("-", pd.NA), errors="coerce")
            
            self.df = self.df.mask(self.df == "-", np.nan)

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
                if col in self.editable_columns:
                    gb.configure_column(col, editable=True)
        else:
            gb.configure_columns(df.columns, editable=False)

        if fuel_market: 
            height = self.df_heights_fuel_market_inputs.get(self.fuel_market, self.df_heights_fuel_market_inputs["Electricity"])
        else:
            height = self.df_heights.get(self.fuel_market, self.df_heights["Electricity"])

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
        df["Inputs"] = df["Inputs"].astype(str).str.strip()

        for i, row in df.iterrows():
            input_val = row["Inputs"].lower()

            if "output" in input_val:
                for col in df.columns:
                    if col not in ["Inputs", "Type", "Units"]:
                        df.at[i, col] = np.nan

            elif "depreciation" in input_val:
                for col in df.columns:
                    if col not in ["Inputs", "Type", "Units", "Baseline"]:
                        df.at[i, col] = np.nan

        df.columns = df.columns.map(str)
        self.df.columns = self.df.columns.map(str)
        df = df[self.df.columns]
        
        for col in df.columns:
            if col in self.editable_columns:
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
            return False
        
        return True

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

    def __call__(self):
        
        st.subheader(f"{self.fuel_market} Inputs")
        if self.fuel_market == "Rest of subsidies or taxes":
            if not self.fetch_and_load(fuel_market=False):
                return
            self.show_excel_editor(fuel_market=False)
            if st.button("Save"):
                saved = self.save_changes()
                if saved:
                    time.sleep(2)
                    st.rerun()

            if st.button("Reset"):
                self.reset_sheet()
                time.sleep(2)
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


class CarbonCredits:
    def __init__(self, api_base, subsection, model, cell_validator):
        self.api_base = api_base
        self.subsection = subsection
        self.model = model
        self.cell_validator = cell_validator
        self.get_url = f"{self.api_base}/carbon-credits"
        self.save_url = f"{self.api_base}/save-carbon-credits"
        self.reset_url = f"{self.api_base}/reset-carbon-credits"
        self.df = None
        self.edited_df = None
        self.editable_columns = ["Baseline"] + [str(year) for year in range(2020, 2051)]

    def get_carbon_credits(self):
        try:
            res = requests.get(self.get_url)
            if res.status_code != 200:
                st.error("The file 'carbon-credits.xlsx' was not found.")
                return []
            file_data = BytesIO(res.content)
            xls = pd.ExcelFile(file_data)
            return xls.sheet_names
        except Exception as e:
            st.error(f"Error fetching 'carbon-credits' file: {e}")
            return []
    
    def fetch_and_load(self):
        res = requests.get(self.get_url)
        if res.status_code != 200:
            st.error("Could not load 'carbon-credits.xlsx'")
            return False

        file_data = BytesIO(res.content)
        self.df = pd.read_excel(file_data, sheet_name="Carbon Credits", engine="openpyxl")

        for col in self.df.columns:
            if col in self.editable_columns:
                self.df[col] = pd.to_numeric(self.df[col].replace("-", pd.NA), errors="coerce")
        
        self.df.dropna(how="all", inplace=True)
        self.df.dropna(axis=1, how="all", inplace=True)

        return True
    
    def show_excel_editor(self):

        st.subheader(f"Carbon Credits")
        df = self.df.copy()

        df_filtered = df[~df.apply(lambda row: row.astype(str).str.contains(r"\{model\}", regex=True).any(), axis=1)]

        gb = GridOptionsBuilder.from_dataframe(df_filtered)
        
        for col in df_filtered.columns:
            col_str = str(col)
            if col_str in self.editable_columns:
                gb.configure_column(col_str, editable=True)
            else:
                gb.configure_column(col_str, editable=False)

        grid_options = gb.build()
        grid_response = AgGrid(
            df_filtered,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            allow_unsafe_jscode=True,
            enable_enterprise_modules=False,
            height=320
        )
        self.edited_df = pd.DataFrame(grid_response["data"])
    
    def save_changes(self):
        self.df.columns = self.df.columns.map(str)
        df = self.edited_df.copy()
        df = df[self.df.columns]

        for col in df.columns:
            col_str = str(col)
            if col_str in self.editable_columns:
                df[col_str] = df[col_str].fillna(0)
                
        invalid_cells = self.cell_validator.cell_validations(df, self.editable_columns)

        if invalid_cells:
            for _, input_name, col, value, error in invalid_cells:
                st.warning(f"Invalid value '{value}' in row '{input_name}' (column '{col}'): {error}")
            return False

        data_json = df.reset_index(drop=True).to_dict(orient="records")
        payload = {
            "model": "",
            "sheet_name": "Carbon Credits",
            "data": data_json
        }
        res = requests.post(self.save_url, json=payload)
        if res.status_code == 200:
            st.success(f"Changes saved successfully in 'carbon-credits.xlsx'.")
        else:
            st.error("Error saving the changes, try again later.")

        return True
    
    def reset_sheet(self):
        response = requests.post(self.reset_url, json={"model": "", "sheet_name": "Carbon Credits", "data": []})
        if response.status_code == 200:
            st.success(f"File 'carbon-credits' reset to template version.")
        else:
            st.error(response.json().get("error", "Something went wrong"))

    def __call__(self):
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()
        
        if st.button("Save"):
            saved = self.save_changes()
            if saved:
                time.sleep(2)
                st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            time.sleep(2)
            st.rerun()


class FinancialStatements:
    
    def __init__(self, api_base, subsection, model, fuel_market):
        self.api_base = api_base
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.get_url = f"{self.api_base}/financial-statements"
        self.df = None
        self.subtables = {"Profit & Loss": None, "Balance Sheet": None, "Cash Flow Statement": None, "PP&E - Capex": None, "Working Capital Calculations": None, "Equity Schedule": None, "Capital Structure": None}
        self.section_starts = {
            "Profit & Loss": 0,
            "Balance Sheet": 36,
            "Cash Flow Statement": 65,
            "PP&E - Capex": 84,
            "Working Capital Calculations": 92,
            "Equity Schedule": 102,
            "Capital Structure": 108
        }
        self.subtable_heights = {"Profit & Loss": 1070, "Balance Sheet": 870, "Cash Flow Statement": 580, "PP&E - Capex": 260, "Working Capital Calculations": 320, "Equity Schedule": 200, "Capital Structure": 400}
        if self.fuel_market == "Electricity (Low access)":
            self.subtables["Transferred Capex"] = None
            self.section_starts["Transferred Capex"] = 121
            self.subtable_heights["Transferred Capex"] = 170
    
    def get_financial_statements(self):
        try:
            url = f"{self.get_url}/{self.model}"
            res = requests.get(url)
            if res.status_code != 200:
                st.error(f"The file 'financial-statements-{self.model}.xlsx' was not found.")
                return []

            file_data = BytesIO(res.content)
            xls = pd.ExcelFile(file_data)
            return xls.sheet_names

        except Exception as e:
            st.error(f"Error fetching capital structure: {e}")
            return []

    def fetch_and_load(self):
        url = f"{self.get_url}/{self.model}"
        res = requests.get(url)
        if res.status_code != 200:
            st.error(f"Could not load 'financial-statements-{self.model}.xlsx'")
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
        sorted_sections = sorted(self.section_starts.items(), key=lambda x: x[1])

        for i, (key, start_idx) in enumerate(sorted_sections):
            end_idx = sorted_sections[i + 1][1] if i + 1 < len(sorted_sections) else len(df)
            subdf = df.iloc[start_idx:end_idx].reset_index(drop=True)
            
            subdf = subdf.mask(subdf == "-", np.nan)
            subdf.dropna(axis=0, how="all", inplace=True)
            subdf.dropna(axis=1, how="all", inplace=True)
            
            if not subdf.empty:
                header_row = subdf.iloc[0]
                subdf.columns = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in header_row]
                subdf = subdf[1:].reset_index(drop=True)
                subdf = subdf.loc[:, subdf.columns.notna()]
                subdf = subdf.loc[:, subdf.columns.str.strip() != '']
                subdf = subdf.loc[:, ~subdf.columns.duplicated()]
                subdf = subdf.loc[:, ~subdf.columns.str.contains("^Unnamed", case=False)]
            
            self.subtables[key] = subdf

    def show_excel_editor(self):
        st.subheader(f"{self.fuel_market} - FFSS")

        for key, df in self.subtables.items():
            st.markdown(f"##### {key}")
            if df is not None and not df.empty:
                df.columns = df.columns.astype(str)
                
                gb = GridOptionsBuilder.from_dataframe(df)
                for col in df.columns:
                
                    if pd.api.types.is_numeric_dtype(df[col]):
                        gb.configure_column(
                            col,
                            type=["numericColumn", "numberColumnFilter"]
                        )
                    else:
                        gb.configure_column(
                            col,
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

    def __call__(self):
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()


class CapexFuelMarket:
    
    def __init__(self, api_base, subsection, model, fuel_market):
        self.api_base = api_base
        self.subsection = subsection
        self.model = model
        self.fuel_market = fuel_market
        self.get_url = f"{self.api_base}/capex-fuel-market"
        self.df = None
        self.subtables = {
            "Electricity": {"Grid": None, "Off-Grid": None},
            "LPG": {"Local Processing": None, "Local Transport": None, "Upstream": None},
        }
        self.section_starts = {
            "Grid": 0, 
            "Off-Grid": 7, 
            "Local Processing": 0, 
            "Local Transport": 7, 
            "Upstream": 14
        }
    
    def get_capex_markets(self):
        try:
            url = f"{self.get_url}/{self.model}"
            res = requests.get(url)
            if res.status_code != 200:
                st.error(f"The file 'capex-market-{self.model}.xlsx' was not found.")
                return []

            file_data = BytesIO(res.content)
            xls = pd.ExcelFile(file_data)
            return xls.sheet_names

        except Exception as e:
            st.error(f"Error fetching capital structure: {e}")
            return []

    def get_active_subtables(self):
        if self.fuel_market == "Electricity":
            return ["Grid", "Off-Grid"]
        else:
            return ["Local Processing", "Local Transport", "Upstream"]

    def fetch_and_load(self):
        url = f"{self.get_url}/{self.model}"
        res = requests.get(url)
        if res.status_code != 200:
            st.error(f"Could not load 'capex-market-{self.model}.xlsx'")
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
        active_keys = self.get_active_subtables()

        if self.fuel_market not in self.subtables:
            self.subtables[self.fuel_market] = {key: None for key in active_keys}

        active_sections = {k: self.section_starts[k] for k in active_keys}
        sorted_sections = sorted(active_sections.items(), key=lambda x: x[1])

        for i, (key, start_idx) in enumerate(sorted_sections):
            if key not in active_keys:
                continue

            end_idx = sorted_sections[i + 1][1] if i + 1 < len(sorted_sections) else len(df)

            subdf = df.iloc[start_idx:end_idx].reset_index(drop=True)

            subdf = subdf.mask(subdf == "-", np.nan)
            subdf.dropna(axis=0, how="all", inplace=True)
            subdf.dropna(axis=1, how="all", inplace=True)

            if not subdf.empty:
                header_row = subdf.iloc[0]
                subdf.columns = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in header_row]
                subdf = subdf[1:].reset_index(drop=True)
                subdf = subdf.loc[:, subdf.columns.notna()]
                subdf = subdf.loc[:, subdf.columns.str.strip() != '']
                subdf = subdf.loc[:, ~subdf.columns.duplicated()]
                subdf = subdf.loc[:, ~subdf.columns.str.contains("^Unnamed", case=False)]

            self.subtables[self.fuel_market][key] = subdf

    def show_excel_editor(self):
        st.subheader(f"{self.fuel_market} - CAPEX")

        if self.fuel_market not in self.subtables:
            st.warning("No subtables found for this fuel market.")
            return

        for key, df in self.subtables[self.fuel_market].items():
            st.markdown(f"##### {key}")
            if df is not None and not df.empty:
                df.columns = df.columns.astype(str)

                gb = GridOptionsBuilder.from_dataframe(df)
                for col in df.columns:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        gb.configure_column(col, type=["numericColumn", "numberColumnFilter"])
                    else:
                        gb.configure_column(col)

                grid_options = gb.build()
                AgGrid(
                    df,
                    gridOptions=grid_options,
                    update_mode=GridUpdateMode.VALUE_CHANGED,
                    allow_unsafe_jscode=True,
                    enable_enterprise_modules=False,
                    height=230,
                    reload_data=False
            )
                
    def __call__(self):
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()


class SummaryFinancing:
    def __init__(self, api_base):
        self.api_base = api_base

    def show(self):
        pass
