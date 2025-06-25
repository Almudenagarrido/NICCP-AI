import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import pandas as pd
import numpy as np
import requests
import uuid
from io import BytesIO


class GeneralInformation:

    def __init__(self, api_url, subsection):
        self.api_base = api_url
        self.get_url = f"{self.api_base}/general-information"
        self.post_url = f"{self.api_base}/save-general-information"
        self.subsection = subsection
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
        self.df_heights = {"Electricity": 170, "LPG": 350, "Carbon": 170}

    def get_financial_markets(self):
        try:
            res = requests.get(self.get_url)
            if res.status_code != 200:
                st.error("The file 'general-information.xlsx' was not found.")
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
            st.error("Could not load 'general-information.xlsx'")
            return False

        if not self.subsection:
            st.write(self.subsection)
            st.warning("No sheet found in file for selected subsection.")
            return False

        file_data = BytesIO(res.content)
        self.df = pd.read_excel(file_data, sheet_name=self.subsection, engine="openpyxl")

        for col in self.df.columns:
            if col not in ["Inputs", "Units"]:
                self.df[col] = pd.to_numeric(self.df[col].replace("-", pd.NA), errors="coerce")
        
        return True

    def show_excel_editor(self):
        st.subheader(self.subsection)
        df = self.df.copy()
        gb = GridOptionsBuilder.from_dataframe(df)
        height = self.df_heights[self.subsection] if self.subsection in self.df_heights else self.df_heights["Electricity"]
        
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
            if col not in ["Inputs", "Units"]:
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
            "sheet_name": self.subsection,
            "data": data_json
        }
        res = requests.post(self.post_url, json=payload)
        if res.status_code == 200:
            st.success("Changes saved correctly in 'general-information.xlsx'")
        else:
            st.error("Error saving the changes, try again later.")

    def reset_sheet(self):
        response = requests.post("http://localhost:8000/reset-general-information", json={"model": "", "sheet_name": self.subsection, "data": []})
        if response.status_code == 200:
            st.success(response.json().get("message", "Sheet reset"))
        else:
            st.error(response.json().get("error", "Something went wrong"))

    def add_market(self):
        st.subheader("Add New Market")
        new_market = st.text_input("Enter name of the new technology: ")

        if st.button("Create Market"):
            new_market = new_market.strip()

            if new_market == "":
                st.warning("Market name cannot be empty.")
            else:
                try:
                    response = requests.post(
                        "http://localhost:8000/add-market", 
                        json={"name": new_market}
                    )
                    if response.status_code == 200:
                        st.success(f"Market '{new_market}' added successfully.")
                        st.session_state.subsection = f"{new_market}"
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
        url = "http://localhost:8000/delete-market"
        try:
            response = requests.post(url, json={"name": sheet_name})
            if response.status_code == 200:
                return True
            else:
                st.error(f"Failed to delete '{sheet_name}'. Server responded with {response.status_code}.")
                return False
        except Exception as e:
            st.error(f"Error deleting market: {e}")
            return False

    def __call__(self):

        if self.subsection == "Add":
            self.add_market()
            return
        
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()
        
        if st.button("Save"):
            self.save_changes()
            st.rerun()

        if st.button("Reset"):
            self.reset_sheet()
            st.rerun()


class TechnoEconomicModels:
    def __init__(self, api_url, subsection, model, design_market):
        self.api_url = api_url
        self.subsection = subsection
        self.subsections = {
            "Manage": ManageModels(self.api_url),
            "Design Capital Structure": DesignCapitalStructure(api_url, subsection, model, design_market),
            "Summary Financing": SummaryFinancing(self.api_url)
        }
        self.model = model
        self.design_market = design_market

    def __call__(self):
        if self.subsection in self.subsections:
            self.subsections[self.subsection].show()
        else:
            st.info("Please select a valid subsection.")


class ManageModels:
    def __init__(self, api_base):
        self.api_base = api_base
        self.list_url = f"{self.api_base}/technoeconomic-models"
        self.get_url = f"{self.api_base}/technoeconomic-models"
        self.delete_url = f"{self.api_base}/delete-technoeconomic-model"
        self.upload_url = f"{self.api_base}/upload-technoeconomic-model"
        self.valid_extensions = ["xlsx", "xlsm", "xls", "xltx", "xltm"]

    def list_technoeconomic_models(self):
        try:
            res = requests.get(self.list_url)
            return res.json().get("files", []) if res.status_code == 200 else []
        except:
            st.error("Error loading models from server.")
            return []
    
    def get_technoeconomic_model(self, filename):
        res = requests.get(f"{self.get_url}/{filename}")
        if res.status_code == 200:
            return res.content
        else:
            return None

    def remove_extension(self, filename):
        for ext in self.valid_extensions:
            if filename.lower().endswith(f".{ext}"):
                return filename[: -len(ext) - 1]
        return filename
    
    def show_technoeconomic_models(self, files):
        success, message = None, ""
        for model in files:
            clean_name = self.remove_extension(model)
            col1, col2, col3 = st.columns([0.8, 0.1, 0.1])

            with col1:
                if st.button(f"📄 {clean_name}"):
                    st.session_state.page = "Techno-Economic Models"
                    st.session_state.subsection = "Design Capital Structure"
                    st.session_state.model = clean_name
                    st.rerun()

            with col2:
                content = self.get_technoeconomic_model(model)
                if content:
                    st.download_button(
                        label="⬇️",
                        data=content,
                        file_name=model,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            with col3:
                if st.button("❌", key=f"delete_{model}"):
                    success, message = self.delete_technoeconomic_model(model)

            if success is True:
                st.success(message)
                st.rerun()
            elif success is False:
                st.error(message)
                st.rerun()

    def delete_technoeconomic_model(self, filename):
        res = requests.delete(f"{self.delete_url}/{filename}")
        if res.status_code == 200:
            return True, f"Model '{filename}' deleted successfully."
        else:
            return False, "Failed to delete the model."

    def tecnoeconomic_model_uploader(self):
        st.markdown("---")
        st.markdown("##### Upload New Techno-Economic Model")
        if "uploader_key" not in st.session_state:
            st.session_state["uploader_key"] = str(uuid.uuid4())
        uploaded_file = st.file_uploader(
            "Upload file",
            type=self.valid_extensions,
            key=st.session_state["uploader_key"],
            label_visibility="collapsed"
        )

        if uploaded_file is not None:
            if any(uploaded_file.name.lower().endswith(f".{ext}") for ext in self.valid_extensions):
                upload_success = self.upload_technoeconomic_model(uploaded_file)
                if upload_success:
                    st.success(f"File '{uploaded_file.name}' uploaded successfully.")
                    st.session_state["uploader_key"] = str(uuid.uuid4())
                    st.rerun()
            else:
                st.error(f"Only Excel files are allowed: {', '.join(self.valid_extensions)}")

    def upload_technoeconomic_model(self, file):
        files = {"file": (file.name, file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        res = requests.post(self.upload_url, files=files)

        if res.status_code == 200:
            return True
        else:
            error_detail = res.json().get("detail", res.text)
            st.error(f"Upload failed: {error_detail}")
            return False

    def show(self):
        st.subheader("Manage Techno-Economic Models")
        files = self.list_technoeconomic_models()

        if not files:
            st.info("No techno-economic models available.")

        self.show_technoeconomic_models(files)
        self.tecnoeconomic_model_uploader()
        

class DesignCapitalStructure:
    def __init__(self, api_base, subsection, model, design_market):
        self.api_base = api_base
        self.get_url = f"{self.api_base}/design-capital-structure"
        self.save_url = f"{self.api_base}/save-design-capital-structure"
        self.reset_url = f"{self.api_base}/reset-design-capital-structure"
        self.subsection = subsection
        self.model = model
        self.design_market = design_market
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
        design_market, engine="openpyxl", header=None, index_col=None)
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
        st.subheader(f"{self.design_market} Financial Plan")

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
            "sheet_name": self.design_market,
            "data": full_df.to_dict(orient="records")
        }

        res = requests.post(self.save_url, json=payload)
        if res.status_code == 200:
            st.success(f"Saved successfully in model '{self.model}'")
        else:
            st.error(f"Save failed (HTTP {res.status_code}): {res.text}")

    def reset_sheet(self):
        payload = {
            "model": self.model,
            "sheet_name": self.design_market,
            "data": []
        }

        res = requests.post(self.reset_url, json=payload)
        if res.status_code == 200:
            st.success(res.json().get("message", "Sheet reset"))
            st.rerun()
        else:
            st.error(res.json().get("error", "Something went wrong"))
    def show(self):
        if not self.fetch_and_load():
            return
        
        self.show_excel_editor()

        if st.button("Save"):
            self.save_changes()

        if st.button("Reset"):
            self.reset_sheet()


class SummaryFinancing:
    def __init__(self, api_base):
        self.api_base = api_base

    def show(self):
        pass
