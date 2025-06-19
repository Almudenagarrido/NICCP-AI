import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import pandas as pd
import requests
import uuid
from io import BytesIO


class GeneralInformation:

    def __init__(self, subsection):
        self.get_url = "http://localhost:8000/general-information"
        self.post_url = "http://localhost:8000/save-general-information"
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
            "sheet_name": self.subsection,
            "data": data_json
        }
        res = requests.post(self.post_url, json=payload)
        if res.status_code == 200:
            st.success("Changes saved correctly in 'general-information.xlsx'")
        else:
            st.error("Error saving the changes, try again later.")

    def reset_sheet(self):
        response = requests.post("http://localhost:8000/reset-general-information", json={"sheet_name": self.subsection, "data": []})
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

        if st.button("Reset"):
            self.reset_sheet()
            st.rerun()


class TechnoEconomicModels:
    def __init__(self, subsection):
        self.subsection = subsection
        self.subsections = {
            "Manage": ManageModels(),
            "Design Capital Structure": DesignCapitalStructure(),
            "Summary Financing": SummaryFinancing()
        }

    def __call__(self):
        if self.subsection in self.subsections:
            self.subsections[self.subsection].show()
        else:
            st.info("Seleccione una subsección válida.")


class ManageModels:
    def __init__(self):
        self.api_base = "http://localhost:8000"
        self.list_url = f"{self.api_base}/technoeconomic-models"
        self.get_url = f"{self.api_base}/technoeconomic-models"
        self.delete_url = f"{self.api_base}/delete-technoeconomic-model"
        self.upload_url = f"{self.api_base}/upload-technoeconomic-model"

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

    def show_technoeconomic_models(self, files):
        success, message = None, ""
        for model in files:
            clean_name = model.replace(".xlsx", "")
            col1, col2, col3 = st.columns([0.8, 0.1, 0.1])

            with col1:
                if st.button(f"📄 {clean_name}"):
                    st.session_state.model = model
                    st.session_state.page = "Techno-Economic Models"
                    st.session_state.subsection = "Design Capital Structure"
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
            "", 
            type=["xlsx", "xlsm", "xls", "xltx", "xltm"], 
            key=st.session_state["uploader_key"]
        )
        valid_excel_extensions = (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")

        if uploaded_file is not None:
            if uploaded_file.name.endswith(valid_excel_extensions):
                upload_success = self.upload_technoeconomic_model(uploaded_file)
                if upload_success:
                    st.success(f"File '{uploaded_file.name}' uploaded successfully.")
                    st.session_state["uploader_key"] = str(uuid.uuid4())
                    st.rerun()
                else:
                    st.error("Upload failed.")
            else:
                st.error("Only Excel files are allowed (.xlsx, .xlsm, .xls, .xltx, .xltm).")

    def upload_technoeconomic_model(self, file):
        try:
            files = {"file": (file.name, file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            res = requests.post(self.upload_url, files=files)
            return res.status_code == 200
        except Exception as e:
            st.error(f"Upload error: {e}")
            return False

    def show(self):
        st.subheader("Manage Techno-Economic Models")
        files = self.list_technoeconomic_models()

        if not files:
            st.info("No techno-economic models available.")

        self.show_technoeconomic_models(files)
        self.tecnoeconomic_model_uploader()
        

class DesignCapitalStructure:
    def __init__(self):
        self.api_base = "api_base"

    def show(self):
        pass


class SummaryFinancing:
    def __init__(self):
        self.api_base = ""

    def show(self):
        pass
