from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
import os
from openpyxl import load_workbook
from io import BytesIO
import shutil
import pandas as pd
from pydantic import BaseModel
from typing import List, Dict


GENERAL_INFORMATION_FOLDER = "./general-information"
GENERAL_INFORMATION_PATH = os.path.join(GENERAL_INFORMATION_FOLDER, "general-information.xlsx")
GENERAL_INFORMATION_TEMPLATE = os.path.join(GENERAL_INFORMATION_FOLDER, "general-information-template.xlsx")
TECHNOECONOMIC_MODELS_DIR = "technoeconomic-models"
DESIGN_CAPITAL_STRUCTURE_FOLDER = "./design-capital-structure"
DESIGN_CAPITAL_STRUCTURE_TEMPLATE = os.path.join(DESIGN_CAPITAL_STRUCTURE_FOLDER, "design-capital-structure-template.xlsx")
DESIGN_CAPITAL_STRUCTURE_MODEL = os.path.join(DESIGN_CAPITAL_STRUCTURE_FOLDER, "design-capital-structure-")
TECHNOECONOMIC_INPUTS_FOLDER = "./technoeconomic-inputs"
TECHNOECONOMIC_INPUTS_TEMPLATE = os.path.join(TECHNOECONOMIC_INPUTS_FOLDER, "technoeconomic-inputs-template.xlsx")
TECHNOECONOMIC_INPUTS_MODEL = os.path.join(TECHNOECONOMIC_INPUTS_FOLDER, "technoeconomic-inputs-")

VALID_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")
START_YEAR_TECHNO_MODELS = None
END_YEAR_TECHNO_MODELS = None


class SheetUpdate(BaseModel):
    model:str
    sheet_name: str
    data: List[Dict] 

class MarketName(BaseModel):
    name: str

app = FastAPI()


# General Information
@app.get("/general-information")
def get_general_information():
    global START_YEAR_TECHNO_MODELS, END_YEAR_TECHNO_MODELS
    
    techno_models = {
        os.path.splitext(f)[0] 
        for f in os.listdir(TECHNOECONOMIC_MODELS_DIR) 
        if any(f.endswith(ext) for ext in VALID_EXTENSIONS)
    }

    if not os.path.exists(GENERAL_INFORMATION_PATH):
        if os.path.exists(GENERAL_INFORMATION_TEMPLATE):
            shutil.copy(GENERAL_INFORMATION_TEMPLATE, GENERAL_INFORMATION_PATH)
        else:
            return {"error": "Template file is missing"}
    
    wb = load_workbook(GENERAL_INFORMATION_PATH)

    for ws in wb.worksheets:
        year_row = None
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=row, column=col).value
                if isinstance(cell_val, str) and cell_val.strip().lower() == "baseline":
                    year_row = row
                    break
            if year_row:
                break
        
        if year_row is None:
            continue

        cols_to_delete = []
        for col in range(3, ws.max_column + 1):
            cell_val = ws.cell(row=year_row, column=col).value
            try:
                year = int(cell_val)
                if year < START_YEAR_TECHNO_MODELS or year > END_YEAR_TECHNO_MODELS:
                    cols_to_delete.append(col)
            except (TypeError, ValueError):
                pass

        for col_idx in reversed(cols_to_delete):
            ws.delete_cols(col_idx)

        if ws.title == "LPG":
            existing_rows = {}

            for i in range(2, ws.max_row + 1):
                val = ws.cell(row=i, column=1).value
                if val:
                    key = str(val).strip()
                    existing_rows[key] = i
            
            expected_rows = {
                f"% EBITDA Margin - {model}" for model in techno_models
            }.union({
                f"OPEX subsidies - {model}" for model in techno_models
            })

            for row_label, row_num in list(existing_rows.items())[::-1]:
                if (row_label.startswith("% EBITDA Margin - ") or row_label.startswith("OPEX subsidies - ")) and row_label not in expected_rows:
                    ws.delete_rows(row_num)

            existing_inputs = {
                str(ws.cell(row=i, column=1).value).strip()
                for i in range(2, ws.max_row + 1)
                if ws.cell(row=i, column=1).value
            }

            for model_name in techno_models:
                for row_label in [f"% EBITDA Margin - {model_name}", f"OPEX subsidies - {model_name}"]:
                    if row_label not in existing_inputs:
                        new_row = ws.max_row + 1
                        ws.cell(row=new_row, column=1).value = row_label
                        ws.cell(row=new_row, column=2).value = "%"
                        for col in range(3, ws.max_column + 1):
                            ws.cell(row=new_row, column=col).value = 0

    wb.save(GENERAL_INFORMATION_PATH)

    return FileResponse(
        GENERAL_INFORMATION_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="general-information.xlsx"
    )

@app.post("/save-general-information")
async def save_general_information(update: SheetUpdate):

    df_new = pd.DataFrame(update.data)

    if 'index' in df_new.columns:
        df_new.set_index('index', inplace=True)

    excel_path = GENERAL_INFORMATION_PATH
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_new.to_excel(writer, sheet_name=update.sheet_name, index=False)

    return {"message": "The file with the general information was updated successfully"}

@app.post("/reset-general-information")
def reset_general_information(update: SheetUpdate):
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        return {"error": "General information file does not exist"}
    if not os.path.exists(GENERAL_INFORMATION_TEMPLATE):
        return {"error": "Template file is missing"}

    try:
        if update.sheet_name not in {"Carbon", "LPG"}:
            template_sheet = "Electricity"
        else:
            template_sheet = update.sheet_name

        template_df = pd.read_excel(GENERAL_INFORMATION_TEMPLATE, sheet_name=template_sheet)
        wb = load_workbook(GENERAL_INFORMATION_PATH)

        if update.sheet_name in wb.sheetnames:
            std = wb[update.sheet_name]
            wb.remove(std)

        wb.save(GENERAL_INFORMATION_PATH)

        with pd.ExcelWriter(GENERAL_INFORMATION_PATH, engine='openpyxl', mode='a') as writer:
            template_df.to_excel(writer, sheet_name=update.sheet_name, index=False)

        return {"message": f"'{update.sheet_name}' was reset to template version"}

    except Exception as e:
        return {"error": f"Failed to reset: {str(e)}"}

@app.post("/add-market")
def add_market(market: MarketName):
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        return {"error": "general-information.xlsx not found"}

    xls_template = pd.ExcelFile(GENERAL_INFORMATION_TEMPLATE, engine="openpyxl")
    xls = pd.ExcelFile(GENERAL_INFORMATION_PATH, engine="openpyxl")
    df_electricity = pd.read_excel(xls_template, sheet_name="Electricity", engine="openpyxl")
    new_sheet_name = market.name
    
    if "Electricity" not in xls_template.sheet_names:
        return {"error": "Base sheet 'Electricity' not found"}

    if new_sheet_name in xls.sheet_names:
        return {"error": f"The sheet '{new_sheet_name}' already exists."}

    sheets_dict = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    sheets_dict[new_sheet_name] = df_electricity.copy()

    with pd.ExcelWriter(GENERAL_INFORMATION_PATH, engine="openpyxl", mode="w") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return {"message": f"Market '{new_sheet_name}' added successfully."}

@app.post("/delete-market")
def delete_market(market: MarketName):
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        return {"error": "general-information.xlsx not found"}

    xls = pd.ExcelFile(GENERAL_INFORMATION_PATH, engine="openpyxl")
    sheet_to_delete = market.name

    if sheet_to_delete not in xls.sheet_names:
        return {"error": f"The sheet '{sheet_to_delete}' does not exist."}

    remaining_sheets = {
        sheet: xls.parse(sheet)
        for sheet in xls.sheet_names if sheet != sheet_to_delete
    }

    with pd.ExcelWriter(GENERAL_INFORMATION_PATH, engine="openpyxl", mode="w") as writer:
        for name, df in remaining_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    return {"message": f"Market '{sheet_to_delete}' deleted successfully."}


# Techno-Economic Models
@app.get("/technoeconomic-models")
def list_techno_models():
    global START_YEAR_TECHNO_MODELS, END_YEAR_TECHNO_MODELS
    
    files = [f for f in os.listdir(TECHNOECONOMIC_MODELS_DIR) if f.endswith(VALID_EXTENSIONS)]

    if files and START_YEAR_TECHNO_MODELS == None:
            file_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, files[0])
            with open(file_path, "rb") as f:
                existing_bytes = f.read()
            file = load_workbook(filename=BytesIO(existing_bytes), data_only=True, read_only=True)
            START_YEAR_TECHNO_MODELS = file["Contents"]["G8"].value
            END_YEAR_TECHNO_MODELS = file["Contents"]["G24"].value

    return {"files": files}

@app.get("/technoeconomic-models/{filename}")
def get_techno_model(filename: str):
    model_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, filename)
    if os.path.exists(model_path):
        return FileResponse(
            model_path,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=filename
        )
    raise HTTPException(status_code=404, detail="Model not found")

@app.delete("/delete-technoeconomic-model/{filename}")
def delete_techno_model(filename: str):
    model_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, filename)
    if os.path.exists(model_path):
        os.remove(model_path)
        return {"status": "deleted"}
    raise HTTPException(status_code=404, detail="Model not found")

@app.post("/upload-technoeconomic-model")
def upload_techno_model(file: UploadFile = File(...)):
    global START_YEAR_TECHNO_MODELS, END_YEAR_TECHNO_MODELS

    if not file.filename.endswith(VALID_EXTENSIONS):
        raise HTTPException(status_code=400, detail="Only Excel files are allowed.")
    
    contents = file.file.read()
    new_file = load_workbook(filename=BytesIO(contents), data_only=True, read_only=True)
    new_start = new_file["Contents"]["G8"].value
    new_end = new_file["Contents"]["G24"].value

    techno_models = [f for f in os.listdir(TECHNOECONOMIC_MODELS_DIR) if f.endswith(VALID_EXTENSIONS)]
    
    if techno_models:
        existing_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, techno_models[0])
        with open(existing_path, "rb") as f:
            existing_bytes = f.read()
        existing_file = load_workbook(filename=BytesIO(existing_bytes), data_only=True, read_only=True)
        existing_start = existing_file["Contents"]["G8"].value
        existing_end = existing_file["Contents"]["G24"].value

        if existing_start != new_start or existing_end != new_end:
            raise HTTPException(
                status_code=400,
                detail=f"Year range mismatch: existing {existing_start}-{existing_end}, new {new_start}-{new_end}. Upload cancelled."
            )
    else:
        START_YEAR_TECHNO_MODELS = new_start
        END_YEAR_TECHNO_MODELS = new_end
    
    save_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, file.filename)
    with open(save_path, "wb") as f:
        f.write(contents)

    return {"status": "uploaded"}


# Design Capital Structure
@app.get("/design-capital-structure/{model}")
async def get_design_capital_structure(model: str):
    FULL_DESIGN_MODEL_PATH = f"{DESIGN_CAPITAL_STRUCTURE_MODEL}{model}.xlsx"

    if not os.path.exists(GENERAL_INFORMATION_PATH):
        return {"error": "general-information.xlsx not found"}

    if not os.path.exists(DESIGN_CAPITAL_STRUCTURE_TEMPLATE):
        return {"error": "Template file is missing"}

    if not os.path.exists(FULL_DESIGN_MODEL_PATH):
        shutil.copy(DESIGN_CAPITAL_STRUCTURE_TEMPLATE, FULL_DESIGN_MODEL_PATH)

    xls_general = pd.ExcelFile(GENERAL_INFORMATION_PATH, engine="openpyxl")
    general_sheets = xls_general.sheet_names
    expected_sheets = [
        "JUST ACCESS" if name.strip().lower() == "carbon" else name
        for name in general_sheets
    ]

    xls_model = pd.ExcelFile(FULL_DESIGN_MODEL_PATH, engine="openpyxl")
    existing_sheets = xls_model.sheet_names

    xls_template = pd.ExcelFile(DESIGN_CAPITAL_STRUCTURE_TEMPLATE, engine="openpyxl")
    df_electricity = pd.read_excel(xls_template, sheet_name="Electricity", engine="openpyxl", header=None)

    missing_sheets = [s for s in expected_sheets if s not in existing_sheets]

    if missing_sheets:
        sheets_dict = {
            sheet: pd.read_excel(xls_model, sheet_name=sheet, header=None)
            for sheet in existing_sheets
        }
        for missing in missing_sheets:
            sheets_dict[missing] = df_electricity.copy()
        with pd.ExcelWriter(FULL_DESIGN_MODEL_PATH, engine="openpyxl", mode="w") as writer:
            for sheet_name, df in sheets_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(FULL_DESIGN_MODEL_PATH)

    for ws in wb.worksheets:
        year_row = None
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=row, column=col).value
                if isinstance(cell_val, str) and cell_val.strip().lower() == "baseline":
                    year_row = row
                    break
            if year_row:
                break
        
        if year_row is None:
            continue

        cols_to_delete = []
        for col in range(6, ws.max_column + 1):
            cell_val = ws.cell(row=year_row, column=col).value
            try:
                year = int(cell_val)
                if year < START_YEAR_TECHNO_MODELS or year > END_YEAR_TECHNO_MODELS:
                    cols_to_delete.append(col)
            except (TypeError, ValueError):
                continue

        for col_idx in reversed(cols_to_delete):
            ws.delete_cols(col_idx)

    wb.save(FULL_DESIGN_MODEL_PATH)

    return FileResponse(
        FULL_DESIGN_MODEL_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=FULL_DESIGN_MODEL_PATH
    )

@app.post("/save-design-capital-structure")
async def save_design_capital_structure(update: SheetUpdate):
    FULL_DESIGN_MODEL_PATH = f"{DESIGN_CAPITAL_STRUCTURE_MODEL}{update.model}.xlsx"

    if not os.path.exists(FULL_DESIGN_MODEL_PATH):
        return {"error": f"File for model '{update.model}' not found"}

    try:
        df = pd.DataFrame(update.data)
        if 'index' in df.columns:
            df.set_index('index', inplace=True)

        with pd.ExcelWriter(FULL_DESIGN_MODEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=update.sheet_name, index=False)

        return {"message": f"Design capital structure for model '{update.model}' saved successfully."}
    except Exception as e:
        return {"error": f"Failed to save: {str(e)}"}

@app.post("/reset-design-capital-structure")
def reset_design_capital_structure(update: SheetUpdate):
    FULL_DESIGN_MODEL_PATH = f"{DESIGN_CAPITAL_STRUCTURE_MODEL}{update.model}.xlsx"

    if not os.path.exists(FULL_DESIGN_MODEL_PATH):
        return {"error": f"Model file '{FULL_DESIGN_MODEL_PATH}' not found"}
    
    if not os.path.exists(DESIGN_CAPITAL_STRUCTURE_TEMPLATE):
        return {"error": f"Template file '{DESIGN_CAPITAL_STRUCTURE_TEMPLATE}' not found"}

    try:
        if update.sheet_name not in {"JUST ACCESS", "LPG"}:
            template_sheet = "Electricity"
        else:
            template_sheet = update.sheet_name

        template_df = pd.read_excel(DESIGN_CAPITAL_STRUCTURE_TEMPLATE, sheet_name=template_sheet)
        wb = load_workbook(FULL_DESIGN_MODEL_PATH)

        if update.sheet_name in wb.sheetnames:
            std = wb[update.sheet_name]
            wb.remove(std)
            wb.save(FULL_DESIGN_MODEL_PATH)

        with pd.ExcelWriter(FULL_DESIGN_MODEL_PATH, engine='openpyxl', mode='a') as writer:
            template_df.to_excel(writer, sheet_name=update.sheet_name, index=False)

        return {"message": f"'{update.sheet_name}' was reset to template version"}

    except Exception as e:
        return {"error": f"Failed to reset: {str(e)}"}
    

# Techno-Economic Inputs
@app.get("/technoeconomic-inputs/{model}")
async def get_technoeconomic_inputs(model: str):
    FULL_TECHNO_MODEL_PATH = f"{TECHNOECONOMIC_INPUTS_MODEL}{model}.xlsx"
    print(FULL_TECHNO_MODEL_PATH)
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        return {"error": "general-information.xlsx not found"}

    if not os.path.exists(TECHNOECONOMIC_INPUTS_TEMPLATE):
        return {"error": "Template file is missing"}

    if not os.path.exists(FULL_TECHNO_MODEL_PATH):
        shutil.copy(TECHNOECONOMIC_INPUTS_TEMPLATE, FULL_TECHNO_MODEL_PATH)

    xls_general = pd.ExcelFile(GENERAL_INFORMATION_PATH, engine="openpyxl")
    general_sheets = xls_general.sheet_names
    expected_sheets = [
        "JUST ACCESS" if name.strip().lower() == "carbon" else name
        for name in general_sheets
    ]

    xls_model = pd.ExcelFile(FULL_TECHNO_MODEL_PATH, engine="openpyxl")
    existing_sheets = xls_model.sheet_names

    xls_template = pd.ExcelFile(TECHNOECONOMIC_INPUTS_TEMPLATE, engine="openpyxl")
    df_template = pd.read_excel(xls_template, sheet_name="Electricity", engine="openpyxl")

    missing_sheets = [s for s in expected_sheets if s not in existing_sheets]
    if missing_sheets:
        sheets_dict = {
            sheet: pd.read_excel(xls_model, sheet_name=sheet, header=None)
            for sheet in existing_sheets
        }
        for missing in missing_sheets:
            sheets_dict[missing] = df_template.copy()
        with pd.ExcelWriter(FULL_TECHNO_MODEL_PATH, engine="openpyxl", mode="w") as writer:
            for sheet_name, df in sheets_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(FULL_TECHNO_MODEL_PATH)

    for ws in wb.worksheets:
        year_row = None
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=row, column=col).value
                if isinstance(cell_val, str) and cell_val.strip().lower() == "baseline":
                    year_row = row
                    break
            if year_row:
                break

        if year_row is None:
            continue

        cols_to_delete = []
        for col in range(3, ws.max_column + 1):
            cell_val = ws.cell(row=year_row, column=col).value
            try:
                year = int(cell_val)
                if year < START_YEAR_TECHNO_MODELS or year > END_YEAR_TECHNO_MODELS:
                    cols_to_delete.append(col)
            except (TypeError, ValueError):
                continue

        for col_idx in reversed(cols_to_delete):
            ws.delete_cols(col_idx)

    wb.save(FULL_TECHNO_MODEL_PATH)

    return FileResponse(
        FULL_TECHNO_MODEL_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=FULL_TECHNO_MODEL_PATH
    )
