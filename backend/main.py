from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
import os
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


class SheetUpdate(BaseModel):
    sheet_name: str
    data: List[Dict]

class MarketName(BaseModel):
    name: str

app = FastAPI()


# General Information
@app.get("/general-information")
def get_general_information():
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        if os.path.exists(GENERAL_INFORMATION_TEMPLATE):
            shutil.copy(GENERAL_INFORMATION_TEMPLATE, GENERAL_INFORMATION_PATH)
        else:
            return {"error": "Template file is missing"}
    
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

        with pd.ExcelWriter(GENERAL_INFORMATION_PATH, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            template_df.to_excel(writer, sheet_name=update.sheet_name, index=False)

        return {"message": f"'{update.sheet_name}' was reset to template version"}

    except Exception as e:
        return {"error": f"Failed to reset: {str(e)}"}

@app.post("/add-market")
def add_market(market: MarketName):
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        return {"error": "general-information.xlsx not found"}

    xls = pd.ExcelFile(GENERAL_INFORMATION_PATH, engine="openpyxl")

    if "Electricity" not in xls.sheet_names:
        return {"error": "Base sheet 'Electricity' not found"}

    df_electricity = pd.read_excel(xls, sheet_name="Electricity", engine="openpyxl")
    new_sheet_name = market.name

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
    valid_extensions = (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")
    files = [f for f in os.listdir(TECHNOECONOMIC_MODELS_DIR) if f.endswith(valid_extensions)]
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
    valid_extensions = (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")
    if not file.filename.endswith(valid_extensions):
        raise HTTPException(status_code=400, detail="Only Excel files are allowed.")

    save_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, file.filename)
    with open(save_path, "wb") as f:
        f.write(file.file.read())

    return {"status": "uploaded"}


# Design Capital Structure
@app.get("/design-capital-structure/{model}")
async def get_design_capital_structure(model: str):
    FULL_DESIGN_MODEL_PATH = f"{DESIGN_CAPITAL_STRUCTURE_MODEL}{model}.xlsx"
    
    if not os.path.exists(FULL_DESIGN_MODEL_PATH):
        if os.path.exists(DESIGN_CAPITAL_STRUCTURE_TEMPLATE):
            shutil.copy(DESIGN_CAPITAL_STRUCTURE_TEMPLATE, FULL_DESIGN_MODEL_PATH)
        else:
            return {"error": "Template file is missing"}
    
    return FileResponse(
        FULL_DESIGN_MODEL_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="general-information.xlsx"
    )

