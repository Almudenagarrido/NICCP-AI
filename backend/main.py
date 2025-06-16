from fastapi import FastAPI
from fastapi.responses import FileResponse
import os
import shutil
import pandas as pd
from pydantic import BaseModel
from typing import List, Dict


GENERAL_INFORMATION_FOLDER = "./general-information"
GENERAL_INFORMATION_PATH = os.path.join(GENERAL_INFORMATION_FOLDER, "general-information.xlsx")
TEMPLATE_PATH = os.path.join(GENERAL_INFORMATION_FOLDER, "general-information-template.xlsx")
TECHNOECONOMIC_MODELS_DIR = "technoeconomic-models"

class SheetUpdate(BaseModel):
    sheet_name: str
    data: List[Dict]

app = FastAPI()


@app.get("/general-information")
def get_general_information():
    if not os.path.exists(GENERAL_INFORMATION_PATH):
        if os.path.exists(TEMPLATE_PATH):
            shutil.copy(TEMPLATE_PATH, GENERAL_INFORMATION_PATH)
        else:
            return {"error": "The template of the file for the general information was not found"}
    
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


@app.get("/technoeconomic-models")
def list_techno_models():
    files = [f for f in os.listdir(TECHNOECONOMIC_MODELS_DIR)]
    return {"files": files}

@app.get("/technoeconomic-models/{filename}")
def get_techno_models(filename: str):
    model_path = os.path.join(TECHNOECONOMIC_MODELS_DIR, filename)
    if os.path.exists(model_path):
        return FileResponse(model_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return {"error": "Model not found"}
