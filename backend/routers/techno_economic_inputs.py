from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
import pandas as pd
import shutil
import os

from models import SheetUpdate
import config as c

router = APIRouter()

@router.get("/technoeconomic-inputs/{model}")
async def get_technoeconomic_inputs(model: str):
    if model.lower() in ["none", "bau"]:
        if not os.path.exists(c.TECHNOECONOMIC_INPUTS_TEMPLATE):
            return {"error": "Template file is missing"}

        return FileResponse(
            c.TECHNOECONOMIC_INPUTS_TEMPLATE,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=c.TECHNOECONOMIC_INPUTS_TEMPLATE
        )

    FULL_TECHNO_MODEL_PATH = f"{c.TECHNOECONOMIC_INPUTS_MODEL}{model}.xlsx"

    if not os.path.exists(c.TECHNOECONOMIC_INPUTS_TEMPLATE):
        return {"error": "Template file is missing"}

    if not os.path.exists(FULL_TECHNO_MODEL_PATH):
        shutil.copy(c.TECHNOECONOMIC_INPUTS_TEMPLATE, FULL_TECHNO_MODEL_PATH)

    return FileResponse(
        FULL_TECHNO_MODEL_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=FULL_TECHNO_MODEL_PATH
    )

@router.post("/save-technoeconomic-inputs")
async def save_technoeconomic_inputs(update: SheetUpdate):

    if not update.model or not update.sheet_name or update.data is None:
        return {"error": "Missing model, sheet_name or data"}

    FULL_TECHNO_MODEL_PATH = f"{c.TECHNOECONOMIC_INPUTS_MODEL}{update.model}.xlsx"
    df_new = pd.DataFrame(update.data)

    with pd.ExcelWriter(FULL_TECHNO_MODEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_new.to_excel(writer, sheet_name=update.sheet_name, index=False)

    return {"message": f"Sheet '{update.sheet_name}' saved successfully for model '{update.model}'"}

@router.post("/reset-technoeconomic-inputs")
async def reset_technoeconomic_inputs(update: SheetUpdate):
    model = update.model
    sheet_name = update.sheet_name

    if not model or not sheet_name:
        raise HTTPException(status_code=400, detail="Missing model or sheet_name")

    FULL_TECHNO_MODEL_PATH = f"{c.TECHNOECONOMIC_INPUTS_MODEL}{model}.xlsx"

    if not os.path.exists(FULL_TECHNO_MODEL_PATH):
        if not os.path.exists(c.TECHNOECONOMIC_INPUTS_TEMPLATE):
            raise HTTPException(status_code=500, detail="Template file is missing")
        shutil.copy(c.TECHNOECONOMIC_INPUTS_TEMPLATE, FULL_TECHNO_MODEL_PATH)

    try:
        with pd.ExcelFile(c.TECHNOECONOMIC_INPUTS_TEMPLATE, engine='openpyxl') as xls:
            if sheet_name not in xls.sheet_names:
                raise HTTPException(status_code=400, detail=f"Sheet '{sheet_name}' not found in template")
            df_template = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl')

        with pd.ExcelWriter(FULL_TECHNO_MODEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_template.to_excel(writer, sheet_name=sheet_name, index=False)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error resetting sheet: {e}")

    return {"message": f"Sheet '{sheet_name}' reset to template version for model '{model}'"}
