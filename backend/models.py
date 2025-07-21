from pydantic import BaseModel
from typing import List, Dict

class SheetUpdate(BaseModel):
    model:str
    sheet_name: str
    data: List[Dict]

class MarketName(BaseModel):
    name: str
