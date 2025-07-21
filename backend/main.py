from fastapi import FastAPI

from routers import fuel_market, techno_economic, design_capital_structure, techno_economic_inputs, carbon_credits, financial_statements, capex_fuel_market

app = FastAPI()

app.include_router(fuel_market.router)
app.include_router(techno_economic.router)
app.include_router(design_capital_structure.router)
app.include_router(techno_economic_inputs.router)
app.include_router(carbon_credits.router)
app.include_router(financial_statements.router)
app.include_router(capex_fuel_market.router)
