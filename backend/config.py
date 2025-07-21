import os

# Folders
FUEL_MARKET_INFORMATION_FOLDER = "./fuel-market-information"
TECHNOECONOMIC_MODELS_FOLDER = "uploaded-technoeconomic-models"
DESIGN_CAPITAL_STRUCTURE_FOLDER = "./design-capital-structure"
TECHNOECONOMIC_INPUTS_FOLDER = "./technoeconomic-inputs"
CARBON_CREDITS_FOLDER = "./carbon-credits"
FINANCIAL_STATEMENTS_FOLDER = "./financial-statements"
CAPEX_MARKET_FOLDER = "./capex-fuel-markets"

# Templates
FUEL_MARKET_INFORMATION_TEMPLATE = os.path.join(FUEL_MARKET_INFORMATION_FOLDER, "fuel-market-information-template.xlsx")
DESIGN_CAPITAL_STRUCTURE_TEMPLATE = os.path.join(DESIGN_CAPITAL_STRUCTURE_FOLDER, "design-capital-structure-template.xlsx")
TECHNOECONOMIC_INPUTS_TEMPLATE = os.path.join(TECHNOECONOMIC_INPUTS_FOLDER, "technoeconomic-inputs-template.xlsx")
CARBON_CREDITS_TEMPLATE_PATH = os.path.join(CARBON_CREDITS_FOLDER, "carbon-credits-template.xlsx")
FINANCIAL_STATEMENTS_TEMPLATE = os.path.join(FINANCIAL_STATEMENTS_FOLDER, "financial-statements-template.xlsx")
CAPEX_MARKET_TEMPLATE = os.path.join(CAPEX_MARKET_FOLDER, "capex-market-template.xlsx")

# Models
FUEL_MARKET_INFORMATION_PATH = os.path.join(FUEL_MARKET_INFORMATION_FOLDER, "fuel-market-information.xlsx")
DESIGN_CAPITAL_STRUCTURE_MODEL = os.path.join(DESIGN_CAPITAL_STRUCTURE_FOLDER, "design-capital-structure-")
TECHNOECONOMIC_INPUTS_MODEL = os.path.join(TECHNOECONOMIC_INPUTS_FOLDER, "technoeconomic-inputs-")
CARBON_CREDITS_PATH = os.path.join(CARBON_CREDITS_FOLDER, "carbon-credits.xlsx")
FINANCIAL_STATEMENTS_MODEL = os.path.join(FINANCIAL_STATEMENTS_FOLDER, "financial-statements-")
CAPEX_MARKET_MODEL = os.path.join(CAPEX_MARKET_FOLDER, "capex-market-")

# Others
FORMULAS_JSON_PATH = "./formulas_map.json"
FOLDERS = [
    TECHNOECONOMIC_MODELS_FOLDER,
    TECHNOECONOMIC_INPUTS_FOLDER,
    DESIGN_CAPITAL_STRUCTURE_FOLDER,
    CARBON_CREDITS_FOLDER,
    FUEL_MARKET_INFORMATION_FOLDER,
    FINANCIAL_STATEMENTS_FOLDER,
    CAPEX_MARKET_FOLDER,
]

VALID_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm")
START_YEAR_TECHNO_MODELS = None
END_YEAR_TECHNO_MODELS = None
