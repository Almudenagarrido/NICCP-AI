import layout as l
import classes as c


l.header()
l.start_content()
page, subsection = l.sidebar()

API_URL = "http://127.0.0.1:8000"

if page == "General Information" and subsection:
    c.GeneralInformation(subsection)()
    
if page == "Manage Techno-Economic Models":
    techno_models = c.TechnoEconomicModels()
    techno_models.manage_technoeconomic_models()

l.end_content()
l.footer()
