import layout as l
import classes as c


l.header()
l.start_content()
page, subsection, model, design_market = l.sidebar()

API_URL = "http://127.0.0.1:8000"

if page == "General Information" and design_market:
    c.GeneralInformation(API_URL, design_market)()
    
elif page == "Techno-Economic Models" and subsection:
    c.TechnoEconomicModels(API_URL, subsection, model, design_market)()

l.end_content()
l.footer()
