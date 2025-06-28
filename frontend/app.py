import layout as l
import classes as c


l.header()
l.start_content()
page, subsection, model, fuel_market = l.sidebar()

API_URL = "http://127.0.0.1:8000"

if page == "Fuel Market Information" and fuel_market:
    c.FuelMarketInformation(API_URL, fuel_market)()
    
elif page == "Techno-Economic Models" and subsection:
    c.TechnoEconomicModels(API_URL, subsection, model, fuel_market)()

l.end_content()
l.footer()
