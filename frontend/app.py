import layout as l
import classes as c


l.header()
l.start_content()
page, subsection = l.sidebar()

API_URL = "http://127.0.0.1:8000"

if page == "General Information" and subsection:
    c.GeneralInformation(subsection)()
    
elif page == "Techno-Economic Models" and subsection:
    c.TechnoEconomicModels(subsection)()

l.end_content()
l.footer()
