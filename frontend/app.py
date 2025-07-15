import layout as l
import classes as c
import streamlit as st

if "page" not in st.session_state:
    st.session_state.page = "Welcome"

if st.session_state.page == "Welcome":
    l.start_content()
    st.markdown("<h1 style='text-align: center;'>Welcome to the Financial Clean Cooking Platform</h1>", unsafe_allow_html=True)
    if st.button("Select Scenario"):
        st.session_state.page = "Scenario Selection"
        st.rerun()
    l.end_content()

elif st.session_state.page == "Scenario Selection":
    l.start_content()
    st.markdown("### Choose a country scenario to begin modeling:")
    if st.button("Rwanda"):
        st.session_state.country = "Rwanda"
    if "country" in st.session_state:
        if st.button("Start Modeling"):
            st.session_state.page = "Techno-Economic Models"
            st.session_state.subsection = "Manage Techno-Economic Models"
            st.session_state.model = None
            st.session_state.fuel_market = None
            st.rerun()
    l.end_content()

else:
    l.start_content()
    page, subsection, model, fuel_market = l.sidebar()

    API_URL = "http://127.0.0.1:8000"

    if page == "Fuel Market Information" and fuel_market:
        c.FuelMarketInformation(API_URL, fuel_market)()
    elif page == "Techno-Economic Models" and subsection:
        c.TechnoEconomicModels(API_URL, subsection, model, fuel_market)()
    l.end_content()
