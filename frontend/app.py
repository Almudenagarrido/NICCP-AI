import layout as l
import classes as c
import streamlit as st
import base64

def image_to_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

if "page" not in st.session_state:
    st.session_state.page = "Welcome"

if st.session_state.page == "Welcome":
    l.start_content()

    st.markdown("""
        <style>
        body {
            background-color: #e0f2f1;
        }

        .welcome-title {
            text-align: center;
            font-size: 2.5rem;
            font-weight: 600;
            color: #004d40;
        }

        .center-button {
            display: flex;
            justify-content: center;
        }

        </style>

        <div class="welcome-title">Welcome to the Financial Clean Cooking Platform</div>
    """, unsafe_allow_html=True)

    img_base64 = image_to_base64("imagen.png")
    st.markdown(f"""
        <div style="text-align: center;">
            <img src="data:image/png;base64,{img_base64}" style="height: 400px;" />
        </div>
    """, unsafe_allow_html=True)


    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("Select Scenario"):
        st.session_state.page = "Scenario Selection"
        st.rerun()

elif st.session_state.page == "Scenario Selection":
    l.start_content()
    st.markdown("### Choose a country scenario to begin modeling:")

    if st.button("Rwanda"):
        st.session_state.country = "Rwanda"

    if "country" in st.session_state:
        st.image(
            "imagen2.png",
            caption="Rwanda – Selected Scenario",
            width=200,
        )

        if st.button("Start Modeling"):
            st.session_state.page = "Techno-Economic Inputs"
            st.session_state.subsection = "Manage Techno-Economic Inputs"
            st.session_state.model = None
            st.session_state.fuel_market = None
            st.rerun()

    l.end_content()

else:
    l.start_content()
    page, subsection, model, fuel_market = l.sidebar()

    API_URL = "http://127.0.0.1:8000"

    if page == "Financial Inputs" and fuel_market:
        c.FuelMarketInformation(API_URL, fuel_market)()
    elif page == "Techno-Economic Inputs" and subsection:
        c.TechnoEconomicModels(API_URL, subsection, model, fuel_market)()
    l.end_content()
