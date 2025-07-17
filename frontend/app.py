import layout as l
import classes as c
import streamlit as st

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

        .banner-image {
            width: 100%;
            height: 350px;
            background-image: url('https://thumbs.dreamstime.com/b/mujer-africana-cocina-comida-tradicional-en-la-calle-mujeres-africanas-cocinan-231196870.jpg');
            background-size: cover;
            background-position: center;
            border-radius: 10px;
        }

        .center-button {
            display: flex;
            justify-content: center;
        }

        </style>

        <div class="welcome-title">Welcome to the Financial Clean Cooking Platform</div>
        <div class="banner-image"></div>
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
            "https://tse2.mm.bing.net/th/id/OIP.1O_v7LZzPuKZGK-BtyiRFQHaGW?w=852&h=730&rs=1&pid=ImgDetMain&o=7&rm=3",
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
