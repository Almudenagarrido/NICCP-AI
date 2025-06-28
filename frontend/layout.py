import streamlit as st
import base64
import classes as c

API_URL = "http://127.0.0.1:8000"

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def start_content(padding_top=30, padding_bottom=10):
    st.markdown(f"""
    <style>
    .app-content {{
        padding-top: {padding_top}px;
        padding-bottom: {padding_bottom}px;
    }}
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="app-content">', unsafe_allow_html=True)

def end_content():
    st.markdown('</div>', unsafe_allow_html=True)

def header():

    niccp_path = "public/niccp-logo.png"
    se4all_path = "public/se4all-logo.png"
    iit_path = "public/iit-logo.png"
    niccp_base64 = get_base64_of_bin_file(niccp_path)
    se4all_base64 = get_base64_of_bin_file(se4all_path)
    iit_base64 = get_base64_of_bin_file(iit_path)

    st.markdown("""
    <style>
    .header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        background-color: #f0f2f6;
        border-bottom: 1px solid #ddd;
        padding: 10px 0;
        text-align: center;
        font-size: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 20px;
        z-index: 1000;
    }
    .header img {
        height: 40px;
        width: auto;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown(f'''
                <div class="header">
                National Integrated Clean Cooking Plan (NICCP)
                <img src="data:image/png;base64,{niccp_base64}" />
                <img src="data:image/png;base64,{se4all_base64}" />
                <img src="data:image/png;base64,{iit_base64}" />
                </div>''', unsafe_allow_html=True)

def sidebar():

    if "page" not in st.session_state:
        st.session_state.page = "Techno-Economic Models"
    if "subsection" not in st.session_state:
        st.session_state.subsection = None
    if "model" not in st.session_state:
        st.session_state.model = None
    if "fuel_market" not in st.session_state:
        st.session_state.fuel_market = None

    fm = c.FuelMarketInformation(API_URL, st.session_state.subsection)
    fuel_market_sheets = fm.get_fuel_markets()
    if 'E-Cooking' in fuel_market_sheets:
        fuel_market_sheets = ['E-Cooking'] + [m for m in fuel_market_sheets if m != 'E-Cooking']

    dcs = c.DesignCapitalStructure(API_URL, st.session_state.subsection, st.session_state.model, st.session_state.fuel_market)
    design_capital_sections = dcs.get_design_capital()
    if 'E-Cooking' in design_capital_sections and 'Electricity' in design_capital_sections:
        design_capital_sections = ['E-Cooking', 'Electricity'] + [s for s in design_capital_sections if s != 'E-Cooking' and s != 'Electricity']

    ti = c.TechnoEconomicInputs(API_URL, st.session_state.subsection, st.session_state.model, st.session_state.fuel_market)
    technoeconomic_input_sheets = ti.get_technoeconomic_inputs()
    if 'E-Cooking' in technoeconomic_input_sheets:
        technoeconomic_input_sheets = ['E-Cooking'] + [t for t in technoeconomic_input_sheets if t != 'E-Cooking']

    # Fuel Market Information
    with st.sidebar:
        with st.expander("Fuel Market Information", expanded=False):
            
            for sheet in fuel_market_sheets:
                if sheet != "Carbon":
                    if st.button(f"{sheet} Financial Inputs"):
                        st.session_state.page = "Fuel Market Information"
                        st.session_state.fuel_market = f"{sheet}"
                        st.session_state.subsection = None
                        st.session_state.model = None

            st.markdown("---")
            if st.button("➕ Add new fuel market"):
                st.session_state.page = "Fuel Market Information"
                st.session_state.fuel_market = "Add"
                st.session_state.subsection = None
                st.session_state.model = None
            
            if st.session_state.page == "Fuel Market Information" and st.session_state.fuel_market is None:
                if fuel_market_sheets:
                    st.session_state.fuel_market = f"{fuel_market_sheets[0]}"
                    st.session_state.subsection = None
                    st.session_state.model = None
            
            market_to_delete = st.selectbox("Delete fuel market", options=fuel_market_sheets)
            if st.button("🗑️"):
                if fm.delete_market(market_to_delete):
                    st.session_state.fuel_market = f"{fuel_market_sheets[0]}"
                    st.session_state.subsection = None
                    st.session_state.model = None
                    st.rerun()

            st.markdown("---")
            if "Carbon" in fuel_market_sheets:
                if st.button(f"Carbon Credits Financial Inputs"):
                        st.session_state.page = "Fuel Market Information"
                        st.session_state.fuel_market = "Carbon"
                        st.session_state.subsection = None
                        st.session_state.model = None

    # Techno-Economic Models
    with st.sidebar:
        if st.session_state.model: 
            st.markdown(f"##### Techno-Economic Model: {st.session_state.model}")
        else:
            st.markdown("##### Techno-Economic Models")

        if st.button("Manage Techno-Economic Models"):
            st.session_state.page = "Techno-Economic Models"
            st.session_state.subsection = "Manage"
            st.session_state.model = None
            st.session_state.fuel_market = None
            st.rerun()

        if st.session_state.model:
            with st.expander("Design Capital Structure", expanded=False):
                for sheet in design_capital_sections:
                    if st.button(f"{sheet} Financial Plan"):
                        st.session_state.page = "Techno-Economic Models"
                        st.session_state.subsection = "Design Capital Structure"
                        st.session_state.fuel_market = f"{sheet}"
                        st.rerun()

                if st.session_state.page == "Techno-Economic Models" and st.session_state.subsection == "Design Capital Structure" and st.session_state.fuel_market == None:
                    if not fuel_market_sheets[0] == "Carbon":
                        st.session_state.fuel_market = fuel_market_sheets[0]
                    else:               
                        st.session_state.fuel_market = "Electricity"

            with st.expander("Techno-Economic Inputs", expanded=False):
                for sheet in technoeconomic_input_sheets:
                    if st.button(f"{sheet} Inputs"):
                        st.session_state.page = "Techno-Economic Models"
                        st.session_state.subsection = "Techno-Economic Inputs"
                        st.session_state.fuel_market = f"{sheet}"
                        st.rerun()

                if st.session_state.page == "Techno-Economic Models" and st.session_state.subsection == "Techno-Economic Inputs" and st.session_state.fuel_market == None:
                    st.session_state.fuel_market = fuel_market_sheets[0]
            
            if st.button("Summary Financing"):
                st.session_state.page = "Techno-Economic Models"
                st.session_state.subsection = "Summary Financing"
                st.rerun()

        if st.session_state.page == "Techno-Economic Models" and st.session_state.subsection == None:
            st.session_state.subsection = "Manage"
            st.session_state.model = None

    return st.session_state.page, st.session_state.subsection, st.session_state.model, st.session_state.fuel_market

def footer():
    st.markdown("""
    <style>
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: #f0f2f6;
        border-top: 1px solid #ddd;
        padding: 10px 0;
        text-align: center;
        font-size: 24px;
    }
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="footer">© 2025 Instituto de Investigación Tecnológica</div>', unsafe_allow_html=True)
