import streamlit as st
import base64
import classes as c

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
    if "design_market" not in st.session_state:
        st.session_state.design_market = None

    gi = c.GeneralInformation("http://127.0.0.1:8000", st.session_state.get("subsection"))
    market_sheets = gi.get_financial_markets()

    # General Information
    with st.sidebar:
        with st.expander("General Information", expanded=False):
            
            for sheet in market_sheets:
                if st.button(f"{sheet} Financial Inputs"):
                    st.session_state.page = "General Information"
                    st.session_state.design_market = f"{sheet}"
                    st.session_state.subsection = None
                    st.session_state.model = None

            st.markdown("---")
            if st.button("➕ Add New Market"):
                st.session_state.page = "General Information"
                st.session_state.design_market = "Add"
                st.session_state.subsection = None
                st.session_state.model = None
            
            if st.session_state.page == "General Information" and st.session_state.design_market is None:
                if market_sheets:
                    st.session_state.design_market = f"{market_sheets[0]}"
                    st.session_state.subsection = None
                    st.session_state.model = None
            
            market_to_delete = st.selectbox("Delete Market", options=market_sheets)
            if st.button("🗑️"):
                if gi.delete_market(market_to_delete):
                    st.session_state.design_market = f"{market_sheets[0]}"
                    st.session_state.subsection = None
                    st.session_state.model = None
                    st.rerun()

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
            st.rerun()

        if st.session_state.model:
            with st.expander("Design Capital Structure", expanded=False):
                for sheet in market_sheets:
                    if sheet == "Carbon":
                        sheet = "JUST ACCESS"
                    if st.button(f"{sheet} Financial Plan"):
                        st.session_state.page = "Techno-Economic Models"
                        st.session_state.subsection = "Design Capital Structure"
                        st.session_state.design_market = f"{sheet}"
                        st.rerun()

                if st.session_state.page == "Techno-Economic Models" and st.session_state.subsection == "Design Capital Structure" and st.session_state.design_market == None:
                    if not market_sheets[0] == "Carbon":
                        st.session_state.design_market = market_sheets[0]
                    else:               
                        st.session_state.design_market = "JUST ACCESS"

            with st.expander("Techno-Economic Inputs", expanded=False):
                for sheet in market_sheets:
                    if sheet == "Carbon":
                        sheet = "C02"
                    if st.button(f"{sheet} Inputs"):
                        st.session_state.page = "Techno-Economic Models"
                        st.session_state.subsection = "Techno-Economic Inputs"
                        st.session_state.design_market = f"{sheet}"
                        st.rerun()

                if st.session_state.page == "Techno-Economic Models" and st.session_state.subsection == "Techno-Economic Inputs" and st.session_state.design_market == None:
                    if not market_sheets[0] == "Carbon":
                        st.session_state.design_market = market_sheets[0]
                    else:               
                        st.session_state.design_market = "C02"
            
            if st.button("Summary Financing"):
                st.session_state.page = "Techno-Economic Models"
                st.session_state.subsection = "Summary Financing"
                st.rerun()

        if st.session_state.page == "Techno-Economic Models" and st.session_state.subsection == None:
            st.session_state.subsection = "Manage"
            st.session_state.model = None

    return st.session_state.page, st.session_state.subsection, st.session_state.model, st.session_state.design_market

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
