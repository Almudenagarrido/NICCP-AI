import streamlit as st
import base64

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

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

def sidebar():
    if "page" not in st.session_state:
        st.session_state.page = None
    if "subsection" not in st.session_state:
        st.session_state.subsection = None

    with st.sidebar:
        with st.expander("General Information", expanded=True):
            
            if st.button("Financial Electricity Inputs"):
                st.session_state.subsection = "Financial Electricity Inputs"
            if st.button("Financial LPG Inputs"):
                st.session_state.subsection = "Financial LPG Inputs"
            if st.button("CO2 - Carbon Finance Inputs"):
                st.session_state.subsection = "CO2 - Carbon Finance Inputs"
            
            st.session_state.page = "General Information"

            if st.session_state.page == "General Information" and st.session_state.subsection is None:
                st.session_state.subsection = "Financial Electricity Inputs"

        if st.button("Manage Techno-Economic Models"):
            st.session_state.page = "Manage Techno-Economic Models"
            st.session_state.subsection = None

    return st.session_state.page, st.session_state.subsection

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
