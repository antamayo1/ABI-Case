import streamlit as st
import pandas as pd

st.set_page_config(page_title="Design Transformer", layout="wide")

if not st.user.is_logged_in:
  _, col1, _ = st.columns([5, 3, 5])
  with col1:
    with st.container(border=True):
      st.image("LMC_Logo.jpeg", use_container_width=True)
      st.markdown("## title")
      login_btn = st.button(
        "**Log in** with **Google**",
        use_container_width=True,
        type="primary"
      )
      if login_btn:
        st.login()
else:
  user_details = st.user.to_dict()
  with st.container(border=True):
    col1, col2 = st.columns([7, 1])
    with col1:
      st.markdown("#### Schedule Generator")
    with col2:
      logout_btn = st.button(
        "Log out",
        type="primary",
        use_container_width=True
      )
      if logout_btn:
        st.logout()
  with st.expander("Upload Schedule", expanded=True):
    st.session_state.input_file = st.file_uploader("Please input the raw export file", type=["xlsx", "xls"])
  if st.session_state.input_file:
    df = pd.read_excel(st.session_state.input_file)
    st.session_state.df = df
    st.dataframe(df)
  else:
    st.info("Please upload a schedule file to proceed.")