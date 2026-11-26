import streamlit as st

st.set_page_config(page_title="工具已移至私人入口", layout="centered")
st.title("工具已移至私人入口")
st.info("此公開網址已停止提供工具功能。請登入私人入口後繼續使用。")
st.link_button("前往私人入口", 'https://inventory-app-numbertalk-bda7ktbu7ny5ocblm7pdon.streamlit.app/legacy', use_container_width=True)
