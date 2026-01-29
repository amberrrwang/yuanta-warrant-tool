import streamlit as st
import pandas as pd
from io import BytesIO
from yuanta import launch_driver, scrape_one_wid, HEADER_ORDER, BASIC_LABELS # 假設你把原程式存成 yuanta.py
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import openpyxl, os, re, time
import requests 
# ======= Streamlit 介面 =======
st.set_page_config(page_title="元大權證抓取工具")

st.title("📈 權證資料即時抓取")

# 介面設定
target_wids = st.text_area("請輸入權證代碼 (用逗號或換行隔開)", value="00637L, 03111U")
process_btn = st.button("開始抓取並產製 Excel")

if process_btn:
    wid_list = [w.strip() for w in target_wids.replace('\n', ',').split(',') if w.strip()]
    
    with st.spinner('正在啟動瀏覽器並抓取資料...'):
        # 呼叫你原本的 Selenium 邏輯 (記得 headless 要設為 True)
        driver = launch_driver(headless=True)
        rows = []
        progress_bar = st.progress(0)
        
        for idx, wid in enumerate(wid_list):
            row = scrape_one_wid(driver, wid)
            rows.append(row)
            progress_bar.progress((idx + 1) / len(wid_list))
        
        driver.quit()

    if rows:
        # 顯示預覽表格
        df = pd.DataFrame(rows)[HEADER_ORDER]
        st.write("### 資料預覽", df)

        # 產製 Excel 並提供下載 (這部分改寫你原本的 save_rows_to_excel)
        output = BytesIO()
        # ... 這裡放入你原本用 openpyxl 寫入 output 的邏輯 ...
        
        st.download_button(
            label="📥 下載 Excel 報表",
            data=output.getvalue(),
            file_name=f"yuanta_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )