import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment

# --- 1. 頁面設定 ---
st.set_page_config(page_title="翌新空壓機報價系統", layout="wide")

# CSS 隱藏橘色裝飾圖標
st.markdown("""
    <style>
    header, .stAppHeader, #MainMenu, footer, [data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
    }
    [data-testid="stSidebarCollapsedControl"] { display: none !important; }
    .price-text { color: #E84118; font-size: 24px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 產品規格 (標楷體 10馬規格) ---
product_specs = {
    "10馬永磁變頻空壓機HCV-10PM-A": (
        "型號:HCV-10PM-A | 噪音:68±5\n"
        "電壓:三相220V | IE4永磁高效率馬達\n"
        "5kg~8kg變頻壓力 | 恆壓恆溫控制\n"
        "外型尺寸:745*680*910(mm)"
    )
}

# --- 3. 產品資料庫 ---
products = {
    "空壓機": {
        "HCV高端系列": [
            ("10馬永磁變頻空壓機HCV-10PM-A", "air_10.png"), ("20馬永磁變頻空壓機HCV-20PM-A", "air_20.png"), 
            ("30馬永磁變頻空壓機HCV-30PM-A", "air_30.png"), ("50馬永磁變頻空壓機HCV-50PM-A", "air_50.png"),
            ("75馬永磁變頻空壓機HCV-75PM-A", "air_75.png"), ("100馬永磁變頻空壓機HCV-100PM-A", "air_100.png")
        ],
        "標準型": [("5馬空壓機", "air_5.png")]
    },
    "儲氣筒": [("105儲氣筒", "tank_105.png"), ("360儲氣筒", "tank_360.png"), ("660儲氣筒", "tank_660.png")],
    "乾燥機": {
        "宙升": [("5馬宙升乾燥機SD-005", "zs_dryer_5.png"), ("10馬宙升乾燥機SD-010", "zs_dryer_10.png")],
        "艾冷": [("5馬艾冷乾燥機", "al_dryer_5.png"), ("10馬艾冷乾燥機", "al_dryer_10.png")]
    },
    "超精密過濾器組": {
        "合成牌": [("超精密過濾器組(合成牌)", "filter_com.png")],
        "PARK": [("超精密過濾器組(PARK)", "filter_park.png")]
    },
    "選配配件": [("Ckd自動排水器", "drainer_ckd.png"), ("電子式自動排水器", "drainer_e.png")]
}

unit_map = {"105儲氣筒": "只", "360儲氣筒": "只", "Ckd自動排水器": "只"}

if 'cart' not in st.session_state: st.session_state.cart = {}
if 'price_config' not in st.session_state:
    st.session_state.price_config = {n: 0 for series in products.values() for n, _ in (series if isinstance(series, list) else [i for sub in series.values() for i in sub])}

# --- 4. 展示與選擇 ---
st.title("請選擇設備類別")
tabs = st.tabs(["空壓機", "儲氣筒", "乾燥機", "超精密過濾器組", "選配配件"])

def display_items(item_list):
    cols = st.columns(3)
    for i, (name, img) in enumerate(item_list):
        with cols[i % 3]:
            if os.path.exists(img): st.image(img, width=220)
            st.write(f"**{name}**")
            if st.button(f"➕ 加入", key=f"btn_{name}"):
                st.session_state.cart[name] = st.session_state.cart.get(name, 0) + 1
                st.rerun()

with tabs[0]: display_items(products["空壓機"][st.radio("系列", ["HCV高端系列", "標準型"], horizontal=True)])
with tabs[1]: display_items(products["儲氣筒"])
with tabs[2]: display_items(products["乾燥機"][st.radio("品牌", ["宙升", "艾冷"], horizontal=True)])
with tabs[3]: display_items(products["超精密過濾器組"][st.radio("品牌 ", ["合成牌", "PARK"], horizontal=True)])
with tabs[4]: display_items(products["選配配件"])

# --- 5. 報價清單 ---
st.divider()
if st.session_state.cart:
    col_list, col_admin = st.columns([2, 1])
    with col_admin:
        with st.popover("⚙️ 修改客戶/價格", use_container_width=True):
            customer_name = st.text_input("客戶名稱")
            contact_person = st.text_input("聯絡人")
            voltage = st.radio("電力規格", ["220V", "380V"], horizontal=True)
            for name in sorted(st.session_state.cart.keys()):
                st.session_state.price_config[name] = st.number_input(f"{name} 單價", value=st.session_state.price_config.get(name, 0), step=100)

    with col_list:
        st.subheader("📋 目前報價清單")
        table_data = []
        total_val = 0
        for name, qty in st.session_state.cart.items():
            p = st.session_state.price_config.get(name, 0)
            sub = p * qty
            total_val += sub
            table_data.append([name, unit_map.get(name, "台"), qty, f"${p:,}", f"${sub:,}"])
        st.table(pd.DataFrame(table_data, columns=["品名", "單位", "數量", "單價", "金額"]))

    # --- 6. Excel 匯出 (強化字體設定) ---
    template_path = "翌新估價單EXCELNEW.xlsx"
    if os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        
        # 定義字體樣式：標楷體 + 粗體
        kai_bold_font = Font(name='標楷體', size=12, bold=True)
        kai_spec_font = Font(name='標楷體', size=10, bold=True)

        # 設定 B11, B12, H14
        ws['B11'].font = kai_bold_font
        ws['B11'] = customer_name
        
        ws['B12'].font = kai_bold_font
        ws['B12'] = contact_person
        
        ws['H14'].font = kai_bold_font
        ws['H14'] = f"日期：{datetime.now().strftime('%Y-%m-%d')}"
        
        current_row = 17
        for i, (name, qty) in enumerate(st.session_state.cart.items()):
            p = st.session_state.price_config.get(name, 0)
            
            # 品名與規格行
            ws.cell(row=current_row, column=1, value=i+1).font = kai_bold_font
            name_cell = ws.cell(row=current_row, column=2, value=f"{name} ({voltage})")
            name_cell.font = kai_bold_font
            
            ws.cell(row=current_row, column=7, value=qty).font = kai_bold_font
            ws.cell(row=current_row, column=8, value=p).font = kai_bold_font
            ws.cell(row=current_row, column=9, value=p * qty).font = kai_bold_font
            
            # 詳細規格填寫 (解決擠壓問題)
            if name in product_specs:
                current_row += 1
                spec_cell = ws.cell(row=current_row, column=2, value=product_specs[name])
                spec_cell.font = kai_spec_font
                spec_cell.alignment = Alignment(wrap_text=True, vertical='center')
                ws.row_dimensions[current_row].height = 65 # 適中的行高，確保排版工整
            
            current_row += 1

        ws['I36'] = total_val
        output = io.BytesIO()
        wb.save(output)
        
        st.download_button("📤 下載標楷體專業報價單", data=output.getvalue(), file_name=f"報價_{customer_name}.xlsx", use_container_width=True)
