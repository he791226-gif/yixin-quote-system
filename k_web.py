import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="翌新空壓機報價系統", layout="wide")

# CSS 修正：保留隱藏功能，但修復手機版選單按鈕不見的問題
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .price-text { color: #E84118; font-size: 24px; font-weight: bold; }
    [data-testid="stVerticalBlock"] > div:has(div.stImage) {
        background-color: white; padding: 20px; border-radius: 15px;
        border: 1px solid #ddd; min-height: 450px; text-align: center;
    }
    
    /* 隱藏頂部裝飾，但保留手機版漢堡選單按鈕 */
    header { visibility: hidden; }
    .stAppHeader { visibility: hidden; }
    [data-testid="stHeader"] { visibility: visible !important; background: transparent; } 
    [data-testid="stToolbar"] { visibility: hidden; }
    
    /* 修復隱藏選單後手機端無法點擊的問題 */
    button[kind="header"] { visibility: visible !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 產品資料庫與規格說明 (包含圖1要求之文字) ---
if 'cart' not in st.session_state:
    st.session_state.cart = {} 

# 產品規格文字庫 (根據圖1及產品特性)
product_specs = {
    "10馬永磁變頻空壓機HCV-10PM-A": "HCV高端系列\n型號:HCV-10PM-A\n馬力:1馬至10馬<隨壓力調整馬力>\n噪音:68±5\nIE4永磁高效率馬達<馬達無軸承><無皮帶>\n5kg~8kg可選變頻壓力\nLCD液晶顯示預警/警告/錯誤跳機保護\n外型尺寸:745*680*910(mm)\n變頻器與電機一體化非外掛變頻空壓機\n啟動電流衝擊小,恆壓恆溫控制,故障率低",
    "5馬空壓機": "標準型高效能空壓機\n穩定風壓輸出\n適合小型加工廠與維修站",
    # 可依此類推增加乾燥機、儲氣筒之說明...
}

products = {
    "空壓機": {
        "HCV高端系列": [
            ("10馬永磁變頻空壓機HCV-10PM-A", "air_10.png"), ("30馬永磁變頻空壓機HCV-30PM-A", "air_30.png"), 
            ("75馬永磁變頻空壓機HCV-75PM-A", "air_75.png"), ("20馬永磁變頻空壓機HCV-20PM-A", "air_20.png"),
            ("50馬永磁變頻空壓機HCV-50PM-A", "air_50.png"), ("100馬永磁變頻空壓機HCV-100PM-A", "air_100.png")
        ],
        "標準型": [
            ("5馬空壓機", "air_5.png"), ("air_mu.png", "air_mu.png")
        ]
    },
    "儲氣筒": [
        ("105儲氣筒", "tank_105.png"), ("360儲氣筒", "tank_360.png"), ("660儲氣筒", "tank_660.png")
    ],
    "乾燥機": {
        "宙升": [
            ("5馬宙升乾燥機SD-005", "zs_dryer_5.png"), ("10馬宙升乾燥機SD-010", "zs_dryer_10.png"),
            ("20馬宙升乾燥機SD-020", "zs_dryer_20.png"), ("30馬宙升乾燥機SD-030", "zs_dryer_30.png"), 
            ("50馬宙升乾燥機SD-050", "zs_dryer_50.png"), ("100馬宙升乾燥機", "zs_dryer_100.png")
        ],
        "艾冷": [
            ("5馬艾冷乾燥機", "al_dryer_5.png"), ("10馬艾冷乾燥機", "al_dryer_10.png"), ("20馬艾冷乾燥機", "al_dryer_20.png"), 
            ("30馬艾冷乾燥機", "al_dryer_30.png"), ("50馬艾冷乾燥機", "al_dryer_50.png"), ("100馬艾冷乾燥機", "al_dryer_10.png")
        ]
    },
    "選配配件": [
        ("Ckd自動排水器", "drainer_ckd.png"), ("電子式自動排水器", "drainer_e.png"), 
        ("前置旋風分離器", "separator.png"), ("超精密過濾器組", "filter.png"), ("超精密過濾器芯", "filter_core.png")
    ]
}

unit_map = {
    "105儲氣筒": "只", "360儲氣筒": "只", "660儲氣筒": "只",
    "Ckd自動排水器": "只", "電子式自動排水器": "只", "前置旋風分離器": "只",
    "超精密過濾器組": "只", "超精密過濾器芯": "只"
}

if 'price_config' not in st.session_state:
    st.session_state.price_config = {}
    # 初始化價格 (包含嵌套字典處理)
    def init_prices(d):
        for k, v in d.items():
            if isinstance(v, dict): init_prices(v)
            elif isinstance(v, list):
                for name, _ in v: st.session_state.price_config[name] = 0
    init_prices(products)

# --- 3. 側邊欄 ---
st.sidebar.title("🏢 翌新後台管理")
customer_name = st.sidebar.text_input("客戶名稱", value="")
contact_person = st.sidebar.text_input("聯絡人", value="")

# 功能 3: 增加電壓選擇
voltage = st.sidebar.radio("電力規格", ["220V", "380V"], horizontal=True)

with st.sidebar.expander("⚙️ 價格調整"):
    for name in sorted(st.session_state.price_config.keys()):
        st.session_state.price_config[name] = st.sidebar.number_input(f"{name} 單價", value=st.session_state.price_config[name])

# --- 4. 主展示介面 ---
st.title("請選擇設備類別")
tabs = st.tabs(["空壓機", "儲氣筒", "乾燥機", "選配配件"])

def display_items(item_list):
    cols = st.columns(3)
    for i, (name, img) in enumerate(item_list):
        with cols[i % 3]:
            if os.path.exists(img):
                st.image(img, width=250)
            st.write(f"**{name}**")
            if st.button(f"➕ 加入報價單", key=f"btn_{name}"):
                st.session_state.cart[name] = st.session_state.cart.get(name, 0) + 1
                st.toast(f"已加入: {name}")

with tabs[0]:
    # 功能 1: 空壓機增加子選項
    air_type = st.radio("選擇類型", ["HCV高端系列", "標準型"], horizontal=True)
    display_items(products["空壓機"][air_type])
with tabs[1]: display_items(products["儲氣筒"])
with tabs[2]:
    brand = st.radio("選擇品牌", ["宙升", "艾冷"], horizontal=True)
    display_items(products["乾燥機"][brand])
with tabs[3]: display_items(products["選配配件"])

# --- 5. 報價清單與 EXCEL 輸出 ---
st.divider()
if st.session_state.cart:
    st.subheader("📋 目前報價清單")
    table_data = []
    total_val = 0
    for name, qty in st.session_state.cart.items():
        p = st.session_state.price_config.get(name, 0)
        sub = p * qty
        total_val += sub
        table_data.append([name, unit_map.get(name, "台"), qty, f"${p:,}", f"${sub:,}"])
    
    st.table(pd.DataFrame(table_data, columns=["品名及規格", "單位", "數量", "單價", "金額"]))
    st.write(f"**目前選擇電力：{voltage}**")
    st.markdown(f"### <span class='price-text'>總計金額：${total_val:,}</span>", unsafe_allow_html=True)

    template_path = "翌新估價單EXCELNEW.xlsx"
    if os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        bold_font = Font(name='新細明體', size=11, bold=True)
        spec_font = Font(name='新細明體', size=10)
        right_align = Alignment(horizontal='right', vertical='center')
        center_align = Alignment(horizontal='center', vertical='center')
        top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # 1. 填入客戶資訊與日期
        ws['B11'] = customer_name
        ws['B12'] = contact_person
        ws['H14'] = f"估價日期：{datetime.now().strftime('%Y-%m-%d')}"
        
        # 2. 填入明細與功能 4 的規格文字
        current_row = 17
        for i, (name, qty) in enumerate(st.session_state.cart.items()):
            price = st.session_state.price_config.get(name, 0)
            
            # NO, 品名, 單位, 數量, 單價, 金額
            ws.cell(row=current_row, column=1, value=i+1).font = bold_font
            ws.cell(row=current_row, column=2, value=f"{name} ({voltage})").font = bold_font
            ws.cell(row=current_row, column=6, value=unit_map.get(name, "台")).alignment = center_align
            ws.cell(row=current_row, column=7, value=qty).alignment = center_align
            ws.cell(row=current_row, column=8, value=price).alignment = right_align
            ws.cell(row=current_row, column=9, value=price * qty).alignment = right_align
            
            # 功能 4: 在品名下方插入規格文字說明
            if name in product_specs:
                current_row += 1
                spec_cell = ws.cell(row=current_row, column=2, value=product_specs[name])
                spec_cell.font = spec_font
                spec_cell.alignment = top_left_align
                # 調整列高以容納文字
                ws.row_dimensions[current_row].height = 150 
            
            current_row += 1

        # 3. 合計金額 (對準 I36 或動態計算)
        ws['I36'] = total_val
        ws['H36'] = "合計："

        output = io.BytesIO()
        wb.save(output)
        st.download_button(label="📤 下載 翌新專業報價單 (Excel)", data=output.getvalue(), file_name=f"翌新報價_{customer_name}.xlsx")

    if st.button("🗑️ 清空重選"):
        st.session_state.cart = {}
        st.rerun()
