import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="翌新空壓機報價系統", layout="wide")

# CSS 修正：徹底隱藏 GitHub，並強制撈回手機版左側控制按鈕
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .price-text { color: #E84118; font-size: 24px; font-weight: bold; }
    [data-testid="stVerticalBlock"] > div:has(div.stImage) {
        background-color: white; padding: 20px; border-radius: 15px;
        border: 1px solid #ddd; min-height: 450px; text-align: center;
    }
    
    /* 1. 徹底隱藏頂部裝飾線與 GitHub 連結，確保程式碼不外流 */
    header { visibility: hidden !important; }
    .stAppHeader { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    
    /* 2. 功能 2 修正：強制讓手機版側邊欄按鈕「浮」出來，解決縮小後找不到的問題 */
    [data-testid="stSidebarCollapsedControl"] {
        visibility: visible !important;
        display: flex !important;
        position: fixed !important;
        top: 10px !important;
        left: 10px !important;
        z-index: 999999 !important;
        background-color: #ffffff !important; 
        border-radius: 50% !important;
        box-shadow: 0px 2px 10px rgba(0,0,0,0.2) !important;
        width: 40px !important;
        height: 40px !important;
        justify-content: center !important;
        align-items: center !important;
    }

    /* 確保按鈕內的箭頭圖示是深色的，方便辨識點擊 */
    [data-testid="stSidebarCollapsedControl"] button {
        color: #333 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 產品資料庫與規格說明 (功能 4: 圖 1 文字要求) ---
if 'cart' not in st.session_state:
    st.session_state.cart = {} 

# 圖 1 的規格文字說明庫 (當選購對應馬力時，自動在 Excel 插入)
product_specs = {
    "10馬永磁變頻空壓機HCV-10PM-A": "HCV高端系列\n型號:HCV-10PM-A\n馬力:1馬至10馬<隨壓力調整馬力>\n噪音:68±5\nIE4永磁高效率馬達<馬達無軸承><無皮帶>\n5kg~8kg可選變頻壓力\nLCD液晶顯示預警/警告/錯誤跳機保護\n外型尺寸:745*680*910(mm)\n變頻器與電機一體化非外掛變頻空壓機\n啟動電流衝擊小,恆壓恆溫控制,故障率低",
    "20馬永磁變頻空壓機HCV-20PM-A": "HCV高端系列\n型號:HCV-20PM-A\n規格同高端系列，提供更強勁風量輸出",
    "30馬永磁變頻空壓機HCV-30PM-A": "HCV高端系列\n型號:HCV-30PM-A\n適合中大型工廠，高效穩定",
    # 其他規格可依此格式繼續增加...
}

products = {
    "空壓機": {
        "HCV高端系列": [
            ("10馬永磁變頻空壓機HCV-10PM-A", "air_10.png"), ("30馬永磁變頻空壓機HCV-30PM-A", "air_30.png"), 
            ("75馬永磁變頻空壓機HCV-75PM-A", "air_75.png"), ("20馬永磁變頻空壓機HCV-20PM-A", "air_20.png"),
            ("50馬永磁變頻空壓機HCV-50PM-A", "air_50.png"), ("100馬永磁變頻空壓機HCV-100PM-A", "air_100.png")
        ],
        "標準型": [
            ("5馬空壓機", "air_5.png")
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
    def init_prices(d):
        for k, v in d.items():
            if isinstance(v, dict): init_prices(v)
            elif isinstance(v, list):
                for name, _ in v: st.session_state.price_config[name] = 0
    init_prices(products)

# --- 3. 側邊欄 (後台管理視窗) ---
st.sidebar.title("🏢 翌新後台管理")
customer_name = st.sidebar.text_input("客戶名稱", value="")
contact_person = st.sidebar.text_input("聯絡人", value="")

# 功能 3: 增加電壓選擇
voltage = st.sidebar.radio("電力規格 (輸出於Excel)", ["220V", "380V"], horizontal=True)

with st.sidebar.expander("⚙️ 價格調整"):
    for name in sorted(st.session_state.price_config.keys()):
        st.session_state.price_config[name] = st.sidebar.number_input(f"{name} 單價", value=st.session_state.price_config[name], step=100)

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
    # 功能 1: 空壓機分頁選擇
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
    st.markdown(f"### <span class='price-text'>總計金額：${total_val:,} (電力：{voltage})</span>", unsafe_allow_html=True)

    template_path = "翌新估價單EXCELNEW.xlsx"
    if os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        bold_font = Font(name='新細明體', size=11, bold=True)
        spec_font = Font(name='新細明體', size=10)
        right_align = Alignment(horizontal='right', vertical='center')
        center_align = Alignment(horizontal='center', vertical='center')
        top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # 填入基本資訊
        ws['B11'] = customer_name
        ws['B12'] = contact_person
        ws['H14'] = f"估價日期：{datetime.now().strftime('%Y-%m-%d')}"
        
        current_row = 17
        for i, (name, qty) in enumerate(st.session_state.cart.items()):
            price = st.session_state.price_config.get(name, 0)
            
            # 品名(含電壓)、單位、數量、單價、金額
            ws.cell(row=current_row, column=1, value=i+1).font = bold_font
            ws.cell(row=current_row, column=2, value=f"{name} ({voltage})").font = bold_font
            ws.cell(row=current_row, column=6, value=unit_map.get(name, "台")).alignment = center_align
            ws.cell(row=current_row, column=7, value=qty).alignment = center_align
            ws.cell(row=current_row, column=8, value=price).alignment = right_align
            ws.cell(row=current_row, column=9, value=price * qty).alignment = right_align
            
            # 功能 4: 插入產品說明 (例如 10馬永磁變頻...)
            if name in product_specs:
                current_row += 1
                spec_cell = ws.cell(row=current_row, column=2, value=product_specs[name])
                spec_cell.font = spec_font
                spec_cell.alignment = top_left_align
                ws.row_dimensions[current_row].height = 160 # 足夠容納多行文字
            
            current_row += 1

        # 合計與底部標籤 (對應您的模板位置)
        ws['I36'] = total_val
        ws['H36'] = "合計："

        output = io.BytesIO()
        wb.save(output)
        st.download_button(label="📤 下載 翌新專業報價單 (Excel)", data=output.getvalue(), file_name=f"翌新報價_{customer_name}.xlsx")

    if st.button("🗑️ 清空重選"):
        st.session_state.cart = {}
        st.rerun()
