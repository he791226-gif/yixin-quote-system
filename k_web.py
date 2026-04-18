import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="翌新空壓機報價系統", layout="wide")

# CSS 修正：徹底隱藏官方雜項與裝飾線，確保畫面乾淨
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .price-text { color: #E84118; font-size: 24px; font-weight: bold; }
    
    /* 徹底隱藏頂部裝飾、選單與官方圖標 (解決圈起處問題) */
    header, .stAppHeader, #MainMenu, footer, [data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 隱藏側邊欄控制鈕 */
    [data-testid="stSidebarCollapsedControl"] { display: none !important; }

    /* 產品展示卡片美化 */
    [data-testid="stVerticalBlock"] > div:has(div.stImage) {
        background-color: white; padding: 20px; border-radius: 15px;
        border: 1px solid #ddd; min-height: 450px; text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 產品資料庫與規格 ---
if 'cart' not in st.session_state:
    st.session_state.cart = {} 

product_specs = {
    "10馬永磁變頻空壓機HCV-10PM-A": "HCV高端系列\n型號:HCV-10PM-A\n馬力:1馬至10馬<隨壓力調整馬力>\n噪音:68±5\nIE4永磁高效率馬達<馬達無軸承><無皮帶>\n5kg~8kg可選變頻壓力\nLCD液晶顯示預警/警告/錯誤跳機保護\n外型尺寸:745*680*910(mm)\n變頻器與電機一體化非外掛變頻空壓機\n啟動電流衝擊小,恆壓恆溫控制,故障率低",
    "20馬永磁變頻空壓機HCV-20PM-A": "HCV高端系列\n型號:HCV-20PM-A\n規格同高端系列",
    "30馬永磁變頻空壓機HCV-30PM-A": "HCV高端系列\n型號:HCV-30PM-A"
}

# 重新定義產品結構 (功能要求：超精密過濾器獨立分頁)
products = {
    "空壓機": {
        "HCV高端系列": [
            ("10馬永磁變頻空壓機HCV-10PM-A", "air_10.png"), ("20馬永磁變頻空壓機HCV-20PM-A", "air_20.png"), 
            ("30馬永磁變頻空壓機HCV-30PM-A", "air_30.png"), ("50馬永磁變頻空壓機HCV-50PM-A", "air_50.png"),
            ("75馬永磁變頻空壓機HCV-75PM-A", "air_75.png"), ("100馬永磁變頻空壓機HCV-100PM-A", "air_100.png")
        ],
        "標準型": [("5馬空壓機", "air_5.png"), ("20馬永磁變頻空壓機HCV-20PM-A", "air_20.png"),
             ("30馬永磁變頻空壓機HCV-30PM-A", "air_30.png"), ("50馬永磁變頻空壓機HCV-50PM-A", "air_50.png"),]
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
    "超精密過濾器組": {
        "合成牌": [("超精密過濾器組(合成牌)", "filter_com.png")],
        "英國Parker超精密過濾器": [("超精密過濾器組(PARK)", "filter_park.png")]
    },
    "選配配件": [
        ("Ckd自動排水器", "drainer_ckd.png"), ("電子式自動排水器", "drainer_e.png"), 
        ("前置旋風分離器", "separator.png"), ("超精密過濾器芯", "filter_core.png")
    ]
}

unit_map = {
    "105儲氣筒": "只", "360儲氣筒": "只", "660儲氣筒": "只",
    "Ckd自動排水器": "只", "電子式自動排水器": "只", "前置旋風分離器": "只",
    "超精密過濾器組(合成牌)": "只", "超精密過濾器組(PARK)": "只", "超精密過濾器芯": "只"
}

# 初始化價格設定
if 'price_config' not in st.session_state:
    st.session_state.price_config = {}
    def init_prices(d):
        for k, v in d.items():
            if isinstance(v, dict): init_prices(v)
            elif isinstance(v, list):
                for name, _ in v: st.session_state.price_config[name] = 0
    init_prices(products)

# --- 3. 隱藏的後台管理參數 (放在前面確保全域抓得到) ---
# 將後台管理放在一個獨立區塊，不影響加入按鈕的邏輯
with st.container():
    col_empty, col_admin_btn = st.columns([3, 1])
    with col_admin_btn:
        with st.popover("⚙️ 開啟後台設定", use_container_width=True):
            customer_name = st.text_input("客戶名稱", key="input_cust")
            contact_person = st.text_input("聯絡人", key="input_cont")
            voltage = st.radio("電力規格", ["220V", "380V"], horizontal=True, key="input_volt")
            with st.expander("📝 調整單價"):
                for name in sorted(st.session_state.price_config.keys()):
                    st.session_state.price_config[name] = st.number_input(f"{name}", value=st.session_state.price_config[name], step=100, key=f"p_{name}")

# 預設值確保後續不報錯
if 'customer_name' not in locals(): customer_name = ""
if 'contact_person' not in locals(): contact_person = ""
if 'voltage' not in locals(): voltage = "220V"

# --- 4. 主展示介面 ---
st.title("請選擇設備類別")
tabs = st.tabs(["空壓機", "儲氣筒", "乾燥機", "超精密過濾器組", "選配配件"])

def display_items(item_list):
    cols = st.columns(3)
    for i, (name, img) in enumerate(item_list):
        with cols[i % 3]:
            if os.path.exists(img):
                st.image(img, width=250)
            st.write(f"**{name}**")
            # 修正：加入報價單按鈕
            if st.button(f"➕ 加入報價單", key=f"btn_{name}"):
                st.session_state.cart[name] = st.session_state.cart.get(name, 0) + 1
                st.rerun() # 強制刷新確保清單立即顯示

with tabs[0]:
    air_type = st.radio("選擇類型", ["HCV高端系列", "標準型"], horizontal=True)
    display_items(products["空壓機"][air_type])
with tabs[1]: display_items(products["儲氣筒"])
with tabs[2]:
    brand = st.radio("選擇品牌", ["宙升", "艾冷"], horizontal=True)
    display_items(products["乾燥機"][brand])
with tabs[3]:
    f_brand = st.radio("品牌選取", ["合成牌", "PARK"], horizontal=True)
    display_items(products["超精密過濾器組"][f_brand])
with tabs[4]: display_items(products["選配配件"])

# --- 5. 報價清單展示與 Excel 下載 ---
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

    # Excel 邏輯
    template_path = "翌新估價單EXCELNEW.xlsx"
    if os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        bold_font = Font(name='新細明體', size=11, bold=True)
        spec_font = Font(name='新細明體', size=10)
        
        ws['B11'] = customer_name
        ws['B12'] = contact_person
        ws['H14'] = f"估價日期：{datetime.now().strftime('%Y-%m-%d')}"
        
        current_row = 17
        for i, (name, qty) in enumerate(st.session_state.cart.items()):
            price = st.session_state.price_config.get(name, 0)
            ws.cell(row=current_row, column=1, value=i+1).font = bold_font
            ws.cell(row=current_row, column=2, value=f"{name} ({voltage})").font = bold_font
            ws.cell(row=current_row, column=6, value=unit_map.get(name, "台"))
            ws.cell(row=current_row, column=7, value=qty)
            ws.cell(row=current_row, column=8, value=price)
            ws.cell(row=current_row, column=9, value=price * qty)
            
            if name in product_specs:
                current_row += 1
                spec_cell = ws.cell(row=current_row, column=2, value=product_specs[name])
                spec_cell.font = spec_font
                ws.row_dimensions[current_row].height = 160 
            current_row += 1

        ws['I36'] = total_val
        ws['H36'] = "合計："

        output = io.BytesIO()
        wb.save(output)
        
        col_down, col_clear = st.columns(2)
        with col_down:
            st.download_button("📤 下載 Excel 報價單", data=output.getvalue(), file_name=f"翌新報價_{customer_name}.xlsx", use_container_width=True)
        with col_clear:
            if st.button("🗑️ 清空重選", use_container_width=True):
                st.session_state.cart = {}
                st.rerun()
