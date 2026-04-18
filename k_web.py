import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment

# --- 1. 頁面設定與 CSS (隱藏您圈起來的橘色圖示) ---
st.set_page_config(page_title="翌新空壓機報價系統", layout="wide")

st.markdown("""
    <style>
    /* 徹底隱藏頂部裝飾與官方元件 */
    header, .stAppHeader, #MainMenu, footer, [data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
    }
    [data-testid="stSidebarCollapsedControl"] { display: none !important; }
    .price-text { color: #E84118; font-size: 24px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 產品介紹區 (您要修改的地方就在這裡！) ---
# 這裡對應的是 Excel 匯出時會跟著顯示的詳細規格
product_specs = {
    # 格式為 "產品名稱": "規格內容"
    # 使用 \n 代表在 Excel 儲存格內換行
    "10馬永磁變頻空壓機HCV-10PM-A": (
        "HCV高端系列\n"
        "型號:HCV-10PM-A\n"
        "馬力:1馬至10馬<隨壓力調整馬力>\n"
        "噪音:68±5\n"
        "電壓:三相220V\n"
        "IE4永磁高效率馬達<馬達無軸承><無皮帶>\n"
        "5kg~8kg可選變頻壓力\n"
        "LCD液晶顯示預警/警告/錯誤跳機保護\n"
        "外型尺寸:745*680*910(mm)\n"
        "變頻器與電機一體化非外掛變頻空壓機\n"
        "啟動電流衝擊小,恆壓恆溫控制,故障率低"
    ),
    
    # 剩下的您可以照著上面的格式繼續新增，例如：
    # "20馬永磁變頻空壓機HCV-20PM-A": "規格內容...",
}

# --- 3. 產品清單結構 ---
products = {
    "空壓機": {
        "HCV高端系列": [
            ("10馬永磁變頻空壓機HCV-10PM-A", "air_10.png"), 
            ("20馬永磁變頻空壓機HCV-20PM-A", "air_20.png"),
            ("30馬永磁變頻空壓機HCV-30PM-A", "air_30.png")
        ],
        "標準型": [("5馬空壓機", "air_5.png")]
    },
    "儲氣筒": [("105儲氣筒", "tank_105.png")],
    "乾燥機": {"宙升": [("5馬宙升乾燥機SD-005", "zs_dryer_5.png")], "艾冷": []},
    "超精密過濾器組": {"合成牌": [("超精密過濾器組(合成牌)", "f_com.png")], "PARK": []},
    "選配配件": [("Ckd自動排水器", "drainer_ckd.png")]
}

# 單位對照表
unit_map = {"105儲氣筒": "只", "Ckd自動排水器": "只"}

# 初始化 Session State
if 'cart' not in st.session_state: st.session_state.cart = {}
if 'price_config' not in st.session_state:
    st.session_state.price_config = {}
    def init_p(d):
        for k, v in d.items():
            if isinstance(v, dict): init_p(v)
            elif isinstance(v, list):
                for n, _ in v: st.session_state.price_config[n] = 0
    init_p(products)

# --- 4. 後台管理 (位置移動到報價清單旁，不再置頂) ---
with st.container():
    c_space, c_admin = st.columns([3, 1])
    with c_admin:
        with st.popover("⚙️ 開啟後台設定"):
            cust_name = st.text_input("客戶名稱")
            cont_person = st.text_input("聯絡人")
            volt = st.radio("電力規格", ["220V", "380V"], horizontal=True)
            for n in sorted(st.session_state.price_config.keys()):
                st.session_state.price_config[n] = st.number_input(f"{n}", value=st.session_state.price_config[n], step=100)

# --- 5. 主分頁與按鈕 ---
st.title("請選擇設備類別")
tabs = st.tabs(["空壓機", "儲氣筒", "乾燥機", "超精密過濾器組", "選配配件"])

def display(items):
    cols = st.columns(3)
    for i, (name, img) in enumerate(items):
        with cols[i % 3]:
            if os.path.exists(img): st.image(img, width=200)
            st.write(f"**{name}**")
            # 增加 rerun 確保按鈕點擊後清單立即更新
            if st.button(f"➕ 加入報價單", key=name):
                st.session_state.cart[name] = st.session_state.cart.get(name, 0) + 1
                st.rerun()

with tabs[0]: display(products["空壓機"][st.radio("系列", ["HCV高端系列", "標準型"], horizontal=True)])
with tabs[1]: display(products["儲氣筒"])
with tabs[2]: display(products["乾燥機"][st.radio("品牌", ["宙升", "艾冷"], horizontal=True)])
with tabs[3]: display(products["超精密過濾器組"][st.radio("品牌 ", ["合成牌", "PARK"], horizontal=True)])
with tabs[4]: display(products["選配配件"])

# --- 6. 報價清單與 Excel 匯出 ---
st.divider()
if st.session_state.cart:
    st.subheader("📋 目前報價清單")
    total = 0
    data = []
    for name, qty in st.session_state.cart.items():
        p = st.session_state.price_config.get(name, 0)
        sub = p * qty
        total += sub
        data.append([name, unit_map.get(name, "台"), qty, f"${p:,}", f"${sub:,}"])
    
    st.table(pd.DataFrame(data, columns=["品名及規格", "單位", "數量", "單價", "金額"]))
    st.markdown(f"### <span class='price-text'>總計金額：${total:,} (電力：{volt})</span>", unsafe_allow_html=True)

    # Excel 處理邏輯
    tmp = "翌新估價單EXCELNEW.xlsx"
    if os.path.exists(tmp):
        wb = openpyxl.load_workbook(tmp)
        ws = wb.active
        ws['B11'], ws['B12'] = cust_name, cont_person
        ws['H14'] = f"估價日期：{datetime.now().strftime('%Y-%m-%d')}"
        
        row = 17
        for i, (name, qty) in enumerate(st.session_state.cart.items()):
            p = st.session_state.price_config.get(name, 0)
            ws.cell(row=row, column=1, value=i+1)
            ws.cell(row=row, column=2, value=f"{name} ({volt})").font = Font(bold=True)
            ws.cell(row=row, column=7, value=qty)
            ws.cell(row=row, column=8, value=p)
            ws.cell(row=row, column=9, value=p*qty)
            
            # 這裡就是把 10馬 規格自動塞進 Excel 的關鍵位置！
            if name in product_specs:
                row += 1
                c = ws.cell(row=row, column=2, value=product_specs[name])
                c.alignment = Alignment(wrap_text=True, vertical='top')
                ws.row_dimensions[row].height = 160 
            row += 1
        
        ws['I36'] = total
        out = io.BytesIO()
        wb.save(out)
        
        c_dn, c_cl = st.columns(2)
        with c_dn: st.download_button("📤 下載 Excel 報價單", data=out.getvalue(), file_name=f"報價單_{cust_name}.xlsx", use_container_width=True)
        with c_cl: 
            if st.button("🗑️ 清空重選", use_container_width=True):
                st.session_state.cart = {}; st.rerun()
