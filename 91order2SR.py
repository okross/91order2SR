import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="91 訂單轉換工具", page_icon="📦", layout="wide")

def process_91_data(df):
    # 進行必要的數據清理
    df['數量'] = pd.to_numeric(df['數量'], errors='coerce').fillna(1)
    df['銷售金額(折扣後)'] = pd.to_numeric(df['銷售金額(折扣後)'], errors='coerce').fillna(0)
    
    # 計算單價: 銷售金額(折扣後) / 數量
    df['商品單價'] = df['銷售金額(折扣後)'] / df['數量']
    
    # 映射欄位
    mapped_rows = []
    for _, row in df.iterrows():
        # 物流判斷：如果通路商為空或為空字串，則填新竹物流
        logistics = str(row.get('通路商', '')).strip()
        if logistics == "" or logistics.lower() == "nan":
            logistics = "新竹物流"
            
        mapped_rows.append({
            '订单编号': row.get('訂單編號', ''),
            '订单日期': row.get('轉單日期時間', ''),
            '订单币种': 'TWD',
            '订单金额': row.get('銷售金額(折扣後)', 0),
            '商品名称': row.get('商品名稱', ''),
            '商品数量': row.get('數量', 0),
            '商品单价': row['商品單價'],
            '店铺网址': 'https://www.goddess-shop.com/',
            '快递单号': str(row.get('配送編號', '')).replace("'", ""), # 去除單引號以免Excel變公式
            '物流企业名称': logistics,
            '电商平台英文名称': 'GDS'
        })
    return pd.DataFrame(mapped_rows)

st.header("📦 91 訂單轉檔工具")

uploaded_file = st.file_uploader("上傳 91 訂單 Excel", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    ext = uploaded_file.name.split('.')[-1].lower()
    df = pd.read_excel(uploaded_file, engine='openpyxl' if ext == 'xlsx' else 'xlrd')
    
    if st.button("🚀 開始轉換", type="primary"):
        result_df = process_91_data(df)
        
        st.subheader("📊 預覽轉換結果")
        st.dataframe(result_df.head())

        # 下載按鈕 (產生符合 91 模板格式)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            header = result_df.columns.tolist()
            # 寫入第1列版本資訊
            v_line = ["version", "20201013"] + [""] * (len(header) - 2)
            pd.DataFrame([v_line]).to_excel(writer, index=False, header=False, startrow=0)
            # 寫入第2列標題
            pd.DataFrame([header]).to_excel(writer, index=False, header=False, startrow=1)
            # 寫入數據 (從第3列開始)
            result_df.to_excel(writer, index=False, header=False, startrow=2)
        
        st.download_button(
            label="📥 下載轉換後的訂單檔案",
            data=buf.getvalue(),
            file_name="91_Order_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
