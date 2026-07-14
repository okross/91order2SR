import streamlit as st
import pandas as pd
import io
import datetime

st.set_page_config(page_title="91 訂單轉換工具", page_icon="📦", layout="wide")

def process_91_data(df):
    # 進行必要的數據清理
    df['轉單日期時間'] = pd.to_datetime(df['轉單日期時間'])
    df['數量'] = pd.to_numeric(df['數量'], errors='coerce').fillna(1)
    df['銷售金額(折扣後)'] = pd.to_numeric(df['銷售金額(折扣後)'], errors='coerce').fillna(0)
    
    # 1. 計算日期過濾條件 (350天內，不包含今天以前的概念，即日期 >= 今天 - 350天)
    today = datetime.datetime.now()
    limit_date = today - datetime.timedelta(days=350)
    
    # 篩選數據
    df_filtered = df[df['轉單日期時間'] >= limit_date]
    excluded_count = len(df) - len(df_filtered)
    
    # 計算單價: 銷售金額(折扣後) / 數量
    df_filtered = df_filtered.copy() # 避免 SettingWithCopyWarning
    df_filtered['商品單價'] = df_filtered['銷售金額(折扣後)'] / df_filtered['數量']
    
    # 映射欄位
    mapped_rows = []
    for _, row in df_filtered.iterrows():
        # 物流判斷
        logistics = str(row.get('通路商', '')).strip()
        if logistics == "" or logistics.lower() == "nan":
            logistics = "新竹物流"
            
        mapped_rows.append({
            '订单编号': row.get('訂單編號', ''),
            '订单日期': row['轉單日期時間'].strftime('%Y-%m-%d %H:%M:%S'),
            '订单币种': 'TWD',
            '订单金额': row.get('銷售金額(折扣後)', 0),
            '商品名称': row.get('商品名稱', ''),
            '商品数量': row.get('數量', 0),
            '商品单价': row['商品單價'],
            '店铺网址': 'https://www.goddess-shop.com/',
            '快递单号': str(row.get('配送編號', '')).replace("'", ""),
            '物流企业名称': logistics,
            '电商平台英文名称': 'GDS'
        })
    
    return pd.DataFrame(mapped_rows), excluded_count

st.header("📦 91 訂單轉檔與財務統計工具")

uploaded_file = st.file_uploader("上傳 91 訂單 Excel", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    ext = uploaded_file.name.split('.')[-1].lower()
    df = pd.read_excel(uploaded_file, engine='openpyxl' if ext == 'xlsx' else 'xlrd')
    
    if st.button("🚀 開始處理", type="primary"):
        # 處理資料
        result_df, excluded = process_91_data(df)
        total_twd = result_df['订单金额'].sum()
        
        # 顯示統計資訊
        st.subheader("📊 統計結果")
        st.info(f"共排除 {excluded} 筆超過 350 天的訂單")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("總成交金額 (TWD)", f"{total_twd:,.0f}")
        c2.metric("約合美金 (USD)", f"${total_twd * 0.031:,.2f}")
        c3.metric("約合人民幣 (CNY)", f"¥{total_twd * 0.22:,.2f}")
        
        st.subheader("📋 預覽轉換數據")
        st.dataframe(result_df.head())

        # 準備 Excel 下載 (符合 91 模板格式)
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
            file_name="91_Order_Export_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
