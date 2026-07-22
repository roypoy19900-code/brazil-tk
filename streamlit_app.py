import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

# --- 网页界面部分 ---
st.title("📊 巴西TK处理工具")
st.caption("系统已内置 寻汇模板 文件和 订单编号函数 文件。您只需上传 原始订单 文件并填写店铺网址即可。仅支持巴西TK，其他请移步到导航页。")

# 1. 修改上传框，支持 xlsx, xls, csv 三种格式
file_b = st.file_uploader("上传巴西TK原始订单文件（请勿修改）", type=["xlsx", "xls", "csv"])

# 2. URL 填写框
shop_url = st.text_input("填写店铺网址（必填）", placeholder="例如：https://www.yourshop.com")

# 检查内置文件是否存在
if not os.path.exists("A.xlsx") or not os.path.exists("C.xlsx"):
    st.error("❌ 错误：系统未检测到内置的 A.xlsx 或 C.xlsx 文件，请检查 GitHub 仓库！")
else:
    if st.button("🚀 开始处理"):
        if file_b and shop_url:  # 确保文件和URL都已填写
            try:
                # 1. 读取内置的 A 和 C 文件
                df_a = pd.read_excel("A.xlsx")
                df_c = pd.read_excel("C.xlsx")
                
                # 2. 根据上传文件的后缀名，自动选择读取方式
                if file_b.name.endswith('.csv'):
                    df_b = pd.read_csv(file_b)
                else:  # 如果是 .xlsx 或 .xls
                    df_b = pd.read_excel(file_b)
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # --- 第一步：筛选并删除B表格特定行 ---
                status_text.text("⏳ 正在执行第 1 步：筛选并删除特定行...")
                if 'Order Substatus' in df_b.columns:
                    df_b = df_b[~((df_b['Order Substatus'] == '运输中') | (df_b['Order Substatus'].isna()))]
                progress_bar.progress(10)
                
                # --- 第三步：引用公式计算订单ID ---
                status_text.text("⏳ 正在执行第 3 步：计算订单ID...")
                df_b['Pure_ID'] = df_b['OrderID'].astype(str).str.extract(r'(\d+)')[0]
                counts = df_b.groupby('Pure_ID').cumcount()
                df_b['Processed_ID'] = df_b['Pure_ID'] + counts.apply(lambda x: f'-{x}' if x > 0 else '')
                df_a.loc[2:2+len(df_b)-1, 'OrderID'] = df_b['Processed_ID'].values
                progress_bar.progress(20)
                
                # --- 第四步：日期格式转换与粘贴 ---
                status_text.text("⏳ 正在执行第 4 步：转换日期格式...")
                if 'ShipDate' in df_b.columns:
                    df_b['Formatted_Date'] = pd.to_datetime(df_b['ShipDate']).dt.strftime('%Y/%m/%d')
                    df_a.loc[2:2+len(df_b)-1, 'OrderDate'] = df_b['Formatted_Date'].values
                progress_bar.progress(30)
                
                # --- 第五步：固定字段填充 ---
                status_text.text("⏳ 正在执行第 5 步：填充固定字段...")
                df_a.loc[2:2+len(df_b)-1, 'Currency'] = 'BRL'
                df_a.loc[2:2+len(df_b)-1, 'Warehouse'] = 'TK'
                progress_bar.progress(40)
                
                # --- 第六步：金额清洗与填充 ---
                status_text.text("⏳ 正在执行第 6 步：清洗订单金额...")
                df_b['Clean_Amount'] = df_b['Order Amount'].astype(str).str.replace('BRL', '').str.replace(',', '.').str.strip()
                df_b['Clean_Amount'] = pd.to_numeric(df_b['Clean_Amount'], errors='coerce')
                df_a.loc[2:2+len(df_b)-1, 'OrderAmount'] = df_b['Clean_Amount'].values
                progress_bar.progress(50)
                
                # --- 第七、八、十步：基于订单ID匹配复制数据 ---
                status_text.text("⏳ 正在执行第 7-10 步：匹配商品与物流信息...")
                b_dict = df_b.set_index('Processed_ID').to_dict('index')
                
                a_ids = df_a.loc[2:2+len(df_b)-1, 'OrderID'].astype(str)
                product_names, quantities, tracking_ids, providers = [], [], [], []
                
                for aid in a_ids:
                    match = b_dict.get(aid, {})
                    product_names.append(match.get('Product Name', ''))
                    quantities.append(match.get('Quantity', 0))
                    tracking_ids.append(match.get('Tracking ID', ''))
                    providers.append(match.get('Shipping Provider Name', ''))
                    
                df_a.loc[2:2+len(df_b)-1, 'ProductName'] = product_names
                df_a.loc[2:2+len(df_b)-1, 'Quantity'] = quantities
                df_a.loc[2:2+len(df_b)-1, 'TrackingID'] = tracking_ids
                df_a.loc[2:2+len(df_b)-1, 'ShippingProvider'] = providers
                progress_bar.progress(80)
                
                # --- 第九步：商品单价计算 ---
                status_text.text("⏳ 正在执行第 9 步：计算商品单价...")
                df_a['UnitPrice'] = df_a['OrderAmount'] / df_a['Quantity']
                progress_bar.progress(85)

                # --- 填写店铺网址到H列 ---
                status_text.text("⏳ 正在填写店铺网址...")
                if '店铺网址' in df_a.columns:
                    df_a.loc[2:2+len(df_b)-1, '店铺网址'] = shop_url
                progress_bar.progress(90)
                
                # --- 第十一步：删除0金额 ---
                status_text.text("⏳ 正在执行第 11 步：清理0金额数据...")
                df_a = df_a[df_a['OrderAmount'] != 0]
                progress_bar.progress(100)
                
                status_text.text("") # 清空状态文字
                
                # 生成下载文件
                output = BytesIO()
                df_a.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                
                st.success("🎉 处理完成！所有 11 个步骤已成功执行。")
                st.download_button(
                    label="📥 下载处理后的 A 文件",
                    data=output,
                    file_name="A_final_processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"处理过程中发生错误：{str(e)}")
        else:
            st.warning("⚠️ 巴西TK原始订单文件并填写店铺网址后再点击处理！")
