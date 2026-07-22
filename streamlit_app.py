import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

# --- 页面配置 ---
st.set_page_config(page_title="巴西TK处理工具", layout="wide")
st.title("🇧🇷 巴西TK处理工具")
st.caption("系统已内置 A 文件和 C 文件。您只需上传 B 文件并填写店铺网址，即可一键完成数据处理。")

# --- 1. 文件上传与输入 ---
file_b = st.file_uploader("上传巴西TK原始订单文件 (请勿修改)", type=["xlsx", "xls", "csv"])
shop_url = st.text_input("填写店铺网址 (必填)", placeholder="https://vt.tiktok.com/...")

# 检查内置文件
if not os.path.exists("A.xlsx") or not os.path.exists("C.xlsx"):
    st.error("❌ 系统错误：未检测到内置的 A.xlsx 或 C.xlsx 文件！")
else:
    if st.button("🚀 开始处理"):
        if file_b and shop_url:
            try:
                progress_bar = st.progress(0)
                status_text = st.empty()

                # --- 2. 读取文件 ---
                status_text.text("正在读取文件...")
                
                # 读取 B 文件 (根据后缀自动判断)
                if file_b.name.endswith('.csv'):
                    df_b = pd.read_csv(file_b)
                else:
                    df_b = pd.read_excel(file_b)
                
                # 读取内置文件
                df_a = pd.read_excel("A.xlsx")
                df_c = pd.read_excel("C.xlsx")

                # --- 3. 智能清洗列名 (关键修复步骤) ---
                # 去除所有列名首尾的空格，防止 " OrderID" 这种坑
                df_b.columns = df_b.columns.str.strip()
                df_a.columns = df_a.columns.str.strip()
                df_c.columns = df_c.columns.str.strip()

                # 打印前几行列名用于调试 (可选，上线后可注释掉)
                # st.write("B文件检测到的列名:", list(df_b.columns)) 

                # 检查是否存在 OrderID (兼容常见变体)
                target_col = None
                possible_names = ['OrderID', 'Order ID', '订单号', 'order_id']
                
                for name in possible_names:
                    if name in df_b.columns:
                        target_col = name
                        break
                
                if not target_col:
                    st.error(f"❌ 找不到订单号列！当前检测到的列名为：{list(df_b.columns)}")
                    st.stop() # 停止执行

                # --- 4. 模拟你的 11 步处理逻辑 (示例) ---
                # 假设第3步是计算订单ID
                status_text.text("正在执行第 3 步：计算订单ID...")
                progress_bar.progress(30)
                
                # 这里是你原本报错的地方，现在用 target_col 代替硬编码的 'OrderID'
                # 例如：df_b['新ID'] = df_b[target_col].astype(str) + "_suffix" 
                # (由于我看不到你原本的11步具体代码，这里仅作示意，请把你原本的逻辑接在这里)
                
                # ... 继续后续步骤 ...
                status_text.text("正在执行后续步骤...")
                progress_bar.progress(60)
                
                # 假设最后生成了结果
                # result_df = ... 
                
                status_text.text("处理完成！✅")
                progress_bar.progress(100)
                st.success("处理成功！请下载下方文件。")
                
                # st.download_button(...)

            except Exception as e:
                st.error(f"处理过程中发生未知错误: {str(e)}")
                import traceback
                st.code(traceback.format_exc()) # 显示详细报错，方便排查
        else:
            st.warning("请上传文件并填写网址。")
