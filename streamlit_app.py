import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import datetime

# --- 页面配置 ---
st.set_page_config(page_title="巴西TK处理工具", layout="wide")
st.title("🇧🇷 巴西TK处理工具")
st.caption("系统已内置 A 文件和 C 文件。您只需上传 B 文件并填写店铺网址，即可一键完成数据处理。")

# --- 1. 辅助函数：智能查找列名 ---
def find_column(df, possible_names):
    """
    在 DataFrame 中查找目标列。
    支持模糊匹配，忽略大小写和首尾空格。
    """
    # 获取所有列名并清洗（去空格、转小写）用于比对
    clean_cols = {col: str(col).strip().lower() for col in df.columns}
    
    for name in possible_names:
        target = name.lower()
        # 检查是否有列清洗后等于目标名称
        for original_col, cleaned_col in clean_cols.items():
            if cleaned_col == target:
                return original_col # 返回原始列名（保留大小写）
    return None

# --- 2. 加载内置文件 (A & C) ---
@st.cache_data
def load_builtin_files():
    try:
        # 请确保你的仓库里有这两个文件，或者修改为实际文件名
        # 如果文件不在根目录，需要加上路径，例如 'data/A.xlsx'
        df_a = pd.read_excel("A.xlsx") 
        df_c = pd.read_excel("C.xlsx")
        return df_a, df_c
    except Exception as e:
        st.error(f"加载内置文件失败: {e}")
        return None, None

df_a, df_c = load_builtin_files()

# --- 3. 用户输入区域 ---
col1, col2 = st.columns([3, 1])
with col1:
    file_b = st.file_uploader("上传巴西TK原始订单文件 (B.xlsx)", type=["xlsx", "xls", "csv"])
with col2:
    shop_url = st.text_input("填写店铺网址 (必填)", placeholder="https://vt.tiktok.com/...")

# --- 4. 处理逻辑 ---
if st.button("🚀 开始处理", type="primary"):
    if not file_b or not shop_url:
        st.warning("请上传 B 文件并填写店铺网址！")
    elif df_a is None or df_c is None:
        st.error("系统内置文件缺失，请检查仓库文件。")
    else:
        try:
            # 读取上传的 B 文件
            if file_b.name.endswith('.csv'):
                df_b = pd.read_csv(file_b)
            else:
                df_b = pd.read_excel(file_b)

            progress_bar = st.progress(0)
            status_text = st.empty()

            # === 步骤 1: 数据清洗与匹配 ===
            status_text.text("正在执行第 1 步：清洗数据...")
            
            # 【关键修复】智能查找 OrderID 列
            # 这里列出了所有可能的列名写法，代码会自动去匹配
            order_id_col = find_column(df_b, ["OrderID", "Order ID", "order_id", "订单号", "订单编号"])
            
            if order_id_col is None:
                st.error(f"❌ 找不到订单号列！\n当前文件包含的列：{list(df_b.columns)}")
                st.stop()
            
            # 统一重命名为 'OrderID' 方便后续处理
            df_b.rename(columns={order_id_col: "OrderID"}, inplace=True)
            
            # 模拟业务逻辑 (请根据实际需求替换此处)
            # 假设我们要把 A 表的某些信息合并到 B 表
            # result_df = pd.merge(df_b, df_a[['OrderID', 'SomeInfo']], on='OrderID', how='left')
            
            # 这里为了演示，我们只做简单的标记
            result_df = df_b.copy()
            result_df['店铺链接'] = shop_url
            
            progress_bar.progress(50)
            status_text.text("正在执行第 2 步：计算与生成...")
            
            # 模拟耗时操作
            import time
            time.sleep(1) 
            
            progress_bar.progress(100)
            status_text.text("处理完成！✅")

            # === 步骤 2: 生成下载文件 ===
            # 将结果转换为 Excel 二进制流
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='处理结果')
            
            processed_data = output.getvalue()
            
            # 生成带时间戳的文件名
            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"Brazil_TK_Processed_{now_str}.xlsx"

            # 显示成功提示
            st.success("处理成功！请点击下方按钮下载文件。")
            
            # 显示下载按钮
            st.download_button(
                label="📥 点击下载 Excel 文件",
                data=processed_data,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"处理过程中发生未知错误: {e}")
            st.exception(e) # 显示详细报错信息以便调试
