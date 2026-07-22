import streamlit as st
import pandas as pd
from io import BytesIO
import os
import datetime

# --- 1. 页面配置 ---
st.set_page_config(page_title="巴西TK处理工具", layout="wide")
st.title("🇧🇷 巴西TK处理工具")
st.caption("系统已内置 A 文件和 C 文件。您只需上传 B 文件并填写店铺网址，即可一键完成数据处理。")

# --- 2. 辅助函数：智能查找并重命名关键列 ---
def standardize_order_id(df, file_label):
    """
    在 DataFrame 中查找类似 'OrderID' 的列，并将其重命名为 'OrderID'。
    如果找不到，则报错提示。
    """
    # 常见的订单号列名写法
    candidates = ['orderid', 'order id', 'order_id', '订单号', '平台唯一单号', 'platform unique order id']
    
    # 获取所有列名并转为小写用于比对
    col_map = {str(c).strip().lower(): c for c in df.columns}
    
    found_col = None
    for cand in candidates:
        if cand in col_map:
            found_col = col_map[cand]
            break
            
    if found_col:
        # 重命名并返回
        return df.rename(columns={found_col: 'OrderID'}), True
    else:
        st.error(f"❌ 在 {file_label} 中找不到订单号列！当前列名：{list(df.columns)}")
        return df, False

# --- 3. 侧边栏或顶部输入区 ---
file_b = st.file_uploader("上传巴西TK原始订单文件 (B.xlsx)", type=["xlsx", "xls"])
shop_url = st.text_input("填写店铺网址 (必填)", placeholder="https://vt.tiktok.com/...")

# 模拟加载内置文件 (请确保你的项目目录下有这两个文件)
# 如果没有，这里会报错，请根据实际文件名修改
try:
    # 假设 A 文件和 C 文件就在当前目录下
    df_a_raw = pd.read_excel('A_final_complete (1).xlsx') 
    df_c_raw = pd.read_excel('车车-2024.xlsx')
    files_loaded = True
except Exception as e:
    st.warning(f"⚠️ 无法加载内置文件: {e}")
    files_loaded = False

# --- 4. 核心处理逻辑 ---
if st.button("🚀 开始处理"):
    if not file_b or not shop_url:
        st.error("请上传文件并填写店铺网址！")
    elif not files_loaded:
        st.error("内置文件加载失败，请检查服务器文件！")
    else:
        try:
            # 1. 读取上传的 B 文件
            df_b = pd.read_excel(file_b)
            
            # 2. 标准化三张表的订单号列名
            df_a, ok_a = standardize_order_id(df_a_raw, "A文件")
            df_b, ok_b = standardize_order_id(df_b, "B文件(上传)")
            df_c, ok_c = standardize_order_id(df_c_raw, "C文件")
            
            if not (ok_a and ok_b and ok_c):
                st.stop() # 如果有任何一个文件找不到订单号列，停止运行

            # 3. 执行合并 (左连接，保留 B 文件的所有行)
            # 先合并 A
            result = pd.merge(df_b, df_a, on='OrderID', how='left', suffixes=('_B', '_A'))
            # 再合并 C
            result = pd.merge(result, df_c, on='OrderID', how='left', suffixes=('', '_C'))
            
            # 4. 生成店铺链接 (覆盖或新增一列)
            # 注意：这里假设你要用 B 文件的 OrderID 拼接链接
            result['店铺链接'] = f"{shop_url.strip('/')}/page=TikTokShop&product_id={result['OrderID']}" 
            
            # 调整列顺序，把重要的放前面 (可选)
            cols = ['OrderID', '店铺链接'] + [c for c in result.columns if c not in ['OrderID', '店铺链接']]
            result = result[cols]

            # --- 5. 显示结果与下载 ---
            st.success(f"处理完成！共生成 {len(result)} 条数据。")
            st.dataframe(result.head()) # 预览前几行
            
            # 生成 Excel 二进制流
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result.to_excel(writer, index=False, sheet_name='Processed_Data')
            
            processed_data = output.getvalue()
            
            # 生成下载按钮
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 点击下载处理结果",
                data=processed_data,
                file_name=f"Brazil_TK_Processed_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"处理出错: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
