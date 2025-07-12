import streamlit as st
import pandas as pd
from streamlit_autorefresh import st_autorefresh
from PIL import Image


# 每 10 秒自动刷新页面
st_autorefresh(interval=5 * 1000, key="datarefresh")

# 替换为你本地的 Excel 路径
file_paths = [
    r"C:\Users\Lenovo\OneDrive - Duke University\Work\Pfizer\quiz_automation\test3_auto.xlsx",
    r"C:\Users\Lenovo\OneDrive - Duke University\Work\Pfizer\quiz_automation\test4_auto.xlsx"
]
FILE_PATH = r"C:\Users\Lenovo\OneDrive - Duke University\Work\Pfizer\quiz_automation\summary_auto.xlsx"

def merge_excel_files():
    dfs = []
    for path in file_paths:
        df = pd.read_excel(path)
        # 每个表单内部去重
        df = df.drop_duplicates(subset=['Name'], keep='first')
        dfs.append(df)
    # 合并所有表单
    combined_df = pd.concat(dfs, ignore_index=True)
    # 汇总总积分
    summary_df = combined_df.groupby('Name', as_index=False)['Total points'].sum()
    summary_df.to_excel(FILE_PATH, index=False)
    return summary_df

# logo = Image.open(r"C:\Users\Lenovo\Desktop\logo-pfizer-1024.png")
# st.image(logo, width=150)
# 加载数据
st.title("Ranking List")


# 实时合并并加载数据
df = merge_excel_files()

# 排序数据
df_sorted = df.sort_values('Total points', ascending=False)




# 创建两列布局
col1, col2 = st.columns(2)
col1, _, col2 = st.columns([1, 0.2, 1]) 

# 显示前5名的柱状图在左侧
with col1:
    st.subheader("Top 5 Participants")
    top_5 = df_sorted.head(5)
    st.bar_chart(top_5.set_index('Name')['Total points'])

# 显示所有人的信息表格在右侧
with col2:
    st.subheader("All Participants (Sorted by Total Points)")
    st.dataframe(df_sorted)