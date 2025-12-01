#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# 1.导入必要的库
import streamlit as st
import pandas as pd  


# In[ ]:


# 2.核心算法
def calculate_sum(file_path):
    df = pd.read_excel(file_path)
    sum_result = df.sum(numeric_only=True)
    return sum_result, df


# In[ ]:


# 3. Streamlit UI设计
# 3.1 页面标题和说明
st.title("Excel数据自动求和工具")
st.write("📊 上传Excel文件，自动计算每列的数值总和（支持.xlsx格式）")
st.divider()  # 加一条分割线，UI更整洁

# 3.2 文件上传组件
uploaded_file = st.file_uploader("请上传Excel文件", type="xlsx")  # 只允许上传.xlsx文件

# 3.3 计算按钮（用户点击后才执行算法，避免一上传就运行）
if st.button("开始计算", type="primary"):  # primary是蓝色主按钮，更醒目
    # 容错处理：如果用户没上传文件就点击按钮，提示错误
    if not uploaded_file:
        st.error("❌ 请先上传Excel文件再计算！")
    else:
        # 执行算法
        with st.spinner("正在计算中..."):
            sum_result, original_df = calculate_sum(uploaded_file)#核心赋值语句
        
        # 3.4 展示结果（分区域显示，用户看得清楚）
        st.success("✅ 计算完成！")
        
        # 显示原始数据（前5行）
        st.subheader("原始数据预览")
        st.dataframe(original_df.head(), use_container_width=True)  # use_container_width让表格自适应宽度
        
        # 显示求和结果
        st.subheader("每列求和结果")
        st.dataframe(sum_result, use_container_width=True)
        
        # 3.5 下载结果
        # 把求和结果转成Excel文件
        sum_result_df = pd.DataFrame(sum_result, columns=["总和"])
        with pd.ExcelWriter("求和结果.xlsx", engine="openpyxl") as writer:
            sum_result_df.to_excel(writer, sheet_name="求和结果", index=True)
        
        # 提供下载按钮
        with open("求和结果.xlsx", "rb") as f:
            st.download_button(
                label="📥 下载求和结果",
                data=f,
                file_name="Excel求和结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            

