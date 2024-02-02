from typing import Union

import streamlit as st
import pandas as pd
from nptyping import DataFrame

from occupational_noise_module.occupational_noise import OccupationalNoiseInfo

st.set_page_config(layout="wide", initial_sidebar_state="auto")

Number = Union[int, float]

st.title("随机噪声值生成和等效噪声值计算")
st.markdown("生成正态分布的随机噪声值和计算等效噪声值")



col1, col2 = st.columns(2)
with col1:
    scale_val: Number = st.number_input(r"输入标准差$\sigma$", value=1.0, min_value=0.0, step=0.1)
with col2:
    size: Number = st.number_input("输入噪声值数量", value=3, min_value=1, step=1)

st.subheader("输入基本数据")
noise_df: DataFrame = st.data_editor(
    pd.DataFrame([{
    '采样点编号': 1,
    '岗位': "测试",
    '日接触时间': 8.0,
    '每周工作天数': 6.0,
    '基准值': 80.0,
    }]),
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        '日接触时间': st.column_config.NumberColumn(format="%.2f"),
        '每周工作天数': st.column_config.NumberColumn(format="%.2f"),
        '基准值': st.column_config.NumberColumn(format="%.2f"),
    },
)

size_int: int = int(size)

action: bool = st.button("执行")

if action:
    noise_info = OccupationalNoiseInfo(noise_df, scale_val, size_int)
    xlsx_df: bytes = noise_info.get_df_to_xlsx(noise_info.new_noise_df)

    st.header("计算结果")

    st.download_button("下载", data=xlsx_df, file_name="等效噪声值.xlsx", help="下载等效噪声值表格文件")
    st.dataframe(noise_info.new_noise_df, use_container_width=True)
