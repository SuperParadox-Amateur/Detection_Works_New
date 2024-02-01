# %%
'''生成随机环境噪声值'''

# %%
from typing import List, Dict
import streamlit as st
import numpy as np
import pandas as pd
from pandas.core.frame import DataFrame
# %%
st.set_page_config(layout="wide", initial_sidebar_state="auto")

# %%
def get_noise_info(
    noise_array,
    current_datetime,
    l_eq,
    t_minute,
    sel,
    freq_weighting,
    time_weighting
    ) -> str:
    lmax: float = float(noise_array.max().round(1))
    lmin: float = float(noise_array.min().round(1)) # 最小值和最大值
    l5: float = float(np.percentile(noise_array, 5).round(1))
    l10: float = float(np.percentile(noise_array, 10).round(1))
    l50: float = float(np.percentile(noise_array, 50).round(1))
    l90: float = float(np.percentile(noise_array, 90).round(1))
    l95: float = float(np.percentile(noise_array, 95).round(1))
    sd: float = float(noise_array.std().round(1))

    current_info_list: List[str] = [
        f'{current_datetime}',
        'Stat.-One',
        f'R: 28dB~133dB Ts=00h{t_minute}m00s',
        f'Statistics: {freq_weighting} {time_weighting}',
        f'Leq,T= {l_eq}dB SEL  = {sel}dB',
        f'Lmax = {lmax}dB Lmin = {lmin}dB',
        f'L5   = {l5}dB L10  = {l10}dB',
        f'L50  = {l50}dB L90  = {l90}dB',
        f'L95  = {l95}dB SD   = {str(sd).rjust(4, " ")}dB',
        "\r\n"
    ]
    noise_string: str = "\r\n".join(current_info_list)
    return noise_string
# %%
distr_dict = {
    "正态分布": np.random.normal,
    "拉普拉斯分布": np.random.laplace,
    "逻辑分布": np.random.logistic,
    "耿贝尔分布": np.random.gumbel,
}

# %
st.title("生成随机环境噪声数值")
st.header("输入基本信息")
col1, col2 = st.columns(2)
with col1:
    t_minute = st.number_input("监测时长(min)", value=10)
    freq_weighting = st.selectbox("频率计权方式", ["A", "B", "C"])
with col2:
    count_ps = st.number_input("每秒监测次数", value=2)
    time_weighting = st.selectbox("时间计权方式", ["F", "S", "I"])
tag: str = st.text_input("输入数据标签")
st.header("输入基本数据")
distr_tab, random_tab = st.tabs(["分布", "随机"])
with distr_tab:
    distr_name = st.selectbox("选择分布", list(distr_dict.keys()))
    in_df_distr: DataFrame = st.data_editor(
        pd.DataFrame([{
            "日期时间": pd.to_datetime("2023-01-01 00:00:00"),
            "等效连续声级": 80.0,
            "标准差": 1.0,
        }]),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "日期时间": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm:ss"),
            "等效连续声级": st.column_config.NumberColumn(format="%.1f"),
            "标准差": st.column_config.NumberColumn(format="%.1f"),
        }
    )
    distr_submit: bool = st.button("运行", key="distr_btn")
with random_tab:
    in_df_random: DataFrame = st.data_editor(
        pd.DataFrame([{
            "日期时间": pd.to_datetime("2023-01-01 00:00:00"),
            "随机值下限": 60,
            "随机值上限": 80,
        }]),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "日期时间": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm:ss"),
            "随机值下限": st.column_config.NumberColumn(format="%d"),
            "随机值上限": st.column_config.NumberColumn(format="%d"),
        }
    )
    random_submit: bool = st.button("运行", key="random_btn")
# % 生成结果
if distr_submit:
    dtype_dict: Dict = {
        # "日期时间": np.datetime64,
        # "日期时间": datetime,
        "等效连续声级": float,
        "标准差": float,
    }

    df: DataFrame = in_df_distr.astype(dtype_dict)

    t: int = int(t_minute) * 60
    count: int = t * count_ps # type: ignore

    noise_text_list: List[str] = [tag, "\r\n"]
    for i in range(df.shape[0]):
        current_datetime = str(df.loc[i, "日期时间"]) # 测量日期时间
        l_eq = float(df.loc[i, "等效连续声级"]) # type: ignore 等效声级
        scale = float(df.loc[i, "标准差"]) # type: ignore 分布的偏差值
        v: float = float(10.0 * np.log10(t)) # 暴露声级和等效声级之间的差
        sel: float = round(l_eq + v, 1) # 暴露声级

        noise_array = distr_dict[distr_name](l_eq, scale, count) # 分布的随机噪声值数组
        noise_info: str = get_noise_info(
            noise_array,
            current_datetime,
            l_eq,
            t_minute,
            sel,
            freq_weighting,
            time_weighting
        ) # 噪声信息
        noise_text_list.append(noise_info)

    output_str: str = "\r\n".join(noise_text_list)

    st.download_button(label = "结果下载", data = output_str, file_name=f"{tag}环境噪声值.txt")
    with st.expander("结果预览"):
        st.text(output_str)

if random_submit:
    dtype_dict: Dict = {
        # "日期时间": np.datetime64,
        # "日期时间": datetime,
        "随机值下限": int,
        "随机值上限": int,
    }

    df: DataFrame = in_df_random.astype(dtype_dict)

    t: int = int(t_minute) * 60
    count: int = t * count_ps # type: ignore

    noise_text_list: List[str] = [tag, "\r\n"]
    for i in range(df.shape[0]):
        current_datetime = str(df.loc[i, "日期时间"]) # 测量日期时间
        lower_range_value = int(df.loc[i, "随机值下限"]) # type: ignore 随机值下限
        upper_range_value = int(df.loc[i, "随机值上限"]) # type: ignore 随机值上限

        noise_array = np.random.randint(
            lower_range_value * 10,
            upper_range_value * 10,
            count
        ) / 10 # 正态分布的随机噪声值数组
        l_eq: float = float(np.round(np.mean(noise_array), 1)) # 计算等效声值
        v: float = float(10.0 * np.log10(t)) # 暴露声级和等效声级之间的差
        sel: float = round(l_eq + v, 1) # 暴露声级
        noise_info: str = get_noise_info(
            noise_array,
            current_datetime,
            l_eq,
            t_minute,
            sel,
            freq_weighting,
            time_weighting
        ) # 噪声信息
        noise_text_list.append(noise_info)

    output_str: str = "\r\n".join(noise_text_list)
    st.download_button(label = "结果下载", data = output_str, file_name=f"{tag}环境噪声值.txt")
    with st.expander("结果预览"):
        st.text(output_str)
