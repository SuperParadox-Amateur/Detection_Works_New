from typing import List, Dict
import numpy as np
import pandas as pd
import streamlit as st
from nptyping import DataFrame

from environment_noise_module.environment_noise import EnvNoise

st.set_page_config(layout="wide", initial_sidebar_state="auto")

distr_list: List[str] = [
    "正态分布",
    "拉普拉斯分布",
    "逻辑分布",
    "耿贝尔分布",
]

noise_degrees: List[str] = [
    '0类',
    '1类',
    '2类',
    '3类',
    '4类',
]

leq_generators: Dict[str, str] = {
    '原始': 'original',
    '近似': 'approximate',
    '积分': 'integral',
}


leq_generators_description: str = (
    r'**原始**：输入的LEQ值；'
    r'**近似**：使用近似公式$L_{\text{eq}}\approx L_{50} + (L_{10}-L_{90})^2 \div 60$计算；'
    r'**积分**：使用积分公式$L_{\text{eq}}=10\times\lg\left(\frac{1}{T}\int^{T}_{0}10^{0.1\cdot L_{\text{A}}}dt\right)$计算'
)

st.title("生成随机环境噪声数值")
st.header("输入基本信息")

col1, col2 = st.columns(2)
with col1:
    t_minute = st.number_input("监测时长(min)", value=10)
    freq_weighting = st.selectbox("频率计权方式", ["A", "B", "C"], index=0)
    noise_degree = st.selectbox("噪声等级", noise_degrees, index=0)
with col2:
    count_ps = st.number_input("每秒监测次数", value=10)
    time_weighting = st.selectbox("时间计权方式", ["F", "S", "I"], index=0)
    leq_generator = st.selectbox(
        "LEQ值生成方式",
        list(leq_generators.keys()),
        index=0,
        help=leq_generators_description,
    )
tag: str = st.text_input("输入数据标签")
st.header("输入基本数据")
distr_tab, random_tab = st.tabs(["分布", "随机"])
with distr_tab:
    distr_name = st.selectbox("选择分布", distr_list, index=0)
    in_df_distr: DataFrame = st.data_editor(
        pd.DataFrame({
            "日期时间": [
                pd.to_datetime("2023-01-01 00:32:07"),
                pd.to_datetime("2023-01-01 00:45:18"),
            ],
            "等效连续声级Leq": [60.0, 70.0],
            "标准差SD": [1.0, 1.0],
            '监测范围1': [33, 33],
            '监测范围2': [133, 133],
        }),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "日期时间": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm:ss"),
            "等效连续声级Leq": st.column_config.NumberColumn(format="%.1f"),
            "标准差SD": st.column_config.NumberColumn(format="%.1f"),
            '监测范围1': st.column_config.NumberColumn(format="%d"),
            '监测范围2': st.column_config.NumberColumn(format="%d"),
        }
    )
    # distr_submit: bool = st.button("运行", key="distr_btn")
with random_tab:
    in_df_random: DataFrame = st.data_editor(
        pd.DataFrame({
            "日期时间": [pd.to_datetime("2023-01-01 00:32:07"), pd.to_datetime("2023-01-01 00:45:18")],
            "随机值下限": [40, 40],
            "随机值上限": [65, 65],
            '监测范围1': [33, 33],
            '监测范围2': [133, 133],
        }),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "日期时间": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm:ss"),
            "随机值下限": st.column_config.NumberColumn(format="%d"),
            "随机值上限": st.column_config.NumberColumn(format="%d"),
            '监测范围1': st.column_config.NumberColumn(format="%d"),
            '监测范围2': st.column_config.NumberColumn(format="%d"),
        }
    )
    # random_submit: bool = st.button("运行", key="random_btn")

is_submit: bool = st.button("运行", key="submit_btn")

if is_submit:
    in_t_minute: int = int(t_minute)

    env_noise = EnvNoise(
    in_df_distr = in_df_distr,
    in_df_random = in_df_random,
    tag = 'Test',
    noise_degree='2类',
    t_minute=in_t_minute,
    freq_weighting=freq_weighting, # type: ignore
    count_ps=count_ps, # type: ignore
    time_weighting=time_weighting, # type: ignore
    leq_generator=leq_generators[leq_generator], # type: ignore
    )

    # 获得信息
    distr_noise_info_df: DataFrame = env_noise.create_distr_noise_info_df(distr_name) # type: ignore
    random_noise_info_df: DataFrame = env_noise.create_random_noise_info_df()

    distr_noise_txt: str = env_noise.generate_all_noise_info_str(distr_noise_info_df)
    random_noise_txt: str = env_noise.generate_all_noise_info_str(random_noise_info_df)

    # 输出结果
    output1_tab, output2_tab = st.tabs(["分布噪声结果", "随机噪声结果"])
    with output1_tab:
        st.download_button(
            label = "结果下载",
            data = distr_noise_txt,
            file_name=f"{tag}环境噪声值.txt",
            key="distr_download"
        )
        with st.expander("结果预览"):
            st.text(distr_noise_txt)
    with output2_tab:
        st.download_button(
            label = "结果下载",
            data = random_noise_txt,
            file_name=f"{tag}环境噪声值.txt",
            key="random_download"
        )
        with st.expander("结果预览"):
            st.text(random_noise_txt)
