from typing import List
import streamlit as st
import pandas as pd
from nptyping import DataFrame

from occupational_health_module.occupational_health import OccupationalHealthItemInfo

st.set_page_config(layout="wide", initial_sidebar_state="auto")

st.title("职业卫生编号和记录表生成")
st.markdown("输入职业卫生项目的相应信息，会自动生成各个点位的样品编号")
st.header("输入数据")
st.subheader("输入项目基本信息")

col1, col2 = st.columns(2)
with col1:
    project_num: str = st.text_input("项目编号")
    # output_path: str = st.text_input("记录表输出路径", help="如果路径为空，则在桌面创建文件夹存放")
    # exploded: bool = st.checkbox("是否分为多列")
with col2:
    company_name: str = st.text_input("公司名称")
    types_order: List[str] = st.multiselect("样品类型顺序", ["空白", "定点", "个体"], ["空白", "定点", "个体"])


st.subheader("输入样品信息")
i_tab1, i_tab2 = st.tabs(["定点", "个体"])
with i_tab1:
    in_point_df: DataFrame = st.data_editor(
        pd.DataFrame([{
        "采样点编号": None,
        "单元": None,
        "检测地点": None,
        "工种": None,
        "日接触时间": None,
        "检测因素": None,
        "采样数量/天": None,
        "采样日程": None
        }]),
        num_rows="dynamic",
        use_container_width=False,
        key="point",
        column_config={
            "日接触时间": st.column_config.NumberColumn(format="%.2f"),
            "采样数量/天": st.column_config.NumberColumn(format="%d"),
            # "采样日程": st.column_config.NumberColumn(format="%d"),
        }
    )
with i_tab2:
    in_personnel_df: DataFrame = st.data_editor(
        pd.DataFrame([{
        "采样点编号": None,
        "单元": None,
        "工种": None,
        "日接触时间": None,
        "检测因素": None,
        "采样数量/天": None,
        "采样日程": None
        }]),
        num_rows="dynamic",
        use_container_width=False,
        key="personnel",
        column_config={
            "日接触时间": st.column_config.NumberColumn(format="%.2f"),
            "采样数量/天": st.column_config.NumberColumn(format="%d"),
            # "采样日程": st.column_config.NumberColumn(format="%d"),
        }
    )


run: bool = st.button("执行", key='run')

if run:
    occupy_num: int = 0
    occupational_health_info = OccupationalHealthItemInfo(
        company_name,
        project_num,
        in_point_df,
        in_personnel_df,
        types_order
    )
    # st.dataframe(occupational_health_info.output_deleterious_substance_info_dict['1']['个体'])
    is_process: bool = st.button('处理记录表', key='process')
    if is_process:
    # st.button('处理记录表', on_click=occupational_health_info.write_to_templates)
        try:
            occupational_health_info.write_to_templates()
            st.success(f"完成，已保存到{occupational_health_info.output_path}")
        except Exception:
            st.error('出现错误，无法进行')
