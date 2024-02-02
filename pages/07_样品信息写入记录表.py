# %%
import streamlit as st
from typing import Dict, Any, List
import pandas as pd
from pandas.core.frame import DataFrame
from other_functions.docx_func import handle_point_samples

st.set_page_config(layout="wide", initial_sidebar_state="auto")

# %% 标题和基本信息输入
st.title("样品信息写入记录表")
st.markdown("将生成好的样品信息写入记录表模板中")
st.header("输入数据")
st.subheader("输入项目的基本信息")

# %% 输入项目的基本信息
col1, col2 = st.columns(2)
with col1:
    project_str: str = st.text_input("项目编号")
    company_name: str = st.text_input("公司名称")
    whether_devided: bool = st.checkbox("是否按照单元分开")
with col2:
    output_path: str = st.text_input("记录表输出路径")
    tag_str: str = st.text_input("标签")


in_point_df: DataFrame = st.data_editor(pd.DataFrame([{
    "采样点编号": None,
    "车间/单元": None,
    "检测地点/岗位": None,
    "工种": None,
    "日接触时间(h)": None,
    "检测项目": None,
    "采样数量/天": None,
    "样品编号": None,
}]),
    num_rows="dynamic",
    use_container_width=False,
    key="point",
)
action: bool = st.button("执行")

# %%
template_file_path: str = r"./模板/原始记录表模板.docx"
output_path_exists: bool = len(output_path) != 0
info_dict: Dict[str, Any] = {
    "code": project_str,
    "comp": company_name,
    "item": None,
}

# %%
if action and output_path_exists:
    point_df: DataFrame = in_point_df.fillna(" ")
    if whether_devided:
        workshop_list: List[str] = point_df["车间/单元"].unique().tolist()
        for workshop in workshop_list:
            tag_str: str = workshop
            workshop_df = point_df[point_df["车间/单元"] == workshop]
            handle_point_samples(info_dict, workshop_df, template_file_path, output_path, tag_str)
    else:
        handle_point_samples(info_dict, point_df, template_file_path, output_path, tag_str)
else:
    st.info("未向记录表模板写入样品编号")
