import pandas as pd
from nptyping import DataFrame
import streamlit as st

from occupational_health_module.new_occupational_health import NewOccupationalHealthItemInfo

@st.cache_data
def get_raw_df(file_path) -> DataFrame:
    raw_df: DataFrame = pd.read_excel(file_path)
    return raw_df

col1, col2 = st.columns(2)
with col1:
    project_number: str = st.text_input("项目编号")
    # output_path: str = st.text_input("记录表输出路径", help="如果路径为空，则在桌面创建文件夹存放")
    # exploded: bool = st.checkbox("是否分为多列")
with col2:
    company_name: str = st.text_input("公司名称")

file_path = st.file_uploader('上传文件')



run: bool = st.button("开始处理", key='run')

if run:
    raw_df = get_raw_df(file_path)
    new_project = NewOccupationalHealthItemInfo(project_number, company_name, raw_df)
    st.dataframe(new_project.point_df)
    st.button('处理记录表',on_click=new_project.write_to_templates)
