from typing import Any, List
import pandas as pd
from nptyping import DataFrame
import streamlit as st

from occupational_health_module.new_occupational_health import NewOccupationalHealthItemInfo

st.set_page_config(layout="wide", initial_sidebar_state="auto")

# @st.cache_data
# def get_raw_df(file_path) -> DataFrame:
#     raw_df: DataFrame = pd.read_excel(file_path)
#     return raw_df

st.title("系统生成职业卫生编号和记录表生成")
st.markdown("输入系统生成的职业卫生项目的相应信息，会自动处理信息")
st.header("输入数据")

st.subheader("输入项目基本信息")

all_templates_categories: List[str] = [
    '定点有害物质',
    '个体有害物质',
    '个体噪声',
    '仪器直读因素',
    '样品流转单',
]

col1, col2 = st.columns(2)
with col1:
    project_number: str = st.text_input("项目编号")
    # output_path: str = st.text_input("记录表输出路径", help="如果路径为空，则在桌面创建文件夹存放")
    # exploded: bool = st.checkbox("是否分为多列")
with col2:
    company_name: str = st.text_input("公司名称")
    in_is_all_factors_split: bool = st.checkbox("是否所有检测因素按照采样日期分开", value=False)

in_templates_categories: List[str] = st.multiselect(
    '要写入的模板分类',
    options=all_templates_categories,
    default=all_templates_categories,
    help="仪器直读因素可写入的检测因素有**噪声**、**高温**、**照度**、**工频电场**和**一氧化碳**"
)

# file_path = st.file_uploader('上传文件')


st.subheader("输入样品信息")
raw_df: DataFrame[Any] = st.data_editor(
    pd.DataFrame([{
        'ID': None,
        '委托编号': None,
        '样品类型': None,
        '样品编号': None,
        '送样编号': None,
        '样品名称': None,
        '检测参数': None,
        '采样/送样日期': None,
        '收样日期': None,
        '样品描述': None,
        '样品状态': None,
        '代表时长/h': None,
        '单元': None,
        '工种/岗位': None,
        '检测地点': None,
        '测点编号': None,
        '第几天': None,
        '第几个频次': None,
        '采样方式': None,
        '作业人数': None,
        '日接触时长/h': None,
        '周工作天数/d': None,
    }]),
    num_rows="dynamic",
    use_container_width=False,
    key="info",
    column_config={
        "日接触时长/h": st.column_config.NumberColumn(format="%.2f"),
        "采样数量/天": st.column_config.NumberColumn(format="%d"),
        "第几天": st.column_config.NumberColumn(format="%d"),
        "第几个频次": st.column_config.NumberColumn(format="%d"),
        "测点编号": st.column_config.NumberColumn(format="%d"),
        '采样/送样日期': st.column_config.DateColumn(format='YYYY-MM-DD'),
        '样品编号': st.column_config.TextColumn(),
        "周工作天数/d": st.column_config.NumberColumn(format="%.2f"),
    }
)



run: bool = st.button("执行", key='run')

if run:
    # raw_df = get_raw_df(file_path)
    new_project: NewOccupationalHealthItemInfo = NewOccupationalHealthItemInfo(
        project_number,
        company_name,
        raw_df,
        is_all_factors_split=in_is_all_factors_split,
        in_templates_categories=in_templates_categories
    )
    st.header('处理得到的所有样品信息')
    st.dataframe(
        new_project.stat_df,
        column_config={
            '采样/送样日期': st.column_config.DateColumn(format='YYYY-MM-DD'),
        }
    )
    # st.dataframe(new_project.factor_reference_df)
    st.button('处理记录表', on_click=new_project.write_to_templates)
    # is_process: bool = st.button('处理记录表', key='process')
    # if is_process:
    #     try:
    #         new_project.write_to_templates()
    #         st.success(f"完成，已保存到{new_project.output_path}")
    #     except Exception:
    #         st.error('出现错误，无法进行')
