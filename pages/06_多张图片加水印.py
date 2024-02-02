# %%

from typing import Union, Any, Optional, Tuple
import pandas as pd
from pandas.core.frame import DataFrame

import streamlit as st
from other_functions.options_dict import _fonts_dict, _align_dict
from other_functions.img_func import add_waterprint_to_local, handle_waterprint_info_df
# %%
st.set_page_config(layout="wide", initial_sidebar_state="auto")
Number = Union[int, float]

# %%
st.header("多张图片加水印")
st.markdown("用于向多张图片添加水印。")
with st.form("输入信息"):
    st.subheader("目标图片位置和保存位置")
    _target_folder_col, _output_folder_col = st.columns(2)
    with _target_folder_col:
        _target_folder_path: str = st.text_input("输入要处理的图片所在路径")
    with _output_folder_col:
        _output_folder_path: str = st.text_input("输入存放处理完成的图片所在路径")
    
    st.subheader("水印信息")
    _in_waterprint_df = st.data_editor(
        pd.DataFrame([{
            "文件名":' 111.jpg',
            "经度": '119.338175',
            "纬度": '25.907093',
            "地址": '福建省福州市闽侯县辅翼村辅前路10号',
            # "时间": pd.to_datetime("2023-01-01"),
            "时间": pd.to_datetime("2023-01-01 00:00:00"),
            "备注": '永正生态'
        }]),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "经度": st.column_config.NumberColumn(format="%.6f"),
            "纬度": st.column_config.NumberColumn(format="%.6f"),
            # "时间": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "时间": st.column_config.DatetimeColumn(format="YYYY-MM-DD hh:mm:ss"),
        }
    )

    st.subheader("输入水印位置")
    # 双列布局
    _x_col, _y_col = st.columns(2)
    with _x_col:
        _x: Number = st.number_input("x位置", value=5, step=1, format="%d")
    with _y_col:
        _y: Number = st.number_input("y位置", value=5, step=1, format="%d")
    st.subheader("文本样式")
    style_col1, style_col2, style_col3 = st.columns(3)
    with style_col1:
        _font_family: Optional[str] = st.selectbox(
            "字体",
            list(_fonts_dict.keys()),
            index=0
        )
        _spacing: Number = st.number_input(
            "间隔",
            value=4,
            step=1,
            format="%d"
        )
    with style_col2:
        _font_size: Number = st.number_input(
            "字体大小",
            value=50,
            step=1,
            format="%d"
        )
        _stroke_width: Number = st.number_input(
            "描边宽度大小",
            value=0,
            step=1,
            format="%d"
        )
    with style_col3:
        _fill: str = st.color_picker("选择颜色", value="#000000")
        _stroke_fill: str = st.color_picker("选择描边颜色", value="#000000")
    _align: Union[str, Tuple[str, str], None] = st.select_slider(
        "对齐",
        list(_align_dict.keys()),
        value=list(_align_dict.keys())[1]
    )
    _submited: bool = st.form_submit_button("执行")

# %%
if _submited:
    # _watermarker_df: DataFrame = merge_multi_cols_info(_watermarker_file)
    _watermarker_df: DataFrame = handle_waterprint_info_df(_in_waterprint_df)
    
    for _img_name, _wm_text in zip(_watermarker_df["文件名"], _watermarker_df["水印信息"]):
        add_waterprint_to_local(
            _img_name,
            _target_folder_path,
            _output_folder_path,
            _x,
            _y,
            _wm_text,
            _fonts_dict[_font_family],  # type: ignore
            int(_font_size),
            _fill,
            int(_spacing),
            _align_dict[_align],  # type: ignore
            _stroke_fill,
            int(_stroke_width)
        )
else:
    st.info("请输入信息。")
