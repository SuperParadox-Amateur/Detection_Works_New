# %%

from io import BytesIO
from typing import Union, Any, Optional, Tuple

import streamlit as st
from functions.options_dict import _fonts_dict, _align_dict
from functions.img_func import add_waterprint_to_bytes
# %%
st.set_page_config(layout="wide", initial_sidebar_state="auto")
Number = Union[int, float]
# %%


# def add_waterprint(
#     _img_file: Any,
#     _x: Union[int, float],
#     _y: Union[int, float],
#     _text: str,
#     _font_family: Optional[str],
#     _font_size: int,
#     _fill: str,
#     _spacing: int,
#     # _direction: Optional[str], # 文本方向需安装libraqm库才可以使用
#     _align: str,
#     _stroke_fill: str,
#     _stroke_width: int
# ) -> BytesIO:
#     _imgbyte: BytesIO = BytesIO()
#     _img: Any = Image.open(_img_file)
#     _draw: Any = ImageDraw.Draw(_img)
#     _font: Any = ImageFont.truetype(font=_font_family, size=_font_size)
#     _draw.text(
#         xy=(_x, _y),
#         text=_text,
#         fill=_fill,
#         font=_font,
#         spacing=_spacing,
#         # direction=_direction, # 文本方向需安装libraqm库才可以使用
#         align=_align,
#         stroke_fill=_stroke_fill,
#         stroke_width=_stroke_width
#     )
#     _img.save(_imgbyte, format="jpeg")
#     # _imgbytevalue: bytes = _imgbyte.getvalue()

#     # return _imgbytevalue
#     return _imgbyte

# %%

st.header("单张图片加水印")
st.markdown("用于向单张图片添加水印。")
with st.form("输入信息"):
    # _img_path: str = st.text_input("输入图片路径")
    st.subheader("选择图片")
    _img_file: Any = st.file_uploader(
        "选择图片文件",
        accept_multiple_files=False
    )
    st.subheader("输入水印位置")
    # 双列布局
    _x_col, _y_col = st.columns(2)
    with _x_col:
        _x: Number = st.number_input("x位置", value=5, step=1, format="%d")
    with _y_col:
        _y: Number = st.number_input("y位置", value=5, step=1, format="%d")
    # _x: Number = st.number_input("x位置", value=5, step=1, format="%d") #单列布局
    # _y: Number = st.number_input("y位置", value=5, step=1, format="%d") #单列布局
    _text: str = st.text_area("输入水印名称", value="水印测试")
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
if _submited:
    _img_byte: BytesIO = add_waterprint_to_bytes(
        _img_file,
        _x,
        _y,
        _text,
        _fonts_dict[_font_family],  # type: ignore
        int(_font_size),
        _fill,
        int(_spacing),
        # _directions_dict[_direction],
        _align_dict[_align],  # type: ignore
        _stroke_fill,
        int(_stroke_width)
    )
    st.download_button("下载",
                       data=_img_byte.getvalue(),
                       file_name=f"wm_{_img_file.name}"
                       )
    st.text(f"{_img_file.name}")
    st.image(_img_byte)
else:
    st.info("请输入信息。")
