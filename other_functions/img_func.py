'''图片加水印函数'''

# %% 图片加水印

from typing import Any, Union, Optional
from io import BytesIO
from PIL import Image, ImageDraw ,ImageFont
import pandas as pd
from pandas.core.frame import DataFrame
#%% [markdown]

# ### 将多列信息合并成一列

#%%
def merge_multi_cols_info(_file_path) -> DataFrame:
    '''
    目的:
        一个excel表格文件中除第一列是文件名之外，其他列都是其他文本信息。
        本函数将其他列的信息合并成包含换行符的一列，用于为图片添加水印
    参数:
        _file_path: excel表格文件的路径
    返回:
        包含图片文件名和将为其添加的水印文本的Dataframe
    '''
    # 本字典用于指定文件各列的格式，不需要时可隐藏本字典
    _dtype_dict: dict = {
        "文件名": str,
        "经度": float,
        "纬度": float,
        "地址": str,
        # "时间": np.datetime64,
        "备注": str
    }

    _wm_df: DataFrame = pd.read_excel(
        _file_path,
        dtype=_dtype_dict  # type: ignore # 不需要时可隐藏
    )

    # 将经度和纬度的精度保留到小数点后六位，不需要时可隐藏
    _wm_df["经度"] = _wm_df["经度"].apply(lambda x: format(x, ".6f"))
    _wm_df["纬度"] = _wm_df["纬度"].apply(lambda x: format(x, ".6f"))
    _info_cols: list = [col for col in _wm_df.columns.tolist() if col != "文件名"]
    for _col in _info_cols:
        _wm_df[_col] = f"{_col}: " + _wm_df[_col].astype(str)

    _wm_df["水印信息"] = _wm_df[_info_cols].astype(str).apply("\n".join, axis=1)
    return _wm_df[["文件名", "水印信息"]]

# %% [markdown]

# ### 处理水印信息df

# %%
def handle_waterprint_info_df(_in_wp_df: DataFrame) -> DataFrame:
    '''
    目的:
        得到处理过的水印信息df
    参数:
        _in_wp_df: 水印信息df
    返回:
        处理过的水印信息df
    '''
    _dtype_dict: dict = {
        "文件名": str,
        "经度": float,
        "纬度": float,
        "地址": str,
        # "时间": np.datetime64,
        "备注": str
    }
    _wp_df: DataFrame = _in_wp_df.astype(_dtype_dict)
    _wp_df["经度"] = _wp_df["经度"].apply(lambda x: format(x, ".6f"))
    _wp_df["纬度"] = _wp_df["纬度"].apply(lambda x: format(x, ".6f"))

    _info_cols: list = [col for col in _wp_df.columns.tolist() if col != "文件名"]
    for _col in _info_cols:
        _wp_df[_col] = f"{_col}: " + _wp_df[_col].astype(str)
    
    _wp_df["水印信息"] = _wp_df[_info_cols].astype(str).apply("\n".join, axis=1)
    return _wp_df[["文件名", "水印信息"]]

#%% [markdown]

# ### 向图片增加水印，保存到bytes中

#%%
def add_waterprint_to_bytes(
    _img_file:     Any,
    _x:            Union[int, float],
    _y:            Union[int, float],
    _wm_text:      str,
    _font_family:  Optional[str],
    _font_size:    int,
    _fill:         str,
    _spacing:      int,
    # _direction:    Optional[str],
    _align:        str,
    _stroke_fill:  str,
    _stroke_width: int
) -> BytesIO:
    '''
    目的:
        向图片增加水印，并保存到bytes中
    参数:
        _img_file:     图片路径
        _x:            水印在图片的x轴位置
        _y:            水印在图片的y轴位置
        _wm_text:      水印文本内容
        _font_family:  水印字体的类型
        _font_size:    水印字体的大小
        _fill:         水印字体的颜色
        _spacing:      水印文本的行间距
        _direction:    水印文本的方向，需安装libraqm库才可以使用
        _align:        水印文本的对齐
        _stroke_fill:  水印文本的描边颜色
        _stroke_width: 水印文本的描边宽度
    返回:
        存放在bytes中的带水印的图片
    '''
    _imgbyte: BytesIO = BytesIO()
    _img: Any = Image.open(_img_file)
    _draw: Any = ImageDraw.Draw(_img)
    _font: Any = ImageFont.truetype(font=_font_family, size=_font_size)
    _draw.text(
        xy=(_x, _y),
        text=_wm_text,
        fill=_fill,
        font=_font,
        spacing=_spacing,
        # direction=_direction, # 文本方向需安装libraqm库才可以使用
        align=_align,
        stroke_fill=_stroke_fill,
        stroke_width=_stroke_width
    )
    _img.save(_imgbyte, format="jpeg")

    return _imgbyte
#%% [markdown]

# ### 向图片增加水印，保存到本地文件夹中

#%%
def add_waterprint_to_local(
    _img_name:            Any,
    _target_files_folder: Optional[str],
    _output_folder:       Optional[str],
    _x:                   Union[int, float],
    _y:                   Union[int, float],
    _wm_text:             str,
    _font_family:         Optional[str],
    _font_size:           int,
    _fill:                str,
    _spacing:             int,
    # _direction:           Optional[str],
    _align:               str,
    _stroke_fill:         str,
    _stroke_width:        int
) -> None:
    '''
    目的:
        向图片增加水印，并保存到本地文件夹中
    参数:
        _img_name:            图片文件名
        _target_files_folder: 目标图片文件所在的路径,
        _output_folder:       输出文件夹的路径,
        _x:                   水印在图片的x轴位置
        _y:                   水印在图片的y轴位置
        _wm_text:             水印文本内容
        _font_family:         水印字体的类型
        _font_size:           水印字体的大小
        _fill:                水印字体的颜色
        _spacing:             水印文本的行间距
        _direction:           水印文本的方向，需安装libraqm库才可以使用
        _align:               水印文本的对齐
        _stroke_fill:         水印文本的描边颜色
        _stroke_width:        水印文本的描边宽度
    返回:
      None，存放在本地文件夹中的带水印的图片
    '''
    _img: Any = Image.open(f"{_target_files_folder}\\{_img_name}")
    _draw: Any = ImageDraw.Draw(_img)
    _font: Any = ImageFont.truetype(font=_font_family, size=_font_size)
    _draw.text(
        xy=(_x, _y),
        text=_wm_text,
        fill=_fill,
        font=_font,
        spacing=_spacing,
        # direction=_direction, # 文本方向需安装libraqm库才可以使用
        align=_align,
        stroke_fill=_stroke_fill,
        stroke_width=_stroke_width
    )
    _img.save(f"{_output_folder}\\wm_{_img_name}")
