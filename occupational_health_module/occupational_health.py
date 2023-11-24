'''重构，每个函数或者方法都可以直接出结果。考虑使用functools模块'''

from io import BytesIO
import math
import os
import re
from copy import deepcopy
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List, Tuple
from nptyping import DataFrame  # , Structure as S
import numpy as np
import pandas as pd
from pandas.api.types import CategoricalDtype
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

from occupational_health_module.other_infos import templates_info

class OccupationalHealthItemInfo():
    def __init__(
            self,
            company_name: str,
            project_number: str,
            point_info_df: DataFrame,
            personnel_info_df: DataFrame,
    ) -> None:
        self.company_name: str = company_name
        self.project_number: str = project_number
        self.templates_info: Dict = templates_info
        self.default_types_order: List[str] = ['空白', '定点', '个体']
        self.point_info_df: DataFrame = point_info_df
        self.personnel_info_df: DataFrame = personnel_info_df
        # [ ] 创建默认的保存路径
        # self.output_path: str = os.path.join(
        #     os.path.expanduser("~/Desktop"),
        #     f'{self.project_number}记录表'
        # )
        # [ ] 是否默认获得职业卫生所有检测因素的参考信息
        # [ ] 检测信息排序
        # [ ] 获得采样日程下每一天的检测信息
        # [ ] 获得采样日程总天数
        # [ ] 获得所有所有空气有害物质的检测因素，包含定点和个体
        # [ ] 获得一天的空气有害物质检测因素，包含定点和个体
        # [ ] 获得一天的空白样品编号
        # [ ] 处理单日的定点检测信息，为其加上样品编号范围和空白样品编号
        # [ ] 处理单日的个体检测信息，为其加上样品编号范围和空白样品编号
        # [ ] 获得爆炸后的定点样品编号
        # [ ] 整理定点和个体的样品统计信息
        # [ ] 获得样品统计df里的各个检测因素的保存时间
        # [ ] 获得分开的接触时间，使用十进制来计算
        # [ ] 将每日空白信息，定点编号，爆炸定点编号，个体编号和样品统计信息写入excel文件中
        # [ ] 将信息写入记录表模板中
        # [ ] 更新已占用编号数量
        # [ ] 列表的自定义排序
        # [ ] 获得空白、定点和个体的编号范围
        # [ ] 将编号范围转换为字符串
