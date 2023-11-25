'''
重构，每个函数或者方法都可以直接出结果。考虑使用functools模块
'''

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

# from occupational_health_module.other_infos import templates_info


class OccupationalHealthItemInfo():
    '''职业卫生相应信息生成'''

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
        self.output_path: str = os.path.join(
            os.path.expanduser("~/Desktop"),
            f'{self.project_number}记录表'
        )
        # [ ] 数据预先操作方法
        self.factor_reference_df: DataFrame = self.get_occupational_health_factor_reference()
        self.sort_df()
        self.get_detection_days()
        self.schedule_days: int = self.point_info_df['采样日程'].max()  # 采样日程总天数
        (
            self.point_deleterious_substance_df,
            self.personnel_deleterious_substance_df
        ) = self.get_deleterious_substance_df()

    # [x] 创建默认的保存路径
    def create_normal_folder(self) -> None:
        '''创建默认的保存路径'''
        if not os.path.exists(self.output_path):
            os.mkdir(self.output_path)
        else:
            pass

    # [x] 默认获得职业卫生所有检测因素的参考信息
    def get_occupational_health_factor_reference(self) -> DataFrame:
        '''
        获得职业卫生所有检测因素的参考信息
        '''
        # reference_path: str = './info_files/检测因素参考信息.xlsx'
        # reference_df: DataFrame = pd.read_excel(reference_path)  # type: ignore
        reference_path: str = './info_files/检测因素参考信息.csv'
        reference_df: DataFrame = pd.read_csv(reference_path)  # type: ignore
        return reference_df

    # [x] 检测信息排序
    def sort_df(self) -> None:
        '''
        对检测信息里的检测因素进行排序
        先对复合检测因素内部排序，再对所有检测因素
        '''
        # 将检测因素，尤其是复合检测因素分开转换为列表，并重新排序
        factor_reference_list: List[str] = (
            self.factor_reference_df['标识检测因素']
            .tolist()
        )
        self.point_info_df['检测因素'] = (
            self
            .point_info_df['检测因素']
            .str.split('|')
            # type: ignore
            .apply(self.custom_sort, args=(factor_reference_list,))
        )
        self.personnel_info_df['检测因素'] = (
            self
            .personnel_info_df['检测因素']
            .str.split('|')
            # type: ignore
            .apply(self.custom_sort, args=(factor_reference_list,))
        )
        # 将检测因素列表的第一个作为标识
        self.point_info_df['标识检测因素'] = (
            self
            .point_info_df['检测因素']
            .apply(lambda lst: lst[0])  # type: ignore
        )
        self.personnel_info_df['标识检测因素'] = (
            self
            .personnel_info_df['检测因素']
            .apply(lambda lst: lst[0])  # type: ignore
        )
        # 将检测因素列表合并为字符串
        self.point_info_df['检测因素'] = (
            self
            .point_info_df['检测因素']
            .apply(lambda x: '|'.join(x))   # type: ignore
        )
        self.personnel_info_df['检测因素'] = (
            self
            .personnel_info_df['检测因素']
            .apply(lambda x: '|'.join(x))   # type: ignore
        )
        # 将定点和个体的检测因素提取出来，创建CategoricalDtype数据，按照拼音排序
        point_factor_list: List[str] = (
            self
            .point_info_df['检测因素']
            .unique().tolist()  # type: ignore
        )
        point_factor_list: List[str] = sorted(
            point_factor_list,
            key=lambda x: x.encode('gbk')
        )
        point_factor_order: CategoricalDtype = CategoricalDtype(
            point_factor_list, ordered=True)
        personnel_factor_list: List[str] = (
            self
            .personnel_info_df['检测因素']
            .unique().tolist()  # type: ignore
        )
        personnel_factor_list: List[str] = sorted(
            personnel_factor_list,
            key=lambda x: x.encode('gbk')
        )
        personnel_factor_order: CategoricalDtype = CategoricalDtype(
            personnel_factor_list,
            ordered=True
        )
        # 将检测因素按照拼音排序
        self.point_info_df['检测因素'] = (
            self
            .point_info_df['检测因素']
            .astype(point_factor_order)  # type: ignore
        )
        self.point_info_df = (
            self
            .point_info_df
            # type: ignore
            .sort_values(by=['检测因素', '采样点编号'], ascending=True, ignore_index=True)
        )
        self.personnel_info_df['检测因素'] = (
            self
            .personnel_info_df['检测因素']
            .astype(personnel_factor_order)  # type: ignore
        )
        self.personnel_info_df = (
            self
            .personnel_info_df
            # type: ignore
            .sort_values(by=['检测因素', '采样点编号'], ascending=True, ignore_index=True)
        )

    # [x] 获得采样日程下每一天的检测信息，确定将采样日程改为“1|2|3”或者指定的“2”的方式
    def get_detection_days(self) -> None:
        '''
        获得采样日程下每一天的检测信息
        '''
        self.point_info_df['采样日程'] = (
            self
            .point_info_df['采样日程']
            .str.split('|')  # type: ignore
        )
        self.point_info_df = (
            self
            .point_info_df
            .explode('采样日程', ignore_index=True)
        )
        self.point_info_df['采样日程'] = (
            self
            .point_info_df['采样日程']
            .astype(int)  # type: ignore
        )
        self.personnel_info_df['采样日程'] = (
            self
            .personnel_info_df['采样日程']
            .str.split('|')  # type: ignore
        )
        self.personnel_info_df = (
            self
            .personnel_info_df
            .explode('采样日程', ignore_index=True)
        )
        self.personnel_info_df['采样日程'] = (
            self
            .personnel_info_df['采样日程']
            .astype(int)  # type: ignore
        )

    # [x] 获得所有所有空气有害物质的检测因素，包含定点和个体
    def get_deleterious_substance_df(self) -> Tuple[DataFrame, DataFrame]:
        '''
        获得所有空气有害物质的检测因素，包含定点和个体
        '''
        point_deleterious_substance_df: DataFrame = (
            pd.merge(  # type: ignore
                self.point_info_df,
                self.factor_reference_df[
                    ['标识检测因素', '是否仪器直读', '是否需要空白', '复合因素代码']
                ],
                on='标识检测因素',
                how='left'
            )
            .fillna({'是否需要空白': False, '复合因素代码': 0, '是否仪器直读': False})
            .query('是否仪器直读 == False')
        )
        personnel_deleterious_substance_df: DataFrame = (
            pd.merge(  # type: ignore
                self.personnel_info_df,
                self.factor_reference_df[
                    ['标识检测因素', '是否仪器直读', '是否需要空白', '复合因素代码']
                ],
                on='标识检测因素',
                how='left'
            )
            .fillna({'是否需要空白': False, '复合因素代码': 0, '是否仪器直读': False})
            .query('是否仪器直读 == False')
        )
        return point_deleterious_substance_df, personnel_deleterious_substance_df

    # [x] 获得一天的空气有害物质检测因素，包含定点和个体
    def get_single_day_deleterious_substance_df(
        self, schedule_day: int = 1
    ) -> Tuple[DataFrame, DataFrame]:
        '''
        获得一天的空气有害物质检测因素，包含定点和个体
        '''
        single_day_point_deleterious_substance_df: DataFrame = (
            self
            .point_deleterious_substance_df
            [self.point_deleterious_substance_df['采样日程'] == schedule_day]
        )
        single_day_personnel_deleterious_substance_df: DataFrame = (
            self
            .personnel_deleterious_substance_df
            [self.personnel_deleterious_substance_df['采样日程'] == schedule_day]
        )
        single_day_deleterious_substance_df_tuple: Tuple[DataFrame, DataFrame] = (
            single_day_point_deleterious_substance_df,
            single_day_personnel_deleterious_substance_df
        )
        return single_day_deleterious_substance_df_tuple
    # [x] 获得一天的空白样品编号

    def get_single_day_blank_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        获得一天的空白样品编号
        '''
        # 应对空白数量为0的情况
        # 复制定点和个体检测信息的dataframe，避免提示错误
        point_df, personnel_df = (
            self
            .get_single_day_deleterious_substance_df(schedule_day)
        )
        single_day_point_df: DataFrame = point_df.copy()
        single_day_personnel_df: DataFrame = personnel_df.copy()
        # 从定点和个体的dataframe提取检测因素，去重以及合并
        single_day_point_df['检测因素'] = (
            single_day_point_df['检测因素']
            .str.split('|')  # type: ignore
        )
        ex_single_day_point_df: DataFrame = single_day_point_df.explode('检测因素')
        single_day_personnel_df['检测因素'] = (
            single_day_personnel_df['检测因素']
            .str.split('|')  # type: ignore
        )
        ex_single_day_personnel_df: DataFrame = single_day_personnel_df.explode(
            '检测因素')

        # 筛选出需要空白的检测因素
        test_df: DataFrame = (
            pd.concat(
                [
                    ex_single_day_point_df[['检测因素', '是否需要空白', '复合因素代码']],
                    ex_single_day_personnel_df[['检测因素', '是否需要空白', '复合因素代码']]
                ]
            )
            .query('是否需要空白 == True')
            .drop_duplicates('检测因素')
            .reset_index(drop=True)
        )
        # 分别处理非复合因素和复合因素，复合因素要合并。
        # 判断定点和个体的检测因素是否为空
        if test_df.empty:
            single_day_blank_df = pd.DataFrame(columns=['标识检测因素', '空白编号'])
        else:
            raw_group1: DataFrame = test_df.loc[test_df['复合因素代码'] == 0]
            raw_group2: DataFrame = test_df.loc[test_df['复合因素代码'] != 0]
            if raw_group1.empty:
                group1 = pd.DataFrame(columns=['检测因素', '是否需要空白'])
            else:
                group1: DataFrame = raw_group1.loc[:, ['检测因素', '是否需要空白']]
            if raw_group2.empty:
                group2 = pd.DataFrame(columns=['检测因素', '是否需要空白'])
            else:
                group2: DataFrame = (
                    pd.DataFrame(raw_group2.groupby(['复合因素代码'], group_keys=False)['检测因素']
                                 .apply('|'.join))
                    .reset_index(drop=True)
                )
                group2.loc[:, '是否需要空白'] = True

            # 最后合并，排序
            concat_group: DataFrame = pd.concat(  # type: ignore
                [group1, group2],
                ignore_index=True,
                axis=0,
                sort=False
            )
            blank_factor_list: List[str] = sorted(
                concat_group['检测因素'].tolist(),
                key=lambda x: x.encode('gbk')
            )  # type: ignore
            blank_factor_order: CategoricalDtype = CategoricalDtype(
                categories=blank_factor_list,
                ordered=True
            )
            concat_group['检测因素'] = (
                concat_group['检测因素']
                .astype(blank_factor_order)  # type: ignore
            )
            # 筛选出需要空白编号的检测因素，并赋值
            single_day_blank_df: DataFrame = (
                concat_group
                # 必须用`==`才可用，按照提示用`is`会失败
                .loc[concat_group['是否需要空白'] == True]
                .sort_values('检测因素', ignore_index=True)
            )  # type: ignore
            single_day_blank_df['检测因素'] = (
                single_day_blank_df['检测因素']
                .astype(str)
                # type: ignore
                .map(lambda x: [x] + x.split('|') if x.count('|') > 0 else x)
            )
            single_day_blank_df['空白编号'] = (
                np.arange(1, single_day_blank_df.shape[0] + 1)
                + engaged_num  # type: ignore
            )
            single_day_blank_df.drop(
                columns=['是否需要空白'],
                inplace=True
            )  # type: ignore
            single_day_blank_df = (
                single_day_blank_df
                .explode('检测因素')
                .rename(columns={'检测因素': '标识检测因素'})
            )
        return single_day_blank_df
    # [x] 处理单日的定点检测信息，为其加上样品编号范围和空白样品编号

    def get_single_day_point_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        处理单日的定点检测信息，为其加上样品编号范围和空白样品编号
        '''
        # 注：为定点添加空白编号的功能不要放到这里实现
        point_df: DataFrame = (
            self
            .get_single_day_deleterious_substance_df(schedule_day)[0]
            .copy()
        )
        # [ ] 如果数量为0
        point_df['终止编号'] = (
            point_df['采样数量/天'].cumsum()
            + engaged_num  # type: ignore
        )
        point_df['起始编号'] = point_df['终止编号'] - point_df['采样数量/天'] + 1

        return point_df
    # [x] 处理单日的个体检测信息，为其加上样品编号范围和空白样品编号

    def get_single_day_personnel_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        处理单日的个体检测信息，为其加上样品编号范围和空白样品编号
        '''
        personnel_df: DataFrame = (
            self
            .get_single_day_deleterious_substance_df(schedule_day)
            [1].copy()
        )
        # [ ] 如果数量为0
        personnel_df['个体编号'] = (
            personnel_df['采样数量/天'].cumsum()
            + engaged_num  # type: ignore
        )
        return personnel_df
    # [ ] 获得单日的所有编号排列好的样品信息

    def get_single_day_dfs(self, engaged_num: int = 0, schedule_day: int = 1): # type: ignore
        '''为单日的监测信息的样品编号'''
        for type_order in self.default_types_order:
            if type_order == '空白':
                current_blank_df: DataFrame = (
                    self.get_single_day_blank_df(engaged_num, schedule_day)
                )
                engaged_num: int = (
                    self.refresh_engaged_num(
                        current_blank_df,
                        type_order,
                        engaged_num
                    )
                )
            elif type_order == '定点':
                current_point_df: DataFrame = (
                    self.get_single_day_point_df(engaged_num, schedule_day)
                )
                engaged_num: int = (
                    self.refresh_engaged_num(
                        current_point_df,
                        type_order,
                        engaged_num
                    )
                )
            elif type_order == '个体':
                current_personnel_df: DataFrame = (
                    self.get_single_day_personnel_df(engaged_num, schedule_day)
                )
                engaged_num: int = (
                    self.refresh_engaged_num(
                        current_personnel_df,
                        type_order,
                        engaged_num
                    )
                )
        # [ ] 为个体和定点添加空白编号
    # [ ] 获得爆炸后的定点样品编号
    # [ ] 整理定点和个体的样品统计信息
    # [ ] 获得样品统计df里的各个检测因素的保存时间
    # [ ] 获得分开的接触时间，使用十进制来计算
    # [ ] 将每日空白信息，定点编号，爆炸定点编号，个体编号和样品统计信息写入excel文件中
    # [ ] 将信息写入记录表模板中
    # [x] 更新已占用编号数量

    def refresh_engaged_num(
        self,
        current_df: DataFrame,
        current_type: str,
        engaged_num: int
    ) -> int:
        '''更新已占用样品编号数'''
        # 按照df类型来更新编号
        # [x] 如果df长度为0时要
        # [x] 更新，使用字典模式
        type_num_dict: Dict[str, str] = {
            '空白': '空白编号',
            '定点': '终止编号',
            '个体': '个体编号',
        }
        if current_df.shape[0] != 0 and type in self.default_types_order:
            new_engaged_num: int = (
                current_df[type_num_dict[current_type]]
                .astype(int)
                .max()
            )
            return new_engaged_num
        else:
            return engaged_num

    # [x] 列表的自定义排序

    def custom_sort(self, str_list: List[str], key_list: List[str]) -> List[str]:
        '''
        列表的自定义排序
        '''
        if str_list[0] in key_list:
            sorted_str_list: List[str] = sorted(
                str_list,
                key=lambda x: key_list.index(x)
            )
            return sorted_str_list
        else:
            return str_list
    # [ ] 获得空白、定点和个体的编号范围
    # [ ] 将编号范围转换为字符串
