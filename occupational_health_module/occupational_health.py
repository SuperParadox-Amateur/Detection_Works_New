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


# point_df_dtype: Dict[str, type[int] | type[str] | type[float]] = {
#         '采样点编号': int,
#         '单元': str,
#         '检测地点': str,
#         '工种': str,
#         '日接触时间': float,
#         '检测因素': str,
#         '采样数量/天': int,
#         '采样天数': int,
#         }

# pernonnel_df_dtype: Dict[str, type[int] | type[str] | type[float]] = {
#     '采样点编号': int,
#     '单元': str,
#     '工种': str,
#     '日接触时间': float,
#     '检测因素': str,
#     '采样数量/天': int,
#     '采样天数': int,
# }


# [x] 考虑将采样日程改为“1|2|3”或者指定的“2”的方式，这样可以自定义部分只需要一天样品的检测信息的采样日程
# 问题：采样可能是1天的定期或者3天的评价，产生的存储dataframe的变量要如何命名，最后要如何展示在streamlit的多标签里

class OccupationalHealthItemInfo():
    def __init__(
            # [x] 计划将项目基本信息以dict的形式存放
            self,
            company_name: str,
            project_number: str,
            # blank_info_df: DataFrame,
            point_info_df: DataFrame,
            personnel_info_df: DataFrame,
            # templates_info: Dict
    ) -> None:
        self.company_name: str = company_name
        self.project_number: str = project_number
        self.templates_info: Dict = templates_info
        self.output_path: str = os.path.join(
            os.path.expanduser("~/Desktop"),
            f'{self.project_number}记录表'
        )
        # self.point_output_path = os.path.join(self.output_path, '记录表', '定点')
        # self.personnel_output_path = os.path.join(self.output_path, '记录表', '个体')
        self.default_types_order: List[str] = ['空白', '定点', '个体']
        self.point_info_df: DataFrame = point_info_df
        self.personnel_info_df: DataFrame = personnel_info_df
        self.factor_reference_df: DataFrame = self.get_occupational_health_factor_reference()
        # self.point_factor_order, self.personnel_factor_order = self.get_point_personnel_factors_order()  # 应该不需要
        self.sort_df()
        self.get_detection_days()
        # type: ignore
        self.schedule_days: int = self.point_info_df['采样日程'].max()
        (
            self.point_deleterious_substance_df,
            self.personnel_deleterious_substance_df
        ) = self.get_deleterious_substance_df()
        # self.dfs_num: BytesIO = self.get_dfs_num(
        #     types_order=self.default_types_order
        # )
        # self.single_day_func_map = {
        #     '空白': self.get_single_day_blank_df,
        #     '定点': self.get_single_day_point_df,
        #     '个体': self.get_single_day_personnel_df,
        # }
        # self.df_name_map: dict[str, str] = {
        #     '空白': 'blank',
        #     '定点': 'point',
        #     '个体': 'personnel',
        # }


    # def get_point_personnel_factors_order(self) -> Tuple[CategoricalDtype, CategoricalDtype]:
    #     '''
    #     （已废弃）将定点和个体检测信息里的检测因素按照汉字拼音排序，并导出CategoricalDtype
    #     '''
    #     point_factor_list: List[str] = self.point_info_df['检测因素'].unique().tolist()  # type: ignore
    #     point_factor_list: List[str] = sorted(point_factor_list, key=lambda x: x.encode('gbk'))
    #     point_factor_order = CategoricalDtype(point_factor_list, ordered=True)
    #     personnel_factor_list: List[str] = self.personnel_info_df['检测因素'].unique().tolist()  # type: ignore
    #     personnel_factor_list: List[str] = sorted(personnel_factor_list, key=lambda x: x.encode('gbk'))
    #     personnel_factor_order = CategoricalDtype(personnel_factor_list, ordered=True)
    #     return point_factor_order, personnel_factor_order

    def get_occupational_health_factor_reference(self) -> DataFrame:
        '''
        获得职业卫生所有检测因素的参考信息
        '''
        # reference_path: str = './info_files/检测因素参考信息.xlsx'
        # reference_df: DataFrame = pd.read_excel(reference_path)  # type: ignore
        reference_path: str = './info_files/检测因素参考信息.csv'
        reference_df: DataFrame = pd.read_csv(reference_path)  # type: ignore
        return reference_df

    def get_detection_days(self) -> None:
        '''
        获得采样日程下每一天的检测信息
        '''
        self.point_info_df['采样日程'] = self.point_info_df['采样天数'].apply(
            lambda x: list(range(1, x + 1)))  # type: ignore
        self.point_info_df = self.point_info_df.explode(
            '采样日程', ignore_index=True)
        self.point_info_df['采样日程'] = self.point_info_df['采样日程'].astype(
            int)  # type: ignore
        self.personnel_info_df['采样日程'] = self.personnel_info_df['采样天数'].apply(
            lambda x: list(range(1, x + 1)))  # type: ignore
        self.personnel_info_df = self.personnel_info_df.explode(
            '采样日程', ignore_index=True)
        self.personnel_info_df['采样日程'] = self.personnel_info_df['采样日程'].astype(
            int)  # type: ignore

    def sort_df(self) -> None:
        '''
        对检测信息里的检测因素进行排序
        先对复合检测因素内部排序，再对所有检测因素
        '''
        # 将检测因素，尤其是复合检测因素分开转换为列表，并重新排序
        factor_reference_list: List[str] = self.factor_reference_df['标识检测因素'].tolist(
        )
        self.point_info_df['检测因素'] = self.point_info_df['检测因素'].str.split('|').apply(
            self.custom_sort, args=(factor_reference_list,))  # type: ignore
        self.personnel_info_df['检测因素'] = self.personnel_info_df['检测因素'].str.split(
            '|').apply(self.custom_sort, args=(factor_reference_list,))  # type: ignore
        # 将检测因素列表的第一个作为标识
        self.point_info_df['标识检测因素'] = self.point_info_df['检测因素'].apply(
            lambda lst: lst[0])  # type: ignore
        self.personnel_info_df['标识检测因素'] = self.personnel_info_df['检测因素'].apply(
            lambda lst: lst[0])  # type: ignore
        # 将检测因素列表合并为字符串
        self.point_info_df['检测因素'] = self.point_info_df['检测因素'].apply(
            lambda x: '|'.join(x))   # type: ignore
        self.personnel_info_df['检测因素'] = self.personnel_info_df['检测因素'].apply(
            lambda x: '|'.join(x))   # type: ignore
        # 将定点和个体的检测因素提取出来，创建CategoricalDtype数据，按照拼音排序
        point_factor_list: List[str] = self.point_info_df['检测因素'].unique(
        ).tolist()  # type: ignore
        point_factor_list: List[str] = sorted(
            point_factor_list, key=lambda x: x.encode('gbk'))
        point_factor_order = CategoricalDtype(point_factor_list, ordered=True)
        personnel_factor_list: List[str] = self.personnel_info_df['检测因素'].unique(
        ).tolist()  # type: ignore
        personnel_factor_list: List[str] = sorted(
            personnel_factor_list, key=lambda x: x.encode('gbk'))
        personnel_factor_order = CategoricalDtype(
            personnel_factor_list, ordered=True)
        # 将检测因素按照拼音排序
        self.point_info_df['检测因素'] = self.point_info_df['检测因素'].astype(
            point_factor_order)  # type: ignore
        self.point_info_df = self.point_info_df.sort_values(
            by=['检测因素', '采样点编号'], ascending=True, ignore_index=True)  # type: ignore
        self.personnel_info_df['检测因素'] = self.personnel_info_df['检测因素'].astype(
            personnel_factor_order)  # type: ignore
        self.personnel_info_df = self.personnel_info_df.sort_values(
            by=['检测因素', '采样点编号'], ascending=True, ignore_index=True)  # type: ignore

    def get_deleterious_substance_df(self) -> Tuple[DataFrame, DataFrame]:
        '''
        获得所有空气有害物质的检测因素，包含定点和个体
        '''
        # （已废除）将参考信息里的所有空气有害物质检测因素转换为列表
        # deleterious_substance_factor_df: DataFrame = self.factor_reference_df.loc[self.factor_reference_df['收集方式'] != '直读']
        # deleterious_substance_factor_list: List[str] = deleterious_substance_factor_df['标识检测因素'].tolist()
        # （已废除）筛选出定点和个体检测信息里的含有所有空气有害物质检测因素的检测信息
        # point_deleterious_substance_df: DataFrame = self.point_info_df[self.point_info_df['标识检测因素'].isin(deleterious_substance_factor_list)]  # type: ignore
        # personnel_deleterious_substance_df: DataFrame = self.personnel_info_df[self.personnel_info_df['标识检测因素'].isin(deleterious_substance_factor_list)]  # type: ignore
        point_deleterious_substance_df: DataFrame = (
            pd.merge(  # type: ignore
                self.point_info_df,
                self.factor_reference_df[[
                    '标识检测因素', '是否仪器直读', '是否需要空白', '复合因素代码']],
                on='标识检测因素',
                how='left'
            )
            .fillna({'是否需要空白': False, '复合因素代码': 0, '是否仪器直读': False})
            .query('是否仪器直读 == False')
        )
        personnel_deleterious_substance_df: DataFrame = (
            pd.merge(  # type: ignore
                self.personnel_info_df,
                self.factor_reference_df[[
                    '标识检测因素', '是否仪器直读', '是否需要空白', '复合因素代码']],
                on='标识检测因素',
                how='left'
            )
            .fillna({'是否需要空白': False, '复合因素代码': 0, '是否仪器直读': False})
            .query('是否仪器直读 == False')
        )
        return point_deleterious_substance_df, personnel_deleterious_substance_df

    def get_single_day_deleterious_substance_df(self, schedule_day: int = 1) -> Tuple[DataFrame, DataFrame]:
        '''
        获得一天的空气有害物质检测因素，包含定点和个体
        '''
        single_day_point_deleterious_substance_df: DataFrame = self.point_deleterious_substance_df[
            self.point_deleterious_substance_df['采样日程'] == schedule_day]
        single_day_personnel_deleterious_substance_df: DataFrame = self.personnel_deleterious_substance_df[
            self.personnel_deleterious_substance_df['采样日程'] == schedule_day]
        return single_day_point_deleterious_substance_df, single_day_personnel_deleterious_substance_df

    def get_single_day_blank_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        获得一天的空白样品编号
        '''
        # 应对空白数量为0的情况
        # 复制定点和个体检测信息的dataframe，避免提示错误
        point_df, personnel_df = self.get_single_day_deleterious_substance_df(
            schedule_day)
        single_day_point_df: DataFrame = point_df.copy()
        single_day_personnel_df: DataFrame = personnel_df.copy()
        # 从定点和个体的dataframe提取检测因素，去重以及合并
        single_day_point_df['检测因素'] = single_day_point_df['检测因素'].str.split(
            '|')  # type: ignore
        ex_single_day_point_df: DataFrame = single_day_point_df.explode('检测因素')
        single_day_personnel_df['检测因素'] = single_day_personnel_df['检测因素'].str.split(
            '|')  # type: ignore
        ex_single_day_personnel_df: DataFrame = single_day_personnel_df.explode(
            '检测因素')
        # test_df: DataFrame = pd.concat(  # type: ignore
        #     [
        #         ex_single_day_point_df[['检测因素', '是否需要空白', '复合因素代码']],
        #         ex_single_day_personnel_df[['检测因素', '是否需要空白', '复合因素代码']]
        #     ],
        #     ignore_index=True
        # ).drop_duplicates('检测因素')#.reset_index(drop=True)

        # 筛选出需要空白的检测因素
        test_df = (
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
                group2['是否需要空白'] = True

        # group1: DataFrame = test_df.loc[test_df['复合因素代码'] == 0, ['检测因素', '是否需要空白']]
        # raw_group2: DataFrame = test_df.loc[test_df['复合因素代码'] != 0]
        # group2 = pd.DataFrame(raw_group2.groupby(['复合因素代码'], group_keys=False)['检测因素'].apply('|'.join)).reset_index(drop=True)  # type: ignore
        # group2['是否需要空白'] = True
        # 最后合并，排序
            concat_group: DataFrame = pd.concat(  # type: ignore
                [group1, group2],
                ignore_index=True,
                axis=0,
                sort=False
            )
            blank_factor_list: List[str] = sorted(
                concat_group['检测因素'].tolist(), key=lambda x: x.encode('gbk'))  # type: ignore
            blank_factor_order = CategoricalDtype(
                categories=blank_factor_list, ordered=True)
            concat_group['检测因素'] = concat_group['检测因素'].astype(
                blank_factor_order)  # type: ignore
            # 筛选出需要空白编号的检测因素，并赋值
            single_day_blank_df: DataFrame = (
                concat_group.loc[concat_group['是否需要空白'] == True]
                .sort_values('检测因素', ignore_index=True)
            )  # type: ignore
            # 另起一列，用来放置标识检测项目
            # single_day_blank_df['标识检测因素'] = single_day_blank_df['检测因素'].astype(str).map(lambda x: x.split('|'))  # type: ignore
            single_day_blank_df['检测因素'] = single_day_blank_df['检测因素'].astype(str).map(
                lambda x: [x] + x.split('|') if x.count('|') > 0 else x)  # type: ignore
            single_day_blank_df['空白编号'] = np.arange(
                1, single_day_blank_df.shape[0] + 1) + engaged_num  # type: ignore
            single_day_blank_df.drop(
                columns=['是否需要空白'], inplace=True)  # type: ignore
            single_day_blank_df = single_day_blank_df.explode(
                '检测因素').rename(columns={'检测因素': '标识检测因素'})
        # single_day_blank_df = single_day_blank_df.explode('标识检测因素')
        return single_day_blank_df

    def get_single_day_point_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        处理单日的定点检测信息，为其加上样品编号范围和空白样品编号
        '''
        # 注：为定点添加空白编号的功能不要放到这里实现
        # blank_df: DataFrame = self.get_single_day_blank_df(engaged_num, schedule_day)
        point_df: DataFrame = self.get_single_day_deleterious_substance_df(schedule_day)[
            0].copy()
        point_df['终止编号'] = point_df['采样数量/天'].cumsum() + \
            engaged_num  # type: ignore
        point_df['起始编号'] = point_df['终止编号'] - point_df['采样数量/天'] + 1
        # r_point_df: DataFrame = pd.merge(point_df, blank_df, how='left', on=['标识检测因素']).fillna(0)  # type: ignore
        # [x] 可能加上完全的对应空白完全检测因素
        return point_df

    def get_single_day_personnel_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        处理单日的个体检测信息，为其加上样品编号范围和空白样品编号
        '''
        # blank_df: DataFrame = self.get_single_day_blank_df(engaged_num, schedule_day)
        personnel_df = self.get_single_day_deleterious_substance_df(schedule_day)[
            1].copy()
        personnel_df['个体编号'] = personnel_df['采样数量/天'].cumsum() + \
            engaged_num  # type: ignore
        # [x] 可能加上完全的对应空白完全检测因素
        # personnel_df['起始编号'] = personnel_df['终止编号'] - personnel_df['采样数量/天'] + 1
        # r_personnel_df: DataFrame = pd.merge(personnel_df, blank_df, how='left', on=['标识检测因素'])#.fillna(0)  # type: ignore
        # r_personnel_df['空白编号'] = r_personnel_df['空白编号'].astype('int')  # type: ignore
        return personnel_df

    def trim_dfs(self, current_point_df: DataFrame, ex_current_point_df: DataFrame, current_personnel_df: DataFrame) -> Tuple[DataFrame, DataFrame, DataFrame]:
        '''
        整理所有输出的dataframe
        '''
        point_output_cols: List[str] = [
            '采样点编号', '单元', '检测地点',
            '工种', '日接触时间', '检测因素',
            '采样数量/天', '采样天数', '采样日程',
            '空白编号', '起始编号', '终止编号'
        ]
        ex_point_output_cols: List[str] = [
            '采样点编号', '单元', '检测地点',
            '工种', '日接触时间', '检测因素',
            '采样数量/天', '采样天数', '采样日程',
            '样品编号', '代表时长'
        ]
        personnel_output_cols: List[str] = [
            '采样点编号', '单元', '工种', '日接触时间',
            '检测因素', '采样数量/天', '采样天数',
            '采样日程', '个体编号'
        ]
        output_current_point_df: DataFrame = current_point_df[point_output_cols]
        output_ex_current_point_df: DataFrame = ex_current_point_df[ex_point_output_cols]
        output_current_personnel_df: DataFrame = current_personnel_df[personnel_output_cols]
        return output_current_point_df, output_ex_current_point_df, output_current_personnel_df

    def get_single_day_dfs_stat(self, current_point_df: DataFrame, current_personnel_df: DataFrame) -> DataFrame:
        # 整理定点和个体的样品信息
        pivoted_point_df: DataFrame = pd.pivot_table(current_point_df, index=[
                                                     '检测因素'], aggfunc={'空白编号': max, '起始编号': min, '终止编号': max})
        # 增加个体样品数量为0时的处理方法
        # [x] 增加空白样品数量为0时的处理方法
        if current_personnel_df.shape[0] != 0:
            pivoted_personnel_df: DataFrame = (
                pd.pivot_table(current_personnel_df, index=[
                               '检测因素'], values='个体编号', aggfunc=[min, max])
                .stack()
                .reset_index()
                .set_index('检测因素')
                .drop('level_1', axis=1)
                .rename(columns={'min': '个体起始编号', 'max': '个体终止编号'})
            )
        else:
            pivoted_personnel_df = pd.DataFrame(columns=['个体起始编号', '个体终止编号'])
            pivoted_personnel_df.index.name = '检测因素'

        # 合并空白、定点和个体的信息
        counted_df: DataFrame = (
            pd.concat([pivoted_point_df, pivoted_personnel_df], axis=1)
            .fillna(0)
            .applymap(int)
        )
        # 统计空白、定点和个体的数量
        counted_df['空白数量'] = counted_df['空白编号'].apply(
            lambda x: 2 if x != 0 else 0)
        counted_df['定点数量'] = counted_df.apply(
            lambda x: x['终止编号'] - x['起始编号'] + 1 if x['终止编号'] != 0 else 0, axis=1)
        counted_df['个体数量'] = counted_df.apply(
            lambda x: x['个体终止编号'] - x['个体起始编号'] + 1 if x['个体终止编号'] != 0 else 0, axis=1)
        counted_df['总计'] = counted_df['空白数量'] + \
            counted_df['定点数量'] + counted_df['个体数量']
        # 统计空白、定点和个体的编号范围
        counted_df['空白编号范围'] = counted_df.apply(
            self.get_blank_count_range, axis=1)
        counted_df['定点编号范围'] = counted_df.apply(
            self.get_point_count_range, axis=1)
        counted_df['个体编号范围'] = counted_df.apply(
            self.get_personnel_count_range, axis=1)
        counted_df['编号范围'] = (
            self.project_number
            + counted_df
            .apply(self.get_range_str, axis=1)
        )
        counted_df['检测因素c'] = counted_df.index
        counted_df['保存时间'] = counted_df['检测因素c'].apply(
            self.get_counted_df_save_info)

        # counted_df.drop('检测因素c')
        # counted_df['保存时间'] = counted_df.apply(self.get_counted_df_save_info, axis=1)
        # counted_df['编号范围'] = counted_df['初始编号范围'].apply(remove_none)

        # cols: List[str] = ['总计', '编号范围']

        return counted_df

    def get_counted_df_save_info(self, factor: str) -> str:
        '''获得样品统计df里的各个检测因素的保存时间'''
        if factor.count('|') == 0:
            first_factor: str = factor
        else:
            first_factor: str = factor.split('|')[0]

        if first_factor in self.factor_reference_df['标识检测因素'].values:
            save_info_df: DataFrame = (
                self.factor_reference_df
                .query("标识检测因素 == @first_factor")
                .reset_index(drop=True)
            )
        # save_info_df = self
        # .factor_reference_df[self.factor_reference_df['标识检测因素'] == first_factor]
        # .reset_index(drop=True)
            save_info: str = str(save_info_df.loc[0, '保存时间'])
        else:
            save_info: str = '/'
        return save_info

    def get_exploded_point_df(self, r_current_point_df: DataFrame) -> List[str]:
        '''将定点df爆炸成多行的定点df'''
        # 空白编号
        int_list: List[str] = ['终止编号', '起始编号', '空白编号']
        r_current_point_df[int_list] = r_current_point_df[int_list].apply(int)
        if r_current_point_df['空白编号'] != 0:
            blank_list: List[str] = [
                f'{self.project_number}{r_current_point_df["空白编号"]:0>4d}-1',
                f'{self.project_number}{r_current_point_df["空白编号"]:0>4d}-2',
            ]
        else:
            blank_list: List[str] = [' ', ' ']
        # 定点编号
        point_list: List[int] = list(range(
            r_current_point_df['起始编号'], r_current_point_df['终止编号'] + 1))  # type: ignore
        point_str_list: List[str] = [
            f'{self.project_number}{i:0>4d}' for i in point_list]
        point_str_list_extra: List[str] = [' '] * (4 - len(point_str_list))
        point_str_list.extend(point_str_list_extra)
        # 空白加定点
        all_list: List[str] = blank_list + point_str_list
        return all_list

    def get_exploded_contact_duration(self, duration: float, size: int, full_size: int) -> List[str]:
        time_dec: Decimal = Decimal(str(duration))
        size_dec: Decimal = Decimal(str(size))
        time_list_dec: List[Decimal] = []
        if time_dec < Decimal('0.25') * size_dec:
            time_list_dec.append(time_dec)
        elif time_dec < Decimal('0.3') * size_dec:
            front_time_list_dec: List[Decimal] = [Decimal('0.25')] * (int(size) - 1)
            last_time_dec: Decimal = time_dec - sum(front_time_list_dec)
            time_list_dec.extend(front_time_list_dec)
            time_list_dec.append(last_time_dec)
        else:
            time_prec: int = int(time_dec.as_tuple().exponent)
            if time_prec == 2:
                prec_str: str = '0.00'
            else:
                prec_str: str = '0.0'
            judge_result: Decimal = time_dec / size_dec
            for i in range(int(size) - 1):
                result: Decimal = judge_result.quantize(Decimal(prec_str), ROUND_HALF_UP)
                time_list_dec.append(result)
            last_result: Decimal = time_dec - sum(time_list_dec)
            time_list_dec.append(last_result)
        
        time_list: List[float] = sorted(list(map(float, time_list_dec)), reverse=False)
        str_time_list: list[str] = list(map(str, time_list))
        blank_cell_list: list[str] = ['－', '－']
        complement_cell_list: list[str] = [' '] * (full_size - len(time_list))
        all_time_list: list[str] = blank_cell_list + str_time_list + complement_cell_list

        return all_time_list


    # 　（失败）重构将生成的样品编号写入bytesio的功能
    # def write_deleterious_substance_dfs_xlsx(self, types_order: List[str]) -> BytesIO:
    #     '''
    #     获得所有样品信息的编号，并写入bytesio文件里
    #     '''
    #     # 初始化已占用编号和BytesIO文件
    #     engaged_num: int = 0
    #     file_io: BytesIO = BytesIO()
    #     # 保证采样类型齐全
    #     if sorted(types_order) != sorted(self.default_types_order):
    #         types_order = self.default_types_order.copy()
    #     # 初始化采样日程
    #     schedule_days = range(1, self.schedule_days + 1)
    #     # 打开bytesio文件用于存储信息
    #     with pd.ExcelWriter(file_io) as excel_writer:
    #         # 循环读取
    #         for schedule_day in schedule_days:
    #             for type_order in types_order:
    #                 # 获得对应的df名称
    #                 df_name: str = f'current_{self.df_name_map[type_order]}_df'
    #                 # 根据采样类型用不同函数处理
    #                 locals()[df_name] = (
    #                     self.single_day_func_map[type_order]
    #                     (engaged_num, schedule_day)
    #                 )
    #                 # 更新已占用编号
    #                 engaged_num = self.refresh_engaged_num(
    #                     locals()[df_name],
    #                     type_order,
    #                     engaged_num
    #                 )
    #                 # 为定点信息的检测因素加上对应的空白编号
    #                 if type_order == '定点':
    #                 # 写入到excel里


    def get_dfs_num(self, types_order: List[str]) -> None:
        '''
        获得所有样品信息的编号，并写入bytesio文件里
        '''
        engaged_num: int = 0
        file_io: BytesIO = BytesIO()

        if sorted(types_order) != sorted(self.default_types_order):
            types_order = self.default_types_order.copy()
        schedule_list = range(1, self.schedule_days + 1)
        # 打开bytesio文件用于存储信息
        with pd.ExcelWriter(file_io) as excel_writer:
            # 循环采样日程
            for schedule_day in schedule_list:
                # 定点检测信息的空白编号和同一天的空白样品信息不一致
                # 定点检测信息可能要先添加样品编号，再添加空白信息
                # 考虑修改为函数工厂模式（放弃，要考虑空白编号的先后位置）
                for type_order in types_order:
                    if type_order == '空白':
                        current_blank_df: DataFrame = self.get_single_day_blank_df(
                            engaged_num, schedule_day)
                        engaged_num = self.refresh_engaged_num(
                            current_blank_df, type_order, engaged_num)
                    elif type_order == '定点':
                        current_point_df: DataFrame = self.get_single_day_point_df(
                            engaged_num, schedule_day)
                        # 添加一个函数，用于获得定点的空白信息
                        engaged_num = self.refresh_engaged_num(
                            current_point_df, type_order, engaged_num)
                    elif type_order == '个体':
                        current_personnel_df: DataFrame = self.get_single_day_personnel_df(
                            engaged_num, schedule_day)
                        # 添加一个函数，用于获得个体的空白信息
                        engaged_num = self.refresh_engaged_num(
                            current_personnel_df, type_order, engaged_num)

                # 为定点信息加上检测因素对应的空白信息
                if current_blank_df.shape[0] != 0:  # type: ignore
                    r_current_point_df: DataFrame = pd.merge(
                        current_point_df, current_blank_df, how='left', on='标识检测因素').fillna(0)  # type: ignore
                else:
                    r_current_point_df = current_point_df.copy()  # type: ignore
                    r_current_point_df['空白编号'] = 0  # type: ignore
                # 爆炸的定点编号
                r_current_point_df['样品编号'] = r_current_point_df.apply(  # type: ignore
                    self.get_exploded_point_df, axis=1)  # type: ignore
                r_current_point_df['代表时长'] = (  # type: ignore
                    r_current_point_df.apply(lambda df:   # type: ignore
                    self.get_exploded_contact_duration(df['日接触时间'], df['采样数量/天'], 4),
                    axis=1
                    )
                )
                ex_current_point_df: DataFrame = r_current_point_df.explode(['样品编号', '代表时长']) # type: ignore
                # 为定点信息加上空白编号，失败会错位
                counted_df: DataFrame = self.get_single_day_dfs_stat(
                    r_current_point_df, current_personnel_df)  # type: ignore

                # 将处理好的df写入excel文件里
                current_blank_df.to_excel(  # type: ignore
                    excel_writer, sheet_name=f'空白D{schedule_day}', index=False
                )

                (
                    output_current_point_df,
                    output_ex_current_point_df,
                    output_current_personnel_df
                ) = self.trim_dfs(r_current_point_df, ex_current_point_df, current_personnel_df)  # type: ignore
                output_current_point_df.to_excel(
                    excel_writer, sheet_name=f'定点D{schedule_day}', index=False)  # type: ignore
                output_ex_current_point_df.to_excel(
                    excel_writer, sheet_name=f'爆炸定点D{schedule_day}', index=False)  # type: ignore
                output_current_personnel_df.to_excel(
                    excel_writer, sheet_name=f'个体D{schedule_day}', index=False)  # type: ignore
                counted_df.to_excel(
                    excel_writer, sheet_name=f'样品统计D{schedule_day}', index=True)

                # 将点位信息写入记录表模板
                self.write_point_deleterious_substance_docx(schedule_day, output_ex_current_point_df)
                self.write_personnel_deleterious_substance_docx(schedule_day, output_current_personnel_df)

                # 将样品统计信息写入流转单模板
                self.write_traveler_docx(schedule_day, counted_df)
                # 将其他检测因素信息写入记录表模板
        other_factors: List[str] = ["一氧化碳", "噪声", "高温"]
        # 不同检测因素调用不同方法处理
        # other_factors_map = {
        #     "一氧化碳": self.write_co_docx,
        #     "噪声": self.write_point_noise_docx,
        #     "高温": self.write_temperature_docx,
        # }
        for factor in other_factors:
            # 判断是否存在再调用相应方法处理
            factor_exists: bool = (
                self
                .point_info_df['检测因素']
                .isin([f'{factor}'])
                .any(bool_only=True)
            )
            if factor_exists:
                # other_factors_map[factor]()
                self.write_direct_reading_factors_docx(factor)
        personnel_noise_exists: bool =(
            self
            .personnel_info_df['检测因素']
            .isin(['噪声'])
            .any(bool_only=True)
        )
        if personnel_noise_exists:
            self.write_personnel_noise_docx()
        # 将样品编号写入excel文件里
        file_name: str = f'{self.project_number}-{self.company_name}样品信息.xlsx'
        safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
        if not os.path.exists(self.output_path):
            os.mkdir(self.output_path)
        else:
            pass
        output_file_path = os.path.join(self.output_path, safe_file_name)
        with open(output_file_path, 'wb') as output_file:
            output_file.write(file_io.getvalue())


    # 放弃 重构将每天样品信息写入到对应模板的功能
    # def write_sample_info_docx(self) -> None:
    #     '''将样品信息写入模板'''
        #
        # 获得当天样品信息

    # def write_dfs_to_output_folder(self) -> None:
    #     '''将样品信息写入模板'''
    #     # 将样品编号写入excel文件里
    #     file_io: BytesIO = self.get_dfs_num(self.default_types_order)
    #     file_name: str = f'{self.project_number}-{self.company_name}样品信息.xlsx'
    #     safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
    #     if not os.path.exists(self.output_path):
    #         os.mkdir(self.output_path)
    #     else:
    #         pass
    #     output_file_path = os.path.join(self.output_path, safe_file_name)
    #     with open(output_file_path, 'wb') as output_file:
    #         output_file.write(file_io.getvalue())
    #     sheet_names = self.get_sheet_names(file_io)

    # [x] 将定点仪器直读检测因素的信息写入模板的方法合并
    def write_direct_reading_factors_docx(self, other_point_factor: str) -> None:
        # 获得检测因素的信息
        current_factor_info: Dict[str, Any] = self.templates_info[other_point_factor]
        join_char: str = current_factor_info['join_char']
        # 获得检测因素的点位信息
        current_factor_df: DataFrame = (
            self.point_info_df
            .query(f'标识检测因素 == "{other_point_factor}"')
            .reset_index(drop=True)
        )
        # 读取检测因素模板
        current_factor_template: str = current_factor_info['template_path']
        document = Document(current_factor_template)
        # 判断需要的记录表的页数
        table_pages: int = (
            math.ceil(
                (len(current_factor_df) - current_factor_info['first_page_rows'])
                / current_factor_info['late_page_rows']
            )
            + 1
        )
        # 根据不同页数，增减表格
        if table_pages == 1:
            # 删除第二页的表格
            rm_table = document.tables[2]
            t = rm_table._element
            t.getparent().remove(t)
            # 删除最后一个段落
            paragraphs = document.paragraphs
            rm_paragraphs1 = paragraphs[-1]
            rm_p1 = rm_paragraphs1._element
            rm_p1.getparent().remove(rm_p1)
            # 删除倒数第二个段落，即模板的第一页的换页符
            rm_paragraphs2 = paragraphs[-2]
            rm_p2 = rm_paragraphs2._element
            rm_p2.getparent().remove(rm_p2)
        elif table_pages == 2:
            pass # 跳过
        else:
            # 循环增加表格
            for _ in range(table_pages - 2):
                # 复制第二页的表格
                cp_table = document.tables[2]
                new_table = deepcopy(cp_table)
                # 在模板末尾增加段落
                new_paragraph = document.add_page_break()
                # 增加复制的表格
                new_paragraph._p.addnext(new_table._element)
                # 再增加一个段落
                document.add_paragraph()
        # [x] 写入信息
        # # 处理后的模板的所有表格
        tables = document.tables
        # 分析不同表格的写入信息
        for table_page in range(table_pages):
            # 获得当前表格的相应信息的索引
            if table_page == 0:
                index_first: int = 0
                index_last: int = current_factor_info['first_page_rows'] - 1
            else:
                index_first: int = (
                    current_factor_info['late_page_rows']
                    * (table_page - 1)
                    + current_factor_info['first_page_rows']
                )
                index_last: int = (
                    current_factor_info['first_page_rows']
                    + table_page
                    * current_factor_info['late_page_rows']
                    - 1
                )
            # 筛选出当前表格的信息
            if index_first == index_last:
                current_df: DataFrame = (
                    current_factor_df
                    .query(f'index == {index_first}')
                    .reset_index(drop=True)
                )
            else:
                current_table = tables[table_page + 1]
                current_df: DataFrame = (
                    current_factor_df
                    .query(f'index >= {index_first} and index <= {index_last}')
                    .reset_index(drop=True)
                )
            current_table = tables[table_page + 1]
            # 按行循环选取单元格
            for r_i in range(current_df.shape[0]):
                current_row_list = [
                    current_df.loc[r_i, '采样点编号'],
                    f"{current_df.loc[r_i, '单元']}{join_char}{current_df.loc[r_i, '检测地点']}",
                    current_df.loc[r_i, '日接触时间'],
                ]
                # 再循环列选取单元格，并写入相应信息
                for i, c_i in enumerate(current_factor_info['available_cols']):
                    current_cell = (
                        current_table.rows[
                            r_i * current_factor_info['item_rows']
                            + current_factor_info['title_rows']
                            ]
                        .cells[c_i]
                    )
                    current_cell.text = str(current_row_list[i])
                    current_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
                    current_cell.paragraphs[0].runs[0].font.size = Pt(9)
                    # current_cell.paragraphs[0].runs[0].font.name = '宋体'
        # [x] 样式调整
        # 写入基本信息
        info_table = tables[0]
        code_cell = (
            info_table
            .rows[current_factor_info['project_num_row']]
            .cells[current_factor_info['project_num_col']]
        )
        comp_cell = (
            info_table
            .rows[current_factor_info['company_name_row']]
            .cells[current_factor_info['company_name_col']]
        )
        code_cell.text = self.project_number
        code_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
        comp_cell.text = self.company_name
        comp_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
        # [x] 样式调整
        # 保存
        file_name: str = f'{other_point_factor}记录表'
        safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
        if not os.path.exists(self.output_path):
            os.mkdir(self.output_path)
        else:
            pass
        output_file_path: str = os.path.join(
            self.output_path, f'{safe_file_name}.docx')
        document.save(output_file_path)


    # [x] 个体噪声
    def write_personnel_noise_docx(self) -> None:
        '''将个体噪声信息写入模板'''
        current_factor_info: Dict[str, Any] = self.templates_info['个体噪声']
        # 获得个体噪声信息
        current_factor_df: DataFrame = (
            self.personnel_info_df
            .query('标识检测因素 == "噪声"')
            .reset_index(drop=True)
        )
        # 读取个体噪声模板
        current_factor_template: str = current_factor_info['template_path']
        document = Document(current_factor_template)
        # 判断需要的记录表的页数
        table_pages: int = (
            math.ceil(
                (len(current_factor_df) - current_factor_info['first_page_rows'])
                / current_factor_info['late_page_rows']
            )
            + 1
        )
        # 根据不同页数，增减表格
        if table_pages == 1:
            # 删除第二页的表格
            rm_table = document.tables[2]
            t = rm_table._element
            t.getparent().remove(t)
            # 删除最后一个段落
            paragraphs = document.paragraphs
            rm_paragraphs1 = paragraphs[-1]
            rm_p1 = rm_paragraphs1._element
            rm_p1.getparent().remove(rm_p1)
            # 删除倒数第二个段落，即模板的第一页的换页符
            rm_paragraphs2 = paragraphs[-2]
            rm_p2 = rm_paragraphs2._element
            rm_p2.getparent().remove(rm_p2)
        elif table_pages == 2:
            pass # 跳过
        else:
            # 循环增加表格
            for _ in range(table_pages - 2):
                # 复制第二页的表格
                cp_table = document.tables[2]
                new_table = deepcopy(cp_table)
                # 在模板末尾增加段落
                new_paragraph = document.add_page_break()
                # 增加复制的表格
                new_paragraph._p.addnext(new_table._element)
                # 再增加一个段落
                document.add_paragraph()
        # 写入信息
        # 处理后的模板的所有表格
        tables = document.tables
        # 分析不同表格的写入信息
        for table_page in range(table_pages):
            # 获得当前表格的相应信息的索引
            if table_page == 0:
                index_first: int = 0
                index_last: int = current_factor_info['first_page_rows'] - 1
            else:
                index_first: int = (
                    current_factor_info['late_page_rows']
                    * (table_page - 1)
                    + current_factor_info['first_page_rows']
                )
                index_last: int = (
                    current_factor_info['first_page_rows']
                    + table_page
                    * current_factor_info['late_page_rows']
                    - 1
                )
            # 筛选出当前表格的信息
            if index_first == index_last:
                current_df: DataFrame = (
                    current_factor_df
                    .query(f'index == {index_first}')
                    .reset_index(drop=True)
                )
            else:
                current_table = tables[table_page + 1]
                current_df: DataFrame = (
                    current_factor_df
                    .query(f'index >= {index_first} and index <= {index_last}')
                    .reset_index(drop=True)
                )
            current_table = tables[table_page + 1]
            # 按行循环选取单元格
            for r_i in range(current_df.shape[0]):
                current_row_list = [
                    current_df.loc[r_i, '采样点编号'],
                    f"{current_df.loc[r_i, '单元']} {current_df.loc[r_i, '工种']}\n",
                    current_df.loc[r_i, '日接触时间'],
                ]
                # 再循环列选取单元格，并写入相应信息
                for i, c_i in enumerate(current_factor_info['available_cols']):
                    current_cell = (
                        current_table.rows[
                            r_i * current_factor_info['item_rows']
                            + current_factor_info['title_rows']
                            ]
                        .cells[c_i]
                    )
                    current_cell.text = str(current_row_list[i])
                    current_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
                    current_cell.paragraphs[0].runs[0].font.size = Pt(6.5)
                    # current_cell.paragraphs[0].runs[0].font.name = '宋体'
        info_table = tables[0]
        code_cell = info_table.rows[0].cells[1]
        comp_cell = info_table.rows[1].cells[1]

        code_cell.text = self.project_number
        code_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
        comp_cell.text = self.company_name
        comp_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
        # [x] 单元格样式
        file_name: str = '个体噪声记录表'
        safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
        if not os.path.exists(self.output_path):
            os.mkdir(self.output_path)
        else:
            pass
        output_file_path: str = os.path.join(
            self.output_path, f'{safe_file_name}.docx')
        document.save(output_file_path)

#     # [x] 定点噪声
#     def write_point_noise_docx(self) -> None:
#         '''将定点噪声信息写入模板'''
#         # 获得定点噪声信息
#         point_noise_df: DataFrame = (
#             self.point_info_df
#             .query('标识检测因素 == "噪声"')
#             .reset_index(drop=True)
#         )
#         # [x] 模板文件路径和样式
#         point_noise_template: str = './templates/定点噪声.docx'
#         point_noise_document = Document(point_noise_template)
#         # 判断需要的记录表的页数
#         table_pages: int = math.ceil((len(point_noise_df) - 9) / 11) + 1
#         if table_pages == 1:
#             rm_table = point_noise_document.tables[2]
#             t = rm_table._element
#             t.getparent().remove(t)

#             rm_page_break = point_noise_document.paragraphs[-2]
#             pg = rm_page_break._element
#             pg.getparent().remove(pg)
#             rm_page_break2 = point_noise_document.paragraphs[-2]
#             pg2 = rm_page_break2._element
#             pg2.getparent().remove(pg2)
#         elif table_pages == 2:
#             pass
#         else:
#             for _ in range(table_pages - 2):
#                 cp_table = point_noise_document.tables[2]
#                 new_table = deepcopy(cp_table)
#                 new_paragraph = point_noise_document.add_page_break()
#                 new_paragraph._p.addnext(new_table._element)
#                 point_noise_document.add_paragraph()
#         tables = point_noise_document.tables
#         for table_page in range(table_pages):
#             if table_page == 0:
#                 index_first: int = 0
#                 index_last: int = 9
#             else:
#                 index_first: int = 11 * table_page - 1
#                 index_last: int = 11 * table_page + 10
#             current_df: DataFrame = (
#                 point_noise_df
#                 .query(f'index >= {index_first} and index <= {index_last}')
#                 .reset_index(drop=True)
#             )
#             current_table = tables[table_page + 1]
#             for r_i in range(current_df.shape[0]):
#                 current_row_list: List[str] = [
#                     current_df.loc[r_i, '采样点编号'],  # type: ignore
#                     f"{current_df.loc[r_i, '单元']} {current_df.loc[r_i, '工种']}",
#                     current_df.loc[r_i, '日接触时间'],
#                 ]
#                 for c_i in range(3):
#                     current_cell = current_table.rows[r_i + 2].cells[c_i]
#                     current_cell.text = str(current_row_list[c_i])
#                     # [x] 单元格样式
#         info_table = tables[0]
#         code_cell = info_table.rows[0].cells[1]
#         comp_cell = info_table.rows[1].cells[1]

#         code_cell.text = self.project_number
#         comp_cell.text = self.company_name
#         # [x] 单元格样式
#         file_name: str = '定点噪声记录表'
#         safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
#         if not os.path.exists(self.output_path):
#             os.mkdir(self.output_path)
#         else:
#             pass
#         output_file_path: str = os.path.join(
#             self.output_path, f'{safe_file_name}.docx')
#         point_noise_document.save(output_file_path)

#     # [x] 一氧化碳
#     def write_co_docx(self) -> None:
#         # 获得一氧化碳信息
#         co_df: DataFrame = (
#             self.point_info_df
#             .query('标识检测因素 == "一氧化碳"')
#             .reset_index(drop=True)
#         )
#         # [x] 模板文件路径和样式
#         co_template: str = './templates/一氧化碳.docx'
#         co_document = Document(co_template)
#         # 判断需要的记录表的页数
#         table_pages: int = math.ceil(len(co_df) / 5)
#         if table_pages == 1:
#             rm_table = co_document.tables[2]
#             t = rm_table._element
#             t.getparent().remove(t)

#             rm_paragraph = co_document.paragraphs[-1]
#             pg = rm_paragraph._element
#             pg.getparent().remove(pg)
#             rm_paragraph2 = co_document.paragraphs[-1]
#             pg2 = rm_paragraph2._element
#             pg2.getparent().remove(pg2)
#             rm_paragraph3 = co_document.paragraphs[-1]
#             pg3 = rm_paragraph3._element
#             pg3.getparent().remove(pg3)
#             # rm_page_break = co_document.paragraphs[-2]
#             # rm_page_break = rm_page_break._element
#             # rm_page_break.getparent().remove(rm_page_break)
#             # rm_paragraph5 = co_document.paragraphs[-1]
#             # pg5 = rm_paragraph5._element
#             # pg5.getparent().remove(pg5)
#             # rm_page_break = co_document.paragraphs[-1]
#         elif table_pages == 2:
#             pass
#         else:
#             for _ in range(table_pages - 2):
#                 cp_table = co_document.tables[2]
#                 new_table = deepcopy(cp_table)
#                 rm_paragraph = co_document.paragraphs[-1]
#                 pg = rm_paragraph._element
#                 pg.getparent().remove(pg)
#                 new_paragraph = co_document.add_page_break()
#                 new_paragraph._p.addnext(new_table._element)
#                 co_document.add_paragraph()
#         tables = co_document.tables
#         for table_page in range(table_pages):
#             first_index: int = 5 * table_page
#             last_index: int = 5 * table_page + 4
#             current_df = co_df.iloc[first_index:last_index]
#             current_table = tables[table_page + 1]
#             for r_i in range(current_df.shape[0]):
#                 current_row_list = [
#                     current_df.loc[r_i, '采样点编号'],
#                     f"{current_df.loc[r_i, '单元']}\n{current_df.loc[r_i, '检测地点']}",
#                     # current_df.loc[r_i, '日接触时间'],
#                 ]
#                 for c_i in range(2):
#                     current_cell = current_table.rows[r_i * 4 + 2].cells[c_i]
#                     current_cell.text = str(current_row_list[c_i])
#                     # [x] 单元格样式
#         info_table = tables[0]
#         code_cell = info_table.rows[0].cells[1]
#         comp_cell = info_table.rows[0].cells[3]

#         code_cell.text = self.project_number
#         comp_cell.text = self.company_name
#         file_name: str = '一氧化碳CO记录表'
#         safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
#         if not os.path.exists(self.output_path):
#             os.mkdir(self.output_path)
#         else:
#             pass
#         output_file_path: str = os.path.join(
#             self.output_path, f'{safe_file_name}.docx')
#         co_document.save(output_file_path)


# # [x] 二氧化碳（考虑取消）
# # [x] 高温

#     def write_temperature_docx(self) -> None:
#         temp_df: DataFrame = (
#             self.point_info_df
#             .query('标识检测因素 == "高温"')
#             .reset_index(drop=True)
#         )
#         # [x] 模板文件路径和样式
#         temp_template: str = './templates/高温.docx'
#         temp_document = Document(temp_template)
#         # 判断需要的记录表的页数
#         table_pages: int = math.ceil((len(temp_df) - 1) / 2) + 1
#         if table_pages == 1:
#             rm_table = temp_document.tables[2]
#             t = rm_table._element
#             t.getparent().remove(t)

#             rm_paragraph = temp_document.paragraphs[-1]
#             pg = rm_paragraph._element
#             pg.getparent().remove(pg)
#             rm_paragraph2 = temp_document.paragraphs[-1]
#             pg2 = rm_paragraph2._element
#             pg2.getparent().remove(pg2)
#             # rm_paragraph3 = co_document.paragraphs[-1]
#             # pg3 = rm_paragraph3._element
#             # pg3.getparent().remove(pg3)
#             # rm_page_break = temp_document.paragraphs[-2]
#             # rm_page_break = rm_page_break._element
#             # rm_page_break.getparent().remove(rm_page_break)
#             # rm_paragraph5 = co_document.paragraphs[-1]
#             # pg5 = rm_paragraph5._element
#             # pg5.getparent().remove(pg5)
#             # rm_page_break = co_document.paragraphs[-1]
#         elif table_pages == 2:
#             pass
#         else:
#             for _ in range(table_pages - 2):
#                 cp_table = temp_document.tables[2]
#                 new_table = deepcopy(cp_table)
#                 rm_paragraph = temp_document.paragraphs[-1]
#                 pg = rm_paragraph._element
#                 pg.getparent().remove(pg)
#                 new_paragraph = temp_document.add_page_break()
#                 new_paragraph._p.addnext(new_table._element)
#                 temp_document.add_paragraph()
#         tables = temp_document.tables
#         for table_page in range(table_pages):
#             if table_page == 0:
#                 query_str: str = 'index == 0'
#             else:
#                 index_first: int = table_page * 2 - 1
#                 index_last: int = table_page * 2
#                 query_str: str = f'index >= {index_first} and index <= {index_last}'
#             current_df: DataFrame = (
#                 temp_df
#                 .query(query_str)
#                 .reset_index(drop=True)
#             )

#             current_table = tables[table_page + 1]

#             for r_i in range(current_df.shape[0]):
#                 current_row_list = [
#                     current_df.loc[r_i, '采样点编号'],
#                     f"{current_df.loc[r_i, '单元']}\n{current_df.loc[r_i, '检测地点']}",
#                     # current_df.loc[r_i, '日接触时间'],
#                 ]
#                 for c_i in range(2):
#                     current_cell = current_table.rows[r_i * 9 + 3].cells[c_i]
#                     current_cell.text = str(current_row_list[c_i])
#                     # [x] 单元格样式
#         info_table = tables[0]
#         code_cell = info_table.rows[0].cells[1]
#         comp_cell = info_table.rows[1].cells[1]

#         code_cell.text = self.project_number
#         comp_cell.text = self.company_name
#         # [x] 单元格样式
#         file_name: str = '高温记录表.docx'
#         safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
#         if not os.path.exists(self.output_path):
#             os.mkdir(self.output_path)
#         else:
#             pass
#         output_file_path = os.path.join(self.output_path, safe_file_name)
#         temp_document.save(output_file_path)


    def write_traveler_docx(self, schedule_day: int, counted_df: DataFrame) -> None:
        # 将流转单信息写入模板
        traveler_path: str = './templates/样品流转单.docx'
        traveler_document = Document(traveler_path)
        project_num_cell = traveler_document.tables[0].rows[0].cells[1]
        project_num_cell.text = self.project_number
        # [x] 样式
        # 判断需要的流转单的页数
        table_pages: int = math.ceil(len(counted_df) / 8)
        for _ in range(table_pages - 1):
            cp_table = traveler_document.tables[0]
            new_table = deepcopy(cp_table)
            cp_paragraph = traveler_document.paragraphs[0]
            last_paragraph = traveler_document.add_page_break()
            last_paragraph._p.addnext(new_table._element)
            traveler_document.add_paragraph(cp_paragraph.text)

        tables = traveler_document.tables

        for table_page in range(table_pages):
            first_index: int = 8 * table_page
            last_index: int = 8 * table_page + 7
            # .reset_index(drop=True)
            current_df: DataFrame = counted_df.iloc[first_index : last_index + 1]
            current_table = tables[table_page]
            for r_i in range(len(current_df)):
                current_index_name = current_df.iloc[r_i].name
                # print(current_index_name)
                current_row_list = [
                    current_df.loc[current_index_name, "编号范围"],  # type: ignore
                    current_index_name,
                    current_df.loc[current_index_name, "保存时间"],  # type: ignore
                    current_df.loc[current_index_name, "总计"],  # type: ignore
                ]
                for c_i in list(range(4)):
                    match_cols_list: List[int] = [0, 1, 3, 4]
                    current_cell = (
                        current_table
                        .rows[r_i + 2]
                        .cells[match_cols_list[c_i]]
                    )
                    current_cell.text = str(current_row_list[c_i])
                    current_cell.paragraphs[0].runs[0].font.size = Pt(7.5)
                    current_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
                    # [x] 单元格样式
                    if '\\n' in current_cell.text:
                        new_text: str = current_cell.text.replace('\\n', '\n')
                        current_cell.text = new_text
                        current_cell.paragraphs[0].runs[0].font.size = Pt(7.5)
                        current_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
        file_name: str = f'D{schedule_day}-样品流转单'
        safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
        # output_path = f'{os.path.expanduser("~/Desktop")}/{self.project_number}记录表'
        if not os.path.exists(self.output_path):
            os.mkdir(self.output_path)
        else:
            pass
        output_file_path: str = os.path.join(
            self.output_path, f'{safe_file_name}.docx')
        traveler_document.save(output_file_path)

    def write_personnel_deleterious_substance_docx(self, schedule_day: int, current_personnel_df: DataFrame) -> None:
        # 将个体有害物质写入模板
        items = current_personnel_df['检测因素'].drop_duplicates().tolist()
        for item in items:
            # 导入个体模板
            personnel_template_path: str = r'./templates/有害物质个体采样记录.docx'
            personnel_document = Document(personnel_template_path)
            # 获得当前检测因素的dataframe
            current_factor_df = current_personnel_df[current_personnel_df['检测因素'] == item].reset_index(
                drop=True)
            # 计算需要的记录表页数
            table_pages: int = math.ceil((len(current_factor_df) - 11) / 6 + 2)
            if table_pages == 1:
                rm_table = personnel_document.tables[2]
                t = rm_table._element
                t.getparent().remove(t)
                rm_page_break = personnel_document.paragraphs[-2]
                pg = rm_page_break._element
                pg.getparent().remove(pg)
                rm_page_break2 = personnel_document.paragraphs[-2]
                pg2 = rm_page_break2._element
                pg2.getparent().remove(pg2)
            elif table_pages == 2:
                pass
            else:
                for _ in range(table_pages - 2):
                    cp_table = personnel_document.tables[2]
                    new_table = deepcopy(cp_table)
                    # new_paragraph = point_document.add_paragraph()
                    new_paragraph = personnel_document.add_page_break()
                    new_paragraph._p.addnext(new_table._element)
                    # paragraph = point_document.add_paragraph()
                    # paragraph._p.addnext(new_table._element)
                    # point_document.add_page_break()
                    personnel_document.add_paragraph()

            tables = personnel_document.tables

            for table_page in range(table_pages):
                if table_page == 0:
                    index_first: int = 0
                    index_last: int = 5
                else:
                    index_first: int = 6 * table_page - 1
                    index_last: int = 6 * table_page + 5

                current_df = (
                    current_factor_df
                    .query(f'index >= {index_first} and index <= {index_last}')
                    .reset_index(drop=True)
                )
                current_table = tables[table_page + 1]
                for r_i in range(current_df.shape[0]):
                    current_row_list = [
                        current_df.loc[r_i, '采样点编号'],
                        f"{current_df.loc[r_i, '单元']}\n{current_df.loc[r_i, '工种']}",
                        f"{self.project_number}{current_df.loc[r_i, '个体编号']:0>4d}",
                    ]
                    for c_i in range(3):
                        current_cell = (
                            current_table
                            .rows[r_i * 3 + 2]
                            .cells[c_i]
                        )
                        current_cell.text = str(current_row_list[c_i])
                        if c_i <= 1:
                            current_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
                        else:
                            current_cell.paragraphs[0].runs[0].font.size = Pt(6.5)

            # 写入基本信息
            info_table = tables[0]
            code_cell = info_table.rows[0].cells[1]
            comp_cell = info_table.rows[0].cells[4]
            item_cell = info_table.rows[3].cells[1]
            code_cell.text = self.project_number
            comp_cell.text = self.company_name
            item_cell.text = item
            # 基本信息的样式
            for cell in [code_cell, comp_cell, item_cell]:
                p = cell.paragraphs[0]
                p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # type: ignore
                p.runs[0].font.size = Pt(8)

            # 保存到桌面文件夹里
            file_name: str = f'D{schedule_day}-个体-{item}'
            safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
            # output_path = f'{os.path.expanduser("~/Desktop")}/{self.project_number}记录表'
            if not os.path.exists(self.output_path):
                os.mkdir(self.output_path)
            else:
                pass
            output_file_path: str = os.path.join(
                self.output_path, f'{safe_file_name}.docx')
            personnel_document.save(output_file_path)

    def write_point_deleterious_substance_docx(self, schedule_day: int, current_point_df: DataFrame) -> None:
        # 将定点有害物质写入模板
        items = current_point_df['检测因素'].drop_duplicates().tolist()
        for item in items:
            # 导入定点模板
            point_template_path: str = r'./templates/有害物质定点采样记录.docx'
            point_document = Document(point_template_path)

            # 获得当前检测因素的dataframe
            current_factor_df = current_point_df[current_point_df['检测因素'] == item].reset_index(
                drop=True)
            # 计算需要的记录表页数
            table_pages: int = math.ceil(
                (len(current_factor_df) - 42) / 24 + 2)
            # 按照页数来增减表格数量
            if table_pages == 1:
                rm_table = point_document.tables[2]
                t = rm_table._element
                t.getparent().remove(t)
                rm_page_break = point_document.paragraphs[-2]
                pg = rm_page_break._element
                pg.getparent().remove(pg)
                rm_page_break2 = point_document.paragraphs[-2]
                pg2 = rm_page_break2._element
                pg2.getparent().remove(pg2)
            elif table_pages == 2:
                pass
                # rm_page_break = point_document.paragraphs[-2]
                # pg = rm_page_break._element
                # pg.getparent().remove(pg)
            else:
                for _ in range(table_pages - 2):
                    cp_table = point_document.tables[2]
                    new_table = deepcopy(cp_table)
                    new_paragraph = point_document.add_page_break()
                    new_paragraph._p.addnext(new_table._element)
                    point_document.add_paragraph()

            tables = point_document.tables
            for table_page in range(table_pages):
                if table_page == 0:
                    index_first: int = 0
                    index_last: int = 17
                else:
                    index_first: int = 24 * table_page - 6
                    index_last: int = 24 * table_page + 17
                current_df = (
                    current_factor_df
                    .query(f'index >= {index_first} and index <= {index_last}')
                    .reset_index(drop=True)
                )
                # 向指定表格填写数据
                current_table = tables[table_page + 1]
                for r_i in range(current_df.shape[0]):
                    current_row_list = [
                        current_df.loc[r_i, '采样点编号'],
                        f"{current_df.loc[r_i, '单元']}\n{current_df.loc[r_i, '检测地点']}",
                        current_df.loc[r_i, '样品编号'],
                        current_df.loc[r_i, '代表时长'],
                        # current_df.loc[r_i, '采样数量/天'],
                    ]
                    for l_i, c_i in enumerate([0, 1, 2, 9]):
                        current_cell = current_table.rows[r_i + 2].cells[c_i]
                        current_cell.text = str(current_row_list[l_i])
                        # [x] 考虑增加更改字体样式
                        if c_i != 2:
                            current_cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # type: ignore
                        else:
                            current_cell.paragraphs[0].runs[0].font.size = Pt(6.5)
                # FAILED 当采样数量是1时，合并接触时长单元格
                    # judge_merge: bool = current_df.loc[r_i, '日接触时间'] / current_df.loc[r_i, '采样数量/天'] < 0.25
                    # if judge_merge:
                    #     current_merge_first_cell = current_table.cell(r_i + 2, 9)
                    #     current_merge_last_cell = current_table.cell(r_i + 4, 9)
                    #     current_merge_first_cell.merge(current_merge_last_cell)
            # 写入基本信息
            info_table = tables[0]
            code_cell = info_table.rows[0].cells[1]
            comp_cell = info_table.rows[0].cells[4]
            item_cell = info_table.rows[3].cells[1]
            code_cell.text = self.project_number
            comp_cell.text = self.company_name
            item_cell.text = item
            # 基本信息的样式
            # [x] 考虑增加更改字体样式
            for cell in [code_cell, comp_cell, item_cell]:
                p = cell.paragraphs[0]
                p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # type: ignore
                p.runs[0].font.size = Pt(9)

            # 保存到桌面文件夹里
            file_name = f'D{schedule_day}-定点-{item}'
            safe_file_name: str = re.sub(r'[?*/\<>:"|]', ',', file_name)
            # output_path = f'{os.path.expanduser("~/Desktop")}/{self.project_number}记录表'
            if not os.path.exists(self.output_path):
                os.mkdir(self.output_path)
            else:
                pass
            output_file_path = os.path.join(
                self.output_path, f'{safe_file_name}.docx')
            point_document.save(output_file_path)

# 筛选某一天日程的所有定点和个体检测信息

# 新建一个存储在bytesio里的excel文件
# 考虑通过循环来将每日的定点和个体检测信息生成的数据（空白、定点和个体）存储到上述的excel文件
# 获得某一天日程的空气检测信息的空白dataframe，筛选出需要空白的检测因素并生成其空白编号
# 生成定点检测信息的样品编号范围，可以选择将爆炸样品编号的函数编写在这里
# 更新已占用编号的函数
# 每日的样品编号范围，用于流转单
# [x] 生成其他仪器只读的检测因素的记录表
# [x] 生成空白、定点和个体的记录表

# 建立一个基于OccupationalHealthItemInfo类的子类，为OccupationalHealthItemInfo类下的每天检测信息的类
# 可能考虑取消子类，因为部分检测参数（例如物理因素、CO和CO2等只需要一天，完全可以放在一个整体里）

    def refresh_engaged_num(self, current_df: DataFrame, type: str, engaged_num: int) -> int:
        '''更新已占用样品编号数'''
        # 按照df类型来更新编号
        # [x] 如果df长度为0时要
        # [x] 更新，使用字典模式。错误，无法使用
        default_types_order: List[str] = ['空白', '定点', '个体']
        type_num_dict = {
            '空白': '空白编号',
            '定点': '终止编号',
            '个体': '个体编号',
        }
        if current_df.shape[0] != 0 and type in default_types_order:
            new_engaged_num: int = current_df[type_num_dict[type]].astype(int).max()
            return new_engaged_num
        else:
            return engaged_num

        # 可能废弃
        # default_types_order: List[str] = ['空白', '定点', '个体']
        # if df.shape[0] == 0:
        #     return engaged_num
        # elif type in default_types_order:
        #     if type == '空白':
        #         new_engaged_num: int = df['空白编号'].astype(int).max()  # type: ignore
        #     elif type == '个体':
        #         new_engaged_num: int = df['个体编号'].astype(int).max()  # type: ignore
        #     elif type == '定点':
        #         new_engaged_num: int = df['终止编号'].astype(int).max()  # type: ignore
        #     return new_engaged_num  # type: ignore
        # else:
        #     return engaged_num
        # 已废弃
        # if df.shape[0] != 0:
        #     # df_cols: List[str] = df.columns.to_list()
        #     df_cols: List[Any] = list(df.columns)
        #     if '空白编号' in df_cols:
        #         new_engaged_num: int = df['空白编号'].astype(int).max()  # type: ignore
        #     elif '终止编号' in df_cols:
        #         new_engaged_num: int = df['终止编号'].astype(int).max()  # type: ignore
        #     elif '个体编号' in df_cols:
        #         new_engaged_num: int = df['个体编号'].astype(int).max()  # type: ignore
        #     return new_engaged_num
        # else:
        #     return engaged_num

    def custom_sort(self, str_list: List[str], key_list: List[str]) -> List[str]:
        '''
        列表的自定义排序
        '''
        if str_list[0] in key_list:
            sorted_str_list: List[str] = sorted(
                str_list, key=lambda x: key_list.index(x))
            return sorted_str_list
        else:
            return str_list

    def get_blank_count_range(self, blank_df: DataFrame):
        if blank_df['空白数量'] != 0:
            return f'{blank_df["空白编号"]:0>4d}-1, {blank_df["空白编号"]:0>4d}-2'
        else:
            return ' '

    def get_point_count_range(self, point_df: DataFrame):
        if point_df['定点数量'] == 0:
            return ' '
        elif point_df['定点数量'] == 1:
            return f'{point_df["起始编号"]:0>4d}'
        else:
            return f'{point_df["起始编号"]:0>4d}-{point_df["终止编号"]:0>4d}'

    def get_personnel_count_range(self, personnel_df: DataFrame):
        if personnel_df['个体数量'] == 0:
            return ' '
        elif personnel_df['个体数量'] == 1:
            return f'{personnel_df["个体起始编号"]:0>4d}'
        else:
            return f'{personnel_df["个体起始编号"]:0>4d}-{personnel_df["个体终止编号"]:0>4d}'

    def get_range_str(self, counted_df: DataFrame):
        range_list = [
            counted_df['空白编号范围'],
            counted_df['定点编号范围'],
            counted_df['个体编号范围']
        ]
        range_list = [i for i in range_list if i != ' ']
        range_str = ', '.join(range_list)  # type: ignore
        return range_str
