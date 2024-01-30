'''
处理系统生成编号，并写入模板里
主程序放置在这里
'''
import os
from typing import Dict, List, Any
# from datetime import datetime

import pandas as pd
from nptyping import DataFrame

from other_infos import templates_infos
from custom_method import get_exploded_contact_duration

class NewOccupationalHealthItemInfo():
    '''主程序'''
    def __init__(
            self,
            project_number: str,
            company_name: str,
            raw_df: DataFrame,
            templates_info: Dict[str, Dict[str, Any]],
        ) -> None:
        self.project_number: str = project_number
        self.company_name: str = company_name
        self.templates_info: Dict[str, Dict[str, Any]] = templates_info
        self.schedule_col: str = self.initialize_schedule()
        self.schedule_list: list[Any] = self.get_schedule_list()
        self.df: DataFrame = self.initialize_df(raw_df)
        self.blank_df: DataFrame = self.initialize_blank_df()
        self.point_df: DataFrame = self.initialize_point_df()

# 初始化
    def initialize_df(self, raw_df: DataFrame) -> DataFrame:
        '''初始化所有样品信息'''
        available_cols: List[str] = [
            '样品类型',
            '样品编号',
            '样品名称',
            '检测参数',
            '采样/送样日期',
            '样品描述',
            '单元',
            '工种/岗位',
            '检测地点',
            '测点编号',
            '第几天',
            '第几个频次',
            '采样方式',
            '作业人数',
            '日接触时长/h',
            '周工作天数/d',
        ]
        df: DataFrame = raw_df[available_cols]
        df_copy: DataFrame = df.copy()
        df_copy['样品编号'] = df_copy['样品编号'].str.replace(self.project_number, '', regex=False)

        return df_copy

    def initialize_blank_df(self) -> DataFrame:
        '''初始化空白信息'''
        raw_blank_df: DataFrame = (
            self # type: ignore
            .df
            .query('样品类型 == "空白样"')
            .reset_index(drop=True)
        )
        blank_df: DataFrame = (
            raw_blank_df
            .pivot(
                index=['检测参数', self.schedule_col],
                columns='第几个频次',
                values='样品编号'
            )
            .rename(columns={1: '空白编号1', 2: '空白编号2'})
            .reset_index(drop=False)
        )

        return blank_df

    def initialize_schedule(self) -> str:
        '''初始化采样日程'''
        if self.df['采样/送样日期'].isnull().all(): # type: ignore
            schedule_col: str = '第几天'
        else:
            schedule_col: str = '采样/送样日期'
        return schedule_col

    def get_schedule_list(self) -> List[Any]:
        '''获得采样日程'''
        # 可能是整数或者是日期
        schedule_list: List[Any] = (
            self
            .df[self.schedule_col]
            .drop_duplicates()
            .tolist()
        )
        return schedule_list

    def initialize_point_df(self) -> DataFrame:
        '''初始化定点信息'''
        # query_str: str = (
        #     '样品类型 == "普通样"'
        #     ' and '
        #     '采样方式 == "定点"'
        #     ' and '
        #     '样品名称 != "工作场所物理因素"'
        # )
        query_str: str = (
            '样品类型 == "普通样"'
            ' and '
            '采样方式 == "定点"'
            ' and '
            '样品名称 != "工作场所物理因素"'
            ' and '
            '样品描述 != "仪器直读"'
            # '样品编号 != "/"'
        )
        raw_point_df: DataFrame = (
            self # type: ignore
            .df
            .query(query_str)
            .reset_index(drop=True)
        )
        raw_point_df['样品编号'] = (
            raw_point_df['样品编号'] # type: ignore
            .astype(int)
        )
        groupby_point_df: DataFrame = (
            raw_point_df # type: ignore
            .groupby(
                [
                    '测点编号',
                    '单元',
                    '检测地点',
                    '工种/岗位',
                    '检测参数',
                    self.schedule_col,
                    r'日接触时长/h'
                ]
        )
        ['样品编号']
        .agg(list)
        .reset_index(drop=False)
        )
        groupby_point_df['采样数量/天'] = (
            groupby_point_df # type: ignore
            ['样品编号']
            .apply(len)
        )
        # [x] 是否合并代表时长列要改进
        groupby_point_df['是否合并代表时长'] = (
            groupby_point_df # type: ignore
            .apply(lambda df: True if df[r'日接触时长/h'] / df['采样数量/天'] < 0.25 else False, axis=1) # type: ignore
        )
        point_df: DataFrame = groupby_point_df.merge( # type: ignore
            self.blank_df,
            on=['检测参数', self.schedule_col],
            how='left'
        )
        point_df['代表时长'] = (
            point_df
            .apply(
                lambda df: self.get_exploded_contact_duration(
                    df[r'日接触时长/h'], df['采样数量/天']
                ),
                axis=1
            )
        )
        point_df['空白编号1'] = point_df['空白编号1'].fillna('-')
        point_df['空白编号2'] = point_df['空白编号2'].fillna('-')

        return point_df


company_name1: str = 'MSCN'
project_number1: str = '23ZDQ0063'
file_path1: str = os.path.join(os.path.expanduser('~'), 'Desktop', 'WT23ZDQ0063系统生成编号.xlsx')
raw_df1: DataFrame = pd.read_excel(file_path1)

new_project = NewOccupationalHealthItemInfo(
    company_name=company_name1,
    project_number=project_number1,
    raw_df=raw_df1,
    templates_info=templates_infos
)

print(new_project.point_df.head())
# print(raw_df1.columns.tolist())
