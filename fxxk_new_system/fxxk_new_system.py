import os
from typing import Any
import pandas as pd
from nptyping import DataFrame

class NewOccupationalHealthItemInfo():
    def __init__(
            self,
            project_code: str,
            company_name: str,
            raw_df: DataFrame
        ) -> None:
        self.company_name: str = company_name
        self.project_code: str = project_code
        self.df: DataFrame = self.initialize_df(raw_df)
        self.schedule_col: str = self.initialize_schedule()
        self.schedule_list: list[Any] = self.get_schedule_list()
        self.blank_df: DataFrame = self.initialize_blank_df()
        self.point_df: DataFrame = self.initialize_point_df()
    
    def initialize_df(self, raw_df: DataFrame) -> DataFrame:
        available_cols: list[str] = [
            '样品类型',
            '样品编号',
            '样品名称',
            '检测参数',
            '采样/送样日期',
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
        df['样品编号'] = df['样品编号'].apply(lambda x: x.replace(project_code, '')) # type: ignore
        return df
    
    def initialize_blank_df(self) -> DataFrame:
        # schedule: Any = self.schedule_list[schedule_index] # type: ignore
        # query_str: str = (
        #     f'{self.schedule_col} == @schedule'
        #     " and "
        #     f'样品类型 == "空白样"'
        # )
        raw_blank_df: DataFrame = (
            self # type: ignore
            .df
            .query('样品类型 == "空白样"')
            .reset_index(drop=True)
        )
        blank_df: DataFrame = raw_blank_df.pivot(
            index=['检测参数', self.schedule_col],
            columns='第几个频次',
            values='样品编号'
        ).rename(columns={1: '空白编号1', 2: '空白编号2'})
        return blank_df
    
    def initialize_point_df(self) -> DataFrame:
        query_str: str = (
            '样品类型 == "普通样"'
            ' and '
            '采样方式 == "定点"'
            ' and '
            '样品名称 != "工作场所物理因素"'
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
                    '采样/送样日期',
                    '第几天',
                    '日接触时长/h'
                ]
        )
        ['样品编号']
        .agg(list)
        .reset_index(drop=False)
        )
        groupby_point_df['样品数量'] = (
            groupby_point_df # type: ignore
            ['样品编号']
            .apply(len)
        )
        groupby_point_df['是否合并代表时长'] = (
            groupby_point_df # type: ignore
            .apply(lambda df: True if df['日接触时长/h'] / df['样品数量'] < 0.25 else False, axis=1) # type: ignore
        )
        point_df: DataFrame = groupby_point_df.merge( # type: ignore
            self.blank_df,
            on=['检测参数', '采样/送样日期'],
            how='left'
        )
        return point_df

    def initialize_schedule(self) -> str:
        if self.df['采样/送样日期'].isnull().all(): # type: ignore
            schedule_col: str = '第几天'
        else:
            schedule_col: str = '采样/送样日期'
        return schedule_col
    
    def get_schedule_list(self) -> list[Any]:
        schedule_list: list[Any] = (
            self
            .df[self.schedule_col]
            .drop_duplicates()
            .tolist()
        )
        return schedule_list

    def get_all_deleterious_substance_dict(self):
        pass

company_name: str = '万华化学（福建）有限公司'
project_code: str = '23ZKP0019'

file_path: str = './fxxk_new_system/WT23ZKP0019系统生成编号.xlsx'

available_cols: list[str] = [
    '样品类型',
    '样品编号',
    '样品名称',
    '检测参数',
    '采样/送样日期',
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

dtypes_dict: dict[str, type[str] | type[int] | type[float]] = {
    '样品类型': str,
    '样品编号': str,
    '样品名称': str,
    '检测参数': str,
    # '采样/送样日期': 'datetime',
    '单元': str,
    '工种/岗位': str,
    '检测地点': str,
    '测点编号': str,
    '第几天': int,
    '第几个频次': int,
    '采样方式': str,
    '作业人数': str,
    '日接触时长/h': float,
    '周工作天数/d': float,
}

df: DataFrame = pd.read_excel( # type: ignore
    os.path.abspath(file_path),
    sheet_name=0,
    usecols=available_cols,
    dtype=dtypes_dict,
    parse_dates=True
)

new_project = NewOccupationalHealthItemInfo(project_code, company_name, df)

schedule: Any = new_project.schedule_list[0] # type: ignore
query_str: str = (
    f'{new_project.schedule_col} == {schedule}'
    " and "
    '样品类型 == "空白样"'
)


print(new_project.blank_df)

new_project.point_df.to_clipboard(index=False)