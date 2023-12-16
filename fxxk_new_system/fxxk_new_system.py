#%%
import os
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, List, Dict
import pandas as pd
from nptyping import DataFrame
from docx import Document

templates_path_dict: Dict[str, str] = {
    '有害物质定点': './templates/有害物质定点采样记录.docx',
    '有害物质个体': './templates/有害物质个体采样记录.docx',
    '高温定点': './templates/高温定点采样记录.docx',
    '一氧化碳定点': './templates/一氧化碳定点采样记录.docx',
    '噪声定点': './templates/噪声定点采样记录.docx',
    '噪声个体': './templates/噪声个体采样记录.docx',
}

#%%
class NewOccupationalHealthItemInfo():
    def __init__(
            self,
            project_number: str,
            company_name: str,
            templates_path_dict: Dict[str, str],
            raw_df: DataFrame
        ) -> None:
        self.company_name: str = company_name
        self.project_number: str = project_number
        self.templates_path_dict: Dict[str, str] = templates_path_dict
        self.df: DataFrame = self.initialize_df(raw_df)
        self.schedule_col: str = self.initialize_schedule()
        self.schedule_list: list[Any] = self.get_schedule_list()
        self.blank_df: DataFrame = self.initialize_blank_df()
        self.point_df: DataFrame = self.initialize_point_df()
        self.personnel_df: DataFrame = self.initialize_personnel_df()
        self.all_deleterious_substance_dict: Dict[Any, Any] = self.get_all_deleterious_substance_dict()
        self.output_path: str = os.path.join(
            os.path.expanduser("~/Desktop"),
            f'{self.project_number}记录表'
        )

# 初始化

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
        df['样品编号'] = df['样品编号'].apply(lambda x: x.replace(project_number, '')) # type: ignore
        return df
    
    def initialize_blank_df(self) -> DataFrame:
        '''初始化空白信息'''
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
    
    def initialize_point_df(self) -> DataFrame:
        '''初始化定点信息'''
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
        groupby_point_df['采样数量/天'] = (
            groupby_point_df # type: ignore
            ['样品编号']
            .apply(len)
        )
        groupby_point_df['是否合并代表时长'] = (
            groupby_point_df # type: ignore
            .apply(lambda df: True if df['日接触时长/h'] / df['采样数量/天'] < 0.25 else False, axis=1) # type: ignore
        )
        point_df: DataFrame = groupby_point_df.merge( # type: ignore
            self.blank_df,
            on=['检测参数', '采样/送样日期'],
            how='left'
        )
        point_df['代表时长'] = (
            point_df
            .apply(
                lambda df: self.get_exploded_contact_duration(
                    df['日接触时长/h'], df['采样数量/天'], 4
                ),
                axis=1
            )
        )

        return point_df

    def initialize_personnel_df(self) -> DataFrame:
        '''初始化个体信息'''
        query_str: str = (
            '样品类型 == "普通样"'
            ' and '
            '采样方式 == "个体"'
            ' and '
            '样品名称 != "工作场所物理因素"'
        )
        personnel_df: DataFrame = (
            self # type: ignore
            .df
            .query(query_str)
            .reset_index(drop=True)
        )
        return personnel_df

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

    def get_all_deleterious_substance_dict(self) -> Dict[Any, Any]:
        '''将每天的有害物质采样信息合并到一个字典中'''
        # 所有信息的字典
        all_deleterious_substance_dict = {}
        for i, schedule in enumerate(self.schedule_list):
            # 当日信息的字典
            deleterious_substance_dict = {}
            blank_df: DataFrame = (
                self
                .blank_df
                [self.blank_df[self.schedule_col] == schedule]
                # .query(f'{self.schedule_col} == @schedule')
                .sort_values(by=['空白编号1'])
                .reset_index(drop=True)
            )
            point_df: DataFrame = (
                self
                .point_df
                [self.point_df[self.schedule_col] == schedule]
                # .query(f'{self.schedule_col} == @schedule')
                .sort_values(by=['测点编号'])
                .reset_index(drop=True)
            )
            personnel_df: DataFrame = (
                self
                .personnel_df
                [self.personnel_df[self.schedule_col] == schedule]
                # .query(f'{self.schedule_col} == @schedule')
                .sort_values(by=['测点编号'])
                .reset_index(drop=True)
            )
            deleterious_substance_dict['空白'] = blank_df
            deleterious_substance_dict['定点'] = point_df
            deleterious_substance_dict['个体'] = personnel_df
            all_deleterious_substance_dict[i] = deleterious_substance_dict

        return all_deleterious_substance_dict

# 写入模板
    # def write_templates(self):
    #     '''将全部信息写入对应模板'''
    #     pass

    # def write_point_deleterious_substance(self):
        '''将定点有害物质信息写入模板'''
        # 获得模板
        # temp_path: str = self.templates_path_dict['有害物质定点']
        # doc = Document(temp_path)
        # for schedule in self.schedule_list:
        #     pass

# 自定义函数

    def get_exploded_contact_duration(
        self,
        duration: float,
        size: int,
        full_size: int
    ) -> List[float]:
        '''获得分开的接触时间，使用十进制来计算'''
        time_dec: Decimal = Decimal(str(duration))
        size_dec: Decimal = Decimal(str(size))
        time_list_dec: List[Decimal] = []
        if time_dec < Decimal('0.25') * size_dec:
            time_list_dec.append(time_dec)
        elif time_dec < Decimal('0.3') * size_dec:
            front_time_list_dec: List[Decimal] = [
                Decimal('0.25')] * (int(size) - 1)
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
            for _ in range(int(size) - 1):
                result: Decimal = (
                    judge_result
                    .quantize(
                        Decimal(prec_str),
                        ROUND_HALF_UP
                    )
                )
                time_list_dec.append(result)
            last_result: Decimal = time_dec - sum(time_list_dec)
            time_list_dec.append(last_result)

        time_list: List[float] = sorted(
            list(map(float, time_list_dec)),
            reverse=False
        )
        return time_list

#%%
company_name: str = '万华化学（福建）有限公司'
project_number: str = '23ZKP0019'

file_path: str = './WT23ZKP0019系统生成编号.xlsx'

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
    file_path,
    # os.path.abspath(file_path),
    sheet_name=0,
    usecols=available_cols,
    dtype=dtypes_dict,
    parse_dates=True
)

new_project = NewOccupationalHealthItemInfo(project_number, company_name, templates_path_dict, df)


# %%
temp_path = os.path.join(
    os.path.abspath(os.path.join(os.getcwd(), "..")),
    'templates/有害物质定点采样记录.docx'
)

doc = Document(temp_path)

# %%
