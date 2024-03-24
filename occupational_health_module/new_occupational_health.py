'''处理系统生成编号，并写入模板里'''
import os
import json
from decimal import ROUND_HALF_DOWN, Decimal
from typing import Any, Dict, List

from nptyping import DataFrame
import pandas as pd
from pandas import CategoricalDtype


class NewOccupationalHealthItemInfo():
    '''系统生成的职业卫生样品编号处理信息'''
    def __init__(
            self,
            project_number: str,
            company_name: str,
            raw_df: DataFrame[Any],
            ) -> None:
        self.company_name: str = company_name  # 公司名称
        self.project_number: str = project_number  # 项目编号
        # self.templates_path_dict: Dict[str, Any] = (
        #     self.read_json_to_dict('./info_files/默认模板路径信息.json')
        # )
        self.templates_info_dict: Dict[str, Any] = (
            self.read_json_to_dict('./info_files/默认模板格式信息.json')
        )
        self.factor_reference_df: DataFrame[Any] = self.get_occupational_health_factor_reference()
        self.df: DataFrame[Any] = self.initialize_df(raw_df)
        self.schedule_col: str = self.initialize_schedule()
        self.schedule_list: List[Any] = self.get_schedule_list()
        self.blank_df: DataFrame[Any] = self.initialize_blank_df()
        self.point_df: DataFrame[Any] = self.initialize_point_df()
        self.personnel_df: DataFrame[Any] = self.initialize_personnel_df()
        self.stat_df: DataFrame[Any] = self.initialize_stat_df()


# 初始化

    # 默认获得职业卫生所有检测因素的参考信息
    def get_occupational_health_factor_reference(self) -> DataFrame[Any]:
        '''
        获得职业卫生所有检测因素的参考信息
        '''
        reference_path: str = os.path.join(
            'info_files/检测因素参考信息.csv'
        )
        reference_df: DataFrame[Any] = pd.read_csv(reference_path)  # type: ignore
        # 增加不同列的空值为不同的数值
        fill_dict: Dict[str, Any] = {
            '采样仪器': '/',
            '保存时间': '/',
            '流量*时间': '/',
            '定点采样流速': 0.0,
            '定点采样时间': 0,
            '个体采样流速': 0.0,
            '个体采样时间': 0,
        }
        reference_df: DataFrame[Any] = reference_df.fillna(fill_dict)  # type: ignore
        return reference_df

    def read_json_to_dict(self, json_file: str) -> Dict[str, Any]:
        '''读取json文件并转换为dict'''
        with open(json_file, 'r', encoding='utf-8') as f:
            string: str = f.read()
            jdict: Dict[str, Any] = json.loads(string)
        return jdict

    def initialize_df(self, raw_df: DataFrame[Any]) -> DataFrame[Any]:
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
        fillna_dict: Dict[str, Any] = {
            '日接触时长/h': 0.0,
            '周工作天数/d': 0.0,
        }
        cols_dtypes = {
            # '样品类型': str,
            # '样品编号': str,
            # '样品名称': str,
            # '检测参数': str,
            # '采样/送样日期': 'datetime64[ns]',
            # '单元': str,
            # '工种/岗位': str,
            # '检测地点': str,
            '测点编号': int,
            '第几天': int,
            '第几个频次': int,
            # '采样方式': str,
            # '作业人数': int,
            '日接触时长/h': float,
            '周工作天数/d': float,
        }
        df: DataFrame[Any] = raw_df[available_cols]
        df: DataFrame[Any] = df.fillna(fillna_dict) # type: ignore
        df: DataFrame[Any] = df.astype(cols_dtypes) # type: ignore
        #  将检测参数列转换为category类型用于排序
        # （取消，因为会导致groupby性能下降，甚至失败）
        # factor_list: List[str] = df['检测参数'].unique().tolist()
        # sorted_factor_list: List[str] = sorted(factor_list, key=lambda x: x.encode('gbk'))
        # factor_order = CategoricalDtype(sorted_factor_list, ordered=True)
        # df['检测参数'] = df['检测参数'].astype(factor_order)
        # df['样品编号'] = df['样品编号'].replace(self.project_number, '', inplace=True) # type: ignore
        df['样品编号'] = df['样品编号'].apply(self.handle_num_str) # type: ignore
        # df['样品编号'] = (
        #     df['样品编号']
        #     .apply(self.handle_num_str)
        #     .reset_index(drop=True)
        # ) # type: ignore

        return df

    def initialize_blank_df(self) -> DataFrame[Any]:
        '''初始化空白信息'''
        raw_blank_df: DataFrame[Any] = (
            self # type: ignore
            .df
            .query('样品类型 == "空白样"')
            .reset_index(drop=True)
        )
        blank_df: DataFrame[Any] = (
            raw_blank_df # type: ignore
            .pivot(
                index=['检测参数', self.schedule_col],
                columns='第几个频次',
                values='样品编号'
            )
            .rename(columns={1: '空白编号1', 2: '空白编号2'})
            .reset_index(drop=False)
        )

        return blank_df

    def initialize_point_df(self) -> DataFrame[Any]:
        '''初始化定点信息'''
        query_str: str = (
            '样品类型 == "普通样"'
            ' and '
            '采样方式 == "定点"'
            ' and '
            '样品名称 != "工作场所物理因素"'
            ' and '
            '样品描述 != "仪器直读"'
            ' and '
            '样品编号 != "/"'
        )
        raw_point_df: DataFrame[Any] = (
            self # type: ignore
            .df
            .query(query_str)
            .reset_index(drop=True)
        )
        raw_point_df['样品编号'] = (
            raw_point_df['样品编号'] # type: ignore
            .astype(int)
        )
        groupby_point_df: DataFrame[Any] = (
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
            .apply(
                lambda df: True if df[r'日接触时长/h'] / df['采样数量/天'] < 0.25 else False, # type: ignore
                axis=1
            )
        )
        point_df: DataFrame[Any] = groupby_point_df.merge( # type: ignore
            self.blank_df,
            on=['检测参数', self.schedule_col],
            how='left'
        )
        point_df['代表时长'] = (
            point_df # type: ignore
            .apply(
                lambda df: self.get_exploded_contact_duration( # type: ignore
                    df[r'日接触时长/h'], df[r'采样数量/天']#, 4 # type: ignore
                ),
                axis=1
            )
        )
        point_df['空白编号1'] = point_df['空白编号1'].fillna('-') # type: ignore
        point_df['空白编号2'] = point_df['空白编号2'].fillna('-') # type: ignore

        return point_df

    def initialize_personnel_df(self) -> DataFrame[Any]:
        '''初始化个体信息'''
        query_str: str = (
            '样品类型 == "普通样"'
            ' and '
            '采样方式 == "个体"'
            ' and '
            '样品名称 != "工作场所物理因素"'
            ' and '
            '样品编号 != "/"'
        )
        personnel_df: DataFrame[Any] = (
            self # type: ignore
            .df
            .query(query_str)
            .reset_index(drop=True)
        )
        # 去除空行
        new_personnel_df: DataFrame[Any] = (
            personnel_df # type: ignore
            .dropna(how='all')
            .reset_index(drop=True)
        )
        if not new_personnel_df.empty:
            new_personnel_df['样品编号'] = (
                new_personnel_df['样品编号'] # type: ignore
                .astype(int)
            )
        return new_personnel_df

    def initialize_stat_df(self) -> DataFrame[Any]:
        '''获得样本统计信息'''
        # 筛选出定点和个体的所有信息
        query_str: str = (
            '样品名称 != "工作场所物理因素"'
            ' and '
            '样品类型 == "普通样"'
            ' and '
            '样品编号 != "/"'
        )
        df: DataFrame[Any] = (
            self.df # type: ignore
            .query(query_str)
            .reset_index()
        )
        df['样品编号'] = df['样品编号'].astype(int) # 转为整数型 # type: ignore
        # 按照日程和检测参数分组并转换为列表
        groupby_df: DataFrame[Any] = (
            df # type: ignore
            .groupby([self.schedule_col, '检测参数'])
            ['样品编号']
            .agg(list)
            # .agg(self.convert_merge_range)
            .reset_index()
        )
        # 定点和个体的数量
        groupby_df['样品数量'] = groupby_df['样品编号'].apply(len) # type: ignore
        groupby_df['样品编号'] = groupby_df['样品编号'].apply(self.convert_merge_range) # type: ignore
        # 合并空白
        merged_df: DataFrame[Any] = groupby_df.merge( # type: ignore
            self.blank_df, # type: ignore
            on=['检测参数', '采样/送样日期'],
            how='left'
        )
        # 将定点和个体的编号与空白编号合并为一个列表
        merged_df['空白编号1'] = (
            merged_df['空白编号1'] # type: ignore
            .fillna('-')
        )
        merged_df['空白编号2'] = (
            merged_df['空白编号2'] # type: ignore
            .fillna('-')
        )
        # 空白样品数量
        merged_df['空白数量'] = (
            merged_df['空白编号1'] # type: ignore
            .apply(lambda x: 2 if x != '-' else 0) # type: ignore
        )
        merged_df['空白编号'] = (
            merged_df # type: ignore
            .apply(
                lambda x: [x['空白编号1'], x['空白编号2']], # type: ignore
                axis=1
            )
        )
        merged_df['全部样品编号'] = (
            merged_df # type: ignore
            .apply(
                lambda x: x['空白编号'] + x['样品编号'], # type: ignore
                axis=1
            )
        )
        # 去除无空白
        merged_df['全部样品编号'] = (
            merged_df['全部样品编号'] # type: ignore
            .apply(lambda x: [i for i in x if i != '-']) # type: ignore
        )
        # 所有样品数量
        merged_df['全部样品数量'] = (
            merged_df['空白数量']
            + merged_df['样品数量']
        )
        # 检测参数按照拼音排序
        factor_list: List[str] = (
            merged_df['检测参数'] # type: ignore
            .unique()
            .tolist()
        )
        sorted_factor_list: List[str] = sorted(
            factor_list,
            key=lambda x: x.encode('gbk')
        )
        factor_order: CategoricalDtype = CategoricalDtype(sorted_factor_list, ordered=True)
        merged_df['检测参数'] = merged_df['检测参数'].astype(factor_order) # type: ignore
        merged_df: DataFrame[Any] = (
            merged_df # type: ignore
            .sort_values(
                by=[self.schedule_col, '检测参数'],
                ascending=True,
                ignore_index=True
            )
        )
        # 加上检测参数的保存时间
        merged_df['标识检测参数'] = merged_df['检测参数'].apply(self.get_split_str_first) # type: ignore
        all_merged_df: DataFrame[Any] = (
            merged_df # type: ignore
            .merge(
                self.factor_reference_df,
                left_on='标识检测参数',
                right_on='标识检测因素',
                how='left'
            )
            .fillna({'保存时间': '/'})
        )

        return all_merged_df

    def initialize_schedule(self) -> str:
        '''初始化采样日程'''
        if self.df['采样/送样日期'].isnull().all(): # type: ignore
            schedule_col: str = '第几天'
        else:
            schedule_col: str = '采样/送样日期'
        return schedule_col

    def get_schedule_list(self) -> List[Any]:
        '''获得采样日程'''
        if self.schedule_col == '采样/送样日期':
            self.df[self.schedule_col] = pd.to_datetime(self.df[self.schedule_col]) # type: ignore
        # 可能是整数或者是日期
        schedule_list: List[Any] = (
            self.df[self.schedule_col] # type: ignore
            .drop_duplicates()
            .tolist()
        )
        # if self.schedule_col == '采样/送样日期':
        #     schedule_list = [
        #         datetime.strptime(i, '%Y-%m-%d').date() for i in schedule_list # type: ignore
        #     ]
        return schedule_list


# 写入模板文件中

# 自定义函数

    def parse_sample_num(self, sample_num: str) -> str:
        '''整理样品编号'''
        if sample_num != '/':
            parsed_sample_num: str = sample_num.replace(self.project_number, '')
            return parsed_sample_num
        else:
            return '/'


    def get_exploded_contact_duration(
        self,
        duration: float,
        size: int = 4
    ) -> List[float]:
        '''获得分开的接触时间，使用十进制来计算'''
        # 接触时间和数量转为十进制
        time_dec: Decimal = Decimal(str(duration))
        size_dec: Decimal = Decimal(str(size))
        time_list_dec: List[Decimal] = [] # 存放代表时长列表
        # 判断接触时间的小数位数
        if duration == int(duration):
            time_prec: int = 0
        else:
            time_prec: int = int(time_dec.as_tuple().exponent)
        # 确定基本平均值的小数位数
        time_prec_dec_dict: Dict[int, Decimal] = {
            0: Decimal('0'),
            -1: Decimal('0.0'),
            -2: Decimal('0.0')
        }
        prec_dec_str: Decimal = time_prec_dec_dict[time_prec]
        # 如果接触时间不能让每个代表时长大于0.25，则不分开
        if time_dec < Decimal('0.25') * size_dec:
            time_list_dec.append(time_dec)
        elif time_dec < Decimal('0.5') * size_dec:
            front_time_list_dec: List[Decimal] = [
                Decimal('0.25')] * (int(size) - 1)
            last_time_dec: Decimal = time_dec - sum(front_time_list_dec)
            time_list_dec.extend(front_time_list_dec)
            time_list_dec.append(last_time_dec)
        elif time_dec < Decimal('0.7') * size_dec:
            front_time_list_dec: List[Decimal] = [
                Decimal('0.5')] * (int(size) - 1)
        else:
            judge_result: Decimal = time_dec / size_dec
            for _ in range(int(size) - 1):
                result: Decimal = judge_result.quantize(prec_dec_str, ROUND_HALF_DOWN)
                time_list_dec.append(result)
            last_result: Decimal = time_dec - sum(time_list_dec)
            time_list_dec.append(last_result)
        time_list: List[float] = list(map(float, time_list_dec))
        return time_list

    def convert_merge_range(self, raw_lst: List[int]) -> List[str]:
        '''将编号列表里连续的编号合并，并转换为列表'''
        lst: List[int] = sorted(raw_lst)
        # lst: List[int] = [1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 13, 14, 15, 17, 18]
        all_range_list: List[List[int]] = []
        current_range: List[int] = []
        lst.extend([0])

        for i, num in enumerate(lst[:-1]):
            start: int = num
            current_range.append(start)
            end: int = num + 1
            if end == lst[i + 1]:
                # range.append(start)
                pass
            else:
                all_range_list.append(current_range)
                current_range = []

        range_str_list: List[str] = []
        for range_list in all_range_list:
            if len(range_list) != 1:
                range_str: str = f'{range_list[0]:>04d}--{range_list[-1]:>04d}'
                range_str_list.append(range_str)
            else:
                range_str: str = f'{range_list[0]:>04d}'
                range_str_list.append(range_str)

        return range_str_list

    def get_split_str_first(self, input_str: str) -> str:
        '''获得检测因素的标识检测因素'''
        str_list: List[str] = input_str.split('|')
        return str_list[0]

    def handle_num_str(self, num_str: str) -> str:
        '''（废除）'''
        if num_str != '/':
            new_num_str: str = num_str.replace(self.project_number, '')
            return new_num_str
        else:
            return '/'
