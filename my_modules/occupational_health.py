from nptyping import DataFrame#, Structure as S
import numpy as np
import pandas as pd
from pandas.api.types import CategoricalDtype
# from typing import NewType


point_df_dtype: dict[str, type[int] | type[str] | type[float]] = {
        '采样点编号': int,
        '单元': str,
        '检测地点': str,
        '工种': str,
        '日接触时间': float,
        '检测因素': str,
        '采样数量/天': int,
        '采样天数': int,
        }

pernonnel_df_dtype: dict[str, type[int] | type[str] | type[float]] = {
    '采样点编号': int,
    '单元': str,
    '工种': str,
    '日接触时间': float,
    '检测因素': str,
    '采样数量/天': int,
    '采样天数': int,
}

def custom_sort(str_list: list[str], key_list: list[str]) -> list[str]:
    '''
    列表的自定义排序
    '''
    if str_list[0] in key_list:
        sorted_str_list: list[str] = sorted(str_list, key=lambda x: key_list.index(x))
        return sorted_str_list
    else:
        return str_list

# TODO 考虑将采样日程改为“1|2|3”或者指定的“2”的方式，这样可以自定义部分只需要一天样品的检测信息的采样日程
# 问题：采样可能是1天的定期或者3天的评价，产生的存储dataframe的变量要如何命名，最后要如何展示在streamlit的多标签里

class OccupationalHealthItemInfo():
    def __init__(
            # TODO 计划将项目基本信息以dict的形式存放
            self,
            company_name: str,
            project_number: str,
            # blank_info_df: DataFrame,
            point_info_df: DataFrame,
            personnel_info_df: DataFrame
            ) -> None:
        self.company_name: str = company_name
        self.project_number: str = project_number
        self.point_info_df: DataFrame = point_info_df
        self.personnel_info_df: DataFrame = personnel_info_df
        self.factor_reference_df: DataFrame = self.get_occupational_health_factor_reference()
        # self.point_factor_order, self.personnel_factor_order = self.get_point_personnel_factors_order()  # 应该不需要
        self.sort_df()
        self.get_detection_days()
        self.shedule_days: int = self.point_info_df['采样日程'].max()  # type: ignore
        self.point_deleterious_substance_df, self.personnel_deleterious_substance_df = self.get_deleterious_substance_df()
    
    # def get_point_personnel_factors_order(self) -> tuple[CategoricalDtype, CategoricalDtype]:
    #     '''
    #     （已废弃）将定点和个体检测信息中的检测因素按照汉字拼音排序，并导出CategoricalDtype
    #     '''
    #     point_factor_list: list[str] = self.point_info_df['检测因素'].unique().tolist()  # type: ignore
    #     point_factor_list: list[str] = sorted(point_factor_list, key=lambda x: x.encode('gbk'))
    #     point_factor_order = CategoricalDtype(point_factor_list, ordered=True)
    #     personnel_factor_list: list[str] = self.personnel_info_df['检测因素'].unique().tolist()  # type: ignore
    #     personnel_factor_list: list[str] = sorted(personnel_factor_list, key=lambda x: x.encode('gbk'))
    #     personnel_factor_order = CategoricalDtype(personnel_factor_list, ordered=True)
    #     return point_factor_order, personnel_factor_order

    def get_occupational_health_factor_reference(self) -> DataFrame:
        '''
        获得职业卫生所有检测因素的参考信息
        '''
        reference_path: str = './info_files/检测因素参考信息.xlsx'
        reference_df: DataFrame = pd.read_excel(reference_path)  # type: ignore
        return reference_df

    def get_detection_days(self) -> None:
        '''
        获得采样日程下每一天的检测信息
        '''
        self.point_info_df['采样日程'] = self.point_info_df['采样天数'].apply(lambda x: list(range(1, x + 1)))  # type: ignore
        self.point_info_df = self.point_info_df.explode('采样日程', ignore_index=True)
        self.point_info_df['采样日程'] = self.point_info_df['采样日程'].astype(int)  # type: ignore
        self.personnel_info_df['采样日程'] = self.personnel_info_df['采样天数'].apply(lambda x: list(range(1, x + 1)))  # type: ignore
        self.personnel_info_df = self.personnel_info_df.explode('采样日程', ignore_index=True)
        self.personnel_info_df['采样日程'] = self.personnel_info_df['采样日程'].astype(int)  # type: ignore
        
    def sort_df(self) -> None:
        '''
        对检测信息里的检测因素进行排序
        先对复合检测因素内部排序，再对所有检测因素
        '''
        # 将检测因素，尤其是复合检测因素分开转换为列表，并重新排序
        factor_reference_list: list[str] = self.factor_reference_df['标识检测因素'].tolist()
        self.point_info_df['检测因素'] = self.point_info_df['检测因素'].str.split('|').apply(custom_sort, args=(factor_reference_list,))  # type: ignore
        self.personnel_info_df['检测因素'] = self.personnel_info_df['检测因素'].str.split('|').apply(custom_sort, args=(factor_reference_list,))  # type: ignore
        # 将检测因素列表的第一个作为标识
        self.point_info_df['标识检测因素'] = self.point_info_df['检测因素'].apply(lambda lst: lst[0])  # type: ignore
        self.personnel_info_df['标识检测因素'] = self.personnel_info_df['检测因素'].apply(lambda lst: lst[0])  # type: ignore
        # 将检测因素列表合并为字符串
        self.point_info_df['检测因素'] = self.point_info_df['检测因素'].apply(lambda x: '|'.join(x))   # type: ignore
        self.personnel_info_df['检测因素'] = self.personnel_info_df['检测因素'].apply(lambda x: '|'.join(x))   # type: ignore
        # 将定点和个体的检测因素提取出来，创建CategoricalDtype数据，按照拼音排序
        point_factor_list: list[str] = self.point_info_df['检测因素'].unique().tolist()  # type: ignore
        point_factor_list: list[str] = sorted(point_factor_list, key=lambda x: x.encode('gbk'))
        point_factor_order = CategoricalDtype(point_factor_list, ordered=True)
        personnel_factor_list: list[str] = self.personnel_info_df['检测因素'].unique().tolist()  # type: ignore
        personnel_factor_list: list[str] = sorted(personnel_factor_list, key=lambda x: x.encode('gbk'))
        personnel_factor_order = CategoricalDtype(personnel_factor_list, ordered=True)
        # 将检测因素按照拼音排序
        self.point_info_df['检测因素'] = self.point_info_df['检测因素'].astype(point_factor_order)  # type: ignore
        self.point_info_df = self.point_info_df.sort_values(by=['检测因素', '采样点编号'], ascending=True, ignore_index=True)  # type: ignore
        self.personnel_info_df['检测因素'] = self.personnel_info_df['检测因素'].astype(personnel_factor_order)  # type: ignore
        self.personnel_info_df = self.personnel_info_df.sort_values(by=['检测因素', '采样点编号'], ascending=True, ignore_index=True)  # type: ignore

    def get_deleterious_substance_df(self) -> tuple[DataFrame, DataFrame]:
        '''
        获得所有空气有害物质的检测因素，包含定点和个体
        '''
        # （已废除）将参考信息里的所有空气有害物质检测因素转换为列表
        # deleterious_substance_factor_df: DataFrame = self.factor_reference_df.loc[self.factor_reference_df['收集方式'] != '直读']
        # deleterious_substance_factor_list: list[str] = deleterious_substance_factor_df['标识检测因素'].tolist()
        # （已废除）筛选出定点和个体检测信息里的含有所有空气有害物质检测因素的检测信息
        # point_deleterious_substance_df: DataFrame = self.point_info_df[self.point_info_df['标识检测因素'].isin(deleterious_substance_factor_list)]  # type: ignore
        # personnel_deleterious_substance_df: DataFrame = self.personnel_info_df[self.personnel_info_df['标识检测因素'].isin(deleterious_substance_factor_list)]  # type: ignore
        point_deleterious_substance_df: DataFrame = (
            pd.merge(  # type: ignore
                self.point_info_df,
                self.factor_reference_df[['标识检测因素', '是否仪器直读', '是否需要空白', '复合因素代码']],
                on='标识检测因素',
                how='left'
            )
            .fillna({'是否需要空白': False, '复合因素代码': 0, '是否仪器直读': False})
            .query('是否仪器直读 == False')
        )
        personnel_deleterious_substance_df: DataFrame = (
            pd.merge(  # type: ignore
                self.personnel_info_df,
                self.factor_reference_df[['标识检测因素', '是否仪器直读', '是否需要空白', '复合因素代码']],
                on='标识检测因素',
                how='left'
            )
            .fillna({'是否需要空白': False, '复合因素代码': 0, '是否仪器直读': False})
            .query('是否仪器直读 == False')
        )
        return point_deleterious_substance_df, personnel_deleterious_substance_df

    def get_single_day_deleterious_substance_df(self, schedule_day: int = 1) -> tuple[DataFrame, DataFrame]:
        '''
        获得一天的空气有害物质检测因素，包含定点和个体
        '''
        single_day_point_deleterious_substance_df: DataFrame = self.point_deleterious_substance_df[self.point_deleterious_substance_df['采样日程'] == schedule_day]
        single_day_personnel_deleterious_substance_df: DataFrame = self.personnel_deleterious_substance_df[self.personnel_deleterious_substance_df['采样日程'] == schedule_day]
        return single_day_point_deleterious_substance_df, single_day_personnel_deleterious_substance_df
        
    def get_single_day_blank_df(self, engaged_num: int = 0, schedule_day: int = 1) -> DataFrame:
        '''
        获得一天的空白样品编号
        '''
        # 复制定点和个体检测信息的dataframe，避免提示错误
        point_df, personnel_df = self.get_single_day_deleterious_substance_df(schedule_day)
        single_day_point_df: DataFrame = point_df.copy()
        single_day_personnel_df: DataFrame = personnel_df.copy()
        # 从定点和个体的dataframe提取检测因素，去重以及合并
        single_day_point_df['检测因素'] = single_day_point_df['检测因素'].str.split('|')  # type: ignore
        ex_single_day_point_df: DataFrame = single_day_point_df.explode('检测因素')
        single_day_personnel_df['检测因素'] = single_day_personnel_df['检测因素'].str.split('|')  # type: ignore
        ex_single_day_personnel_df: DataFrame = single_day_personnel_df.explode('检测因素')
        test_df: DataFrame = pd.concat(  # type: ignore
            [
                ex_single_day_point_df[['检测因素', '是否需要空白', '复合因素代码']],
                ex_single_day_personnel_df[['检测因素', '是否需要空白', '复合因素代码']]
            ]
        ).drop_duplicates('检测因素').reset_index(drop=True)
        # 分别处理非复合因素和复合因素，复合因素要合并。
        group1: DataFrame = test_df.loc[test_df['复合因素代码'] == 0, ['检测因素', '是否需要空白']]
        raw_group2: DataFrame = test_df.loc[test_df['复合因素代码'] != 0]
        group2 = pd.DataFrame(raw_group2.groupby(['复合因素代码'])['检测因素'].apply('|'.join)).reset_index(drop=True)  # type: ignore
        group2['是否需要空白'] = True
        # 最后合并，排序
        concat_group: DataFrame = pd.concat(  # type: ignore
            [group1, group2],
            ignore_index=True,
            axis=0,
            sort=False
        )
        blank_factor_list: list[str] = sorted(concat_group['检测因素'].tolist(), key=lambda x: x.encode('gbk'))  # type: ignore
        blank_factor_order = CategoricalDtype(categories=blank_factor_list, ordered=True)
        concat_group['检测因素'] = concat_group['检测因素'].astype(blank_factor_order)  # type: ignore
        # 筛选出需要空白编号的检测因素，并赋值
        single_day_blank_df: DataFrame = concat_group.loc[concat_group['是否需要空白'] == True].sort_values('检测因素', ignore_index=True)  # type: ignore
        single_day_blank_df["检测因素"] = single_day_blank_df["检测因素"].astype(str).map(lambda x: [x] + x.split("|") if x.count("|") > 0 else x)  # type: ignore
        single_day_blank_df["空白编号"] = np.arange(1, single_day_blank_df.shape[0] + 1) + engaged_num  # type: ignore
        single_day_blank_df.drop(columns=['是否需要空白'], inplace=True)  # type: ignore
        return single_day_blank_df.explode('检测因素')



# TODO 筛选某一天日程的所有定点和个体检测信息

# TODO 新建一个存储在bytesio里的excel文件
# TODO 考虑通过循环来将每日的定点和个体检测信息生成的数据（空白、定点和个体）存储到上述的excel文件
# TODO 获得某一天日程的空气检测信息的空白dataframe，筛选出需要空白的检测因素并生成其空白编号
# TODO 生成定点检测信息的样品编号范围，可以选择将爆炸样品编号的函数编写在这里
# TODO 更新已占用编号的函数
# TODO 每日的样品编号范围，用于流转单
# TODO 生成其他仪器只读的检测因素的记录表
# TODO 生成空白、定点和个体的记录表

# 建立一个基于OccupationalHealthItemInfo类的子类，为OccupationalHealthItemInfo类下的每天检测信息的类
# TODO 可能考虑取消子类，因为部分检测参数（例如物理因素、CO和CO2等只需要一天，完全可以放在一个整体里）

# class SingleDayOccupationalHealthItemInfo(OccupationalHealthItemInfo):
#     def __init__(
#             self,
#             company_name: str,
#             project_number: str,
#             # working_days: float,
#             point_info_df: DataFrame,
#             personnel_info_df: DataFrame,
#             schedule_day: int = 1,
#             engaged_num: int = 0
#             ) -> None:
#         super().__init__(company_name, project_number, point_info_df, personnel_info_df)
#         self.schedule_day: int = schedule_day
#         self.engaged_num: int = engaged_num
#         self.query_str: str = f'采样日程 == {self.schedule_day}'
#         self.current_point_info_df: DataFrame = self.point_info_df.query(self.query_str).reset_index()  # type: ignore
#         self.current_personnel_info_df: DataFrame = self.personnel_info_df.query(self.query_str).reset_index()  # type: ignore
        # self.standard_info_df# = super().get_standard_detection_item_info()
