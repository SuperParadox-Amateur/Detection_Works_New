from nptyping import DataFrame#, Structure as S
import pandas as pd
from pandas.api.types import CategoricalDtype
# from typing import NewType


point_df_dtype: dict[str, type[int] | type[str] | type[float]] = {
        '采样点编号': int,
        '单元': str,
        '检测地点': str,
        '工种': str,
        '日接触时间': float,
        '检测项目': str,
        '采样数量/天': int,
        '采样天数': int,
        }

pernonnel_df_dtype: dict[str, type[int] | type[str] | type[float]] = {
    '采样点编号': int,
    '单元': str,
    '工种': str,
    '日接触时间': float,
    '检测项目': str,
    '采样数量/天': int,
    '采样天数': int,
}

# 考虑将采样日程改为“1|2|3”或者指定的“2”的方式，这样可以自定义部分只需要一天样品的空气检测项目的采样日程
# 问题：采样可能是1天的定期或者3天的评价，产生的存储dataframe的变量要如何命名，最后要如何展示在streamlit的多标签里

class OccupationalHealthItemInfo():
    def __init__(
            # todo 计划将项目基本信息以dict的形式存放
            self, 
            company_name: str, 
            project_number: str, 
            # working_days: float, # 取消工作天数，因为不同工种的工作天数可能不同
            point_info_df: DataFrame,
            personnel_info_df: DataFrame
            ) -> None:
        self.company_name: str = company_name
        self.project_number: str = project_number
        # self.working_days: float = working_days # 取消工作天数
        self.point_info_df: DataFrame = point_info_df
        self.personnel_info_df: DataFrame = personnel_info_df
        self.reference_info_df: DataFrame = self.get_occupational_health_item_reference()
        self.point_factor_order, self.personnel_factor_order = self.get_point_personnel_factors_order()
        self.sorted_df()
        self.get_detection_days()
        self.shedule_days: int = self.point_info_df['采样日程'].view(int).max()  # type: ignore

    def get_point_personnel_factors_order(self) -> tuple[CategoricalDtype, CategoricalDtype]:
        point_factor_list: list[str] = self.point_info_df['检测项目'].unique().tolist()  # type: ignore
        point_factor_list: list[str] = sorted(point_factor_list, key=lambda x: x.encode('gbk'))
        point_factor_order = CategoricalDtype(point_factor_list, ordered=True)
        personnel_factor_list: list[str] = self.personnel_info_df['检测项目'].unique().tolist()  # type: ignore
        personnel_factor_list: list[str] = sorted(personnel_factor_list, key=lambda x: x.encode('gbk'))
        personnel_factor_order = CategoricalDtype(personnel_factor_list, ordered=True)
        return point_factor_order, personnel_factor_order

    def get_occupational_health_item_reference(self) -> DataFrame:
        reference_path: str = r'info_files\检测项目信息.xlsx'
        reference_df_dtype: dict[str, type[str] | type[bool]] = {
            '检测项目': str,
            '样品收集器': str,
            '采样仪器': str,
            '收集方式': str,
            '是否需要空白': bool,
            '保存时间': str,
            '流量*时间': str,
            '备注': str,
        }
        reference_df: DataFrame = pd.read_excel(reference_path, dtype=reference_df_dtype)  # type: ignore
        return reference_df

    def get_detection_days(self) -> None:
        self.point_info_df['采样日程'] = self.point_info_df['采样天数'].apply(lambda x: list(range(1, x + 1))) # type: ignore
        self.point_info_df = self.point_info_df.explode('采样日程')
        self.personnel_info_df['采样日程'] = self.personnel_info_df['采样天数'].apply(lambda x: list(range(1, x + 1))) # type: ignore
        self.personnel_info_df = self.personnel_info_df.explode('采样日程', ignore_index=True)
        
    def sorted_df(self) -> None:
        # 先按照检测项目按照汉字拼音排序，再按照采样点编号排序，两者要共存
        # self.point_info_df['采样点编号'] = self.point_info_df['采样点编号'].astype(str)  # type: ignore
        # self.point_info_df = self.point_info_df.sort_values(by=['检测项目', '采样点编号'], key=lambda x: x.str.encode('gbk'), ignore_index=True)  # type: ignore
        # self.personnel_info_df = self.personnel_info_df.sort_values(by=['检测项目', '采样点编号'], key=lambda x: x.str.encode('gbk'), ignore_index=True)  # type: ignore

        # self.point_info_df = self.point_info_df.sort_values(by=['检测项目', '采样点编号'], key=lambda x: x.str.encode('gbk'), ignore_index=True, ascending=[True, True]) # type: ignore
        # self.personnel_info_df = self.personnel_info_df.sort_values(by=['检测项目'], key=lambda x: x.str.encode('gbk'), ignore_index=True, ascending=True)  # type: ignore
        self.point_info_df['检测项目'] = self.point_info_df['检测项目'].astype(self.point_factor_order)  # type: ignore
        self.point_info_df = self.point_info_df.sort_values(by=['检测项目', '采样点编号'], ascending=True, ignore_index=True)  # type: ignore
        self.personnel_info_df['检测项目'] = self.personnel_info_df['检测项目'].astype(self.personnel_factor_order)  # type: ignore
        self.personnel_info_df = self.personnel_info_df.sort_values(by=['检测项目', '采样点编号'], ascending=True, ignore_index=True)  # type: ignore
    
# todo 建立一个基于OccupationalHealthItemInfo类的子类，为OccupationalHealthItemInfo类下的每天检测信息的类
# todo 可能考虑取消子类，因为部分检测参数（例如物理因素、CO和CO2等只需要一天，完全可以放在一个整体里）

class SingleDayOccupationalHealthItemInfo(OccupationalHealthItemInfo):
    def __init__(
            self,
            company_name: str,
            project_number: str,
            # working_days: float,
            point_info_df: DataFrame,
            personnel_info_df: DataFrame,
            schedule_day: int = 1,
            engaged_num: int = 0
            ) -> None:
        super().__init__(company_name, project_number, point_info_df, personnel_info_df)
        self.schedule_day: int = schedule_day
        self.engaged_num: int = engaged_num
        self.query_str: str = f'采样日程 == {self.schedule_day}'
        self.current_point_info_df: DataFrame = self.point_info_df.query(self.query_str).reset_index()  # type: ignore
        self.current_personnel_info_df: DataFrame = self.personnel_info_df.query(self.query_str).reset_index()  # type: ignore

