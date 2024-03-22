from datetime import datetime, time, timedelta
from typing import Any, List

import pandas as pd
from nptyping import DataFrame


class SampleScheduleManage():
    '''采样时间安排'''
    def __init__(
            self,
            raw_work_df: DataFrame,
            instrument_df: DataFrame,
            break_time_df: DataFrame,
            time_span: int,
            instrument_time_len: int = 50
        ) -> None:
        self.time_span: int = time_span # 采样时间间隔
        self.instrument_time_len: int = instrument_time_len
        self.break_time_df: DataFrame = break_time_df # 休息时间范围
        self.instruments: DataFrame = self.initialize_instruments(instrument_df) # 仪器信息df
        self.work_df: DataFrame = self.initialize_work_df(raw_work_df) # 采样信息df
        self.sample_time_df: DataFrame = self.initialize_sample_time_df() # 生成仪器采样时间列表
        self.sample_schedule_df: DataFrame = pd.DataFrame() # 空df，用于保存采样点位和采样安排

    def initialize_work_df(self, raw_work_df: DataFrame) -> DataFrame:
        '''初始化采样信息df'''
        raw_work_df = raw_work_df.assign(是否完成=False)
        available_cols: List[str] = [
            '采样点编号',
            '单元',
            '检测地点',
            '工种',
            '日接触时间',
            '检测因素',
            '收集方式',
            '定点采样时间',
            '采样数量/天',
            '采样日程',
            '是否完成',
        ]
        work_df: DataFrame = (
            raw_work_df
            .sort_values(['采样点编号'])
            .reset_index(drop=True)
            [available_cols]
        )

        return work_df

    def initialize_sample_time_df(self) -> DataFrame:
        '''初始化采样时间df'''
        # 创建空的采样时间df
        sample_time_df: DataFrame = pd.DataFrame({
            '采样识别值': pd.Series(dtype=float),
            '小组': pd.Series(dtype=int),
            '次序': pd.Series(dtype=int),
            '采样日程': pd.Series(dtype=int),
            '采样时间': pd.Series(dtype='datetime64[ns]'),
        })
        # 获得所有采样日程
        days: List[int] = self.instruments['采样日程'].drop_duplicates().tolist()
        # 获得所有的小组
        groups: List[int] = self.instruments['小组'].drop_duplicates().tolist()
        # 每天每一小组都建立一个采样次序的df
        for day in days:
            for group in groups:
                # 当前小组最晚的仪器启动时间
                group_instrument_df: DataFrame = (
                    self.instruments
                    .query('小组 == @group and 采样日程 == @day')
                )
                boot_time: time = group_instrument_df['启动时间'].max().time()
                # boot_dt_time = datetime.combine(datetime.today(), boot_time)
                # 每个次序的间隔
                all_time_span: int = 15 + self.time_span
                time_interval: timedelta = timedelta(minutes=all_time_span)
                # 创建一个累加的时间列表
                time_list: List[datetime] = self.generate_time_list(boot_time, time_interval)

                group_df: DataFrame = pd.DataFrame({
                    '小组': [group] * len(time_list),
                    '次序': list(range(1, len(time_list) + 1)),
                    '采样时间': time_list,
                    '采样日程': day,
                })
                group_df['采样识别值'] = (
                    group_df['小组'].astype(float) * 0.1
                    + group_df['次序'] * 10
                    + (group_df['采样日程'] - 1) * 100
                )
                sample_time_df = pd.concat([sample_time_df, group_df], ignore_index=True)
        return sample_time_df

    def is_within_range(self, current_time: datetime) -> bool:
        '''判断时间是否在范围里'''
        judge_list: List[bool] = []
        for i in range(self.break_time_df.shape[0]):
            current_time_ts = current_time.time()#.timestamp()
            judge: bool = (
                current_time_ts >= self.break_time_df.loc[i, '开始时间']
                and
                current_time_ts <= self.break_time_df.loc[i, '结束时间']
            )
            judge_list.append(judge)
        if True in judge_list:
            return True
        else:
            return False


    def generate_time_list(self, boot_time: time, time_interval: timedelta) -> list[time]:
        '''生成一个累加的时间列表，保证不在休息时间范围内'''
        boot_time_dt: datetime = datetime.combine(datetime.today(), boot_time)
        time_list: list[time] = []
        for i in range(self.instrument_time_len):
            time_item: datetime = boot_time_dt + i * time_interval
            is_time_range: bool = self.is_within_range(time_item)
            if not is_time_range:
                time_list.append(time_item.time())
            else:
                pass
        return time_list


    def initialize_instruments(self, instrument_df: DataFrame) -> DataFrame:
        '''初始化仪器信息df'''
        instrument_df['端口'] = (
            instrument_df['端口数']
            .apply(lambda x: list(range(1, int(x) + 1)))
        )
        instrument_df['启动时间'] = (
            instrument_df['启动时间']
            .apply(lambda x: datetime.combine(datetime.today(), x))
            # .apply(lambda x: datetime.fromtimestamp(x))
        )
        instrument_df = (
            instrument_df
            .assign(
                是否完成=False,
                上一个采样点=0,
                采样次数=0
            )
        )

        return instrument_df.set_index('代号')

    def initialize_break_time_df(self, break_time_df: DataFrame):
        '''初始化休息时间范围df'''
        break_time_df['开始时间'] = (
            break_time_df['开始时间']
            .apply(lambda x: datetime.combine(datetime.today(), x))
        )
        break_time_df['结束时间'] = (
            break_time_df['结束时间']
            .apply(lambda x: datetime.combine(datetime.today(), x))
        )

    def judge_is_sample(self, instrument: str) -> None:
        '''判断该仪器是否可以继续采样'''
        gather_type: str = self.instruments.loc[instrument, '收集方式'] # type: ignore
        # 当前仪器可采样的点位数量
        remainder_rows_query_str: str = f'收集方式 == "{gather_type}" and 是否完成 == False'
        remainder_df: DataFrame = self.work_df.query(remainder_rows_query_str)
        remainder_rows: int = remainder_df.shape[0]
        # 仪器是否工作结束
        # is_finished: bool = self.instruments.loc[instrument, '是否完成'].value # type: ignore
        if remainder_rows == 0:
            self.instruments.loc[instrument, '是否完成'] = True
        else:
            pass

    def judge_is_sample_group(self, group: int):
        '''判断该小组的仪器是否可以继续采样'''
        # 小组的仪器可以采样的类型
        gather_types: list[str] = (
            self.instruments
            .query(f'小组 == "{group}"')['收集方式']
            .drop_duplicates()
            .tolist()
        )
        # 当前小组的仪器可采样的点位数量
        remainder_rows_query_str: str = f'收集方式 == {gather_types} and 是否完成 == False'
        remainder_df: DataFrame = self.work_df.query(remainder_rows_query_str)
        remainder_rows: int = remainder_df.shape[0]
        # 仪器是否工作结束
        if remainder_rows == 0:
            self.instruments.loc[self.instruments['小组'] == group, '是否完成'] = True
        else:
            pass


    def select_sample_point(self, instrument: str) -> int:
        '''为当前仪器选取采样点'''
        # [x] 计划增加从同一单元筛选出点位的功能
        gather_type: str = self.instruments.loc[instrument, '收集方式'] # type: ignore
        last_sample_point_num: int = self.instruments.loc[instrument, '上一个采样点'] # type: ignore
        # 如果上一个采样点不存在（即为0），则随机选取采样点
        if last_sample_point_num == 0:
            # 筛选出当前仪器可用的采样信息df
            new_point_query_str: str = f'收集方式 == "{gather_type}" and 是否完成 == False'
            all_sample_point_df: DataFrame = (
                self
                .work_df
                .query(new_point_query_str)
            )
            # 筛选出其中采样数量最多的行
            max_sample_point_rows: DataFrame = (
                all_sample_point_df
                .nlargest(1, '采样数量/天', keep='all')
            )
            random_row: DataFrame = max_sample_point_rows.sample(1)
            new_sample_point_num: int = random_row.iloc[0, 0] # type: ignore
            return new_sample_point_num
        else:
            # 上一个采样点所在的单元
            last_sample_unit: str = (
                self
                .work_df
                .query('采样点编号 == {last_sample_point_num}')
            ).iloc[0, 0]
            # 上一个采样点里当前仪器涉及的采样信息
            last_sample_point_query_str: str = f'采样点编号 == {last_sample_point_num} and 收集方式 == "{gather_type}" and 是否完成 == False'
            last_sample_point_df: DataFrame = (
                self
                .work_df
                .query(last_sample_point_query_str)
            )
            # 上一个采样单元里当前仪器涉及的采样信息
            last_sample_unit_query_str: str = f'采样点编号 == {last_sample_unit} and 收集方式 == "{gather_type}" and 是否完成 == False'
            last_sample_unit_df: DataFrame = (
                self
                .work_df
                .query(last_sample_unit_query_str)
            )
            # 上一个采样点可以让当前仪器采样的检测因素的数量
            last_sample_point_len: int = last_sample_point_df.shape[0]
            # 上一个采样单元可以让当前仪器采样的检测因素的数量
            last_sample_unit_len: int = last_sample_unit_df.shape[0]
            # 如果为0，则从该单元里重新选择点位，优先采样数量多的点位
            if last_sample_point_len == 0:
                new_unit_query_str: str = f'单元 == {last_sample_unit} and 收集方式 == "{gather_type}" and 是否完成 == False'
                all_sample_unit_df: DataFrame = (
                    self
                    .work_df
                    .query(new_unit_query_str)
                )
                # 筛选出其中采样数量最多的行
                max_sample_unit_rows: DataFrame = (
                    all_sample_unit_df
                    .nlargest(1, '采样数量/天', keep='all')
                )
                random_row: DataFrame = max_sample_point_rows.sample(1)
                new_sample_point_num: int = random_row.iloc[0, 0] # type: ignore
                return new_sample_point_num
            elif last_sample_unit_len == 0:
                # 筛选出当前仪器可用的采样信息df
                new_point_query_str: str = f'收集方式 == "{gather_type}" and 是否完成 == False'
                all_sample_point_df: DataFrame = (
                    self
                    .work_df
                    .query(new_point_query_str)
                )
                # 筛选出其中采样数量最多的行
                max_sample_point_rows: DataFrame = (
                    all_sample_point_df
                    .nlargest(1, '采样数量/天', keep='all')
                )
                random_row: DataFrame = max_sample_point_rows.sample(1)
                new_sample_point_num: int = random_row.iloc[0, 0] # type: ignore
                return new_sample_point_num
            else:
                return last_sample_point_num

    def instrument_sample(self, instrument: str) -> None:
        '''仪器采样'''
        gather_type: str = self.instruments.loc[instrument, '收集方式'] # type: ignore
        # 判断仪器能否继续采样
        self.judge_is_sample(instrument)
        is_finished: bool = self.instruments.loc[instrument, '是否完成'] # type: ignore
        if not is_finished:
            # 如果可以采样
            # 仪器采样次数加1
            order: int = self.instruments.loc[instrument, '采样次数'] # type: ignore
            group: int = self.instruments.loc[instrument, '小组'] # type: ignore
            self.instruments.loc[instrument, '采样次数'] = order + 1
            # 选择下一个采样点
            sample_point_num: int = self.select_sample_point(instrument)
            # 获得仪器开始采样的时间
            # sample_time = self.instruments.loc[instrument, '启动时间']
            # 获得采样点编号的采样信息df
            sample_point_str: str = (
                f'采样点编号 == {sample_point_num}'
                ' and '
                f' 收集方式 == "{gather_type}" and 是否完成 == False'
            )
            sample_point_df: DataFrame = self.work_df.query(sample_point_str)
            # 多个的仪器端口采样，并填写到采样信息df里
            if sample_point_df.shape[0] == 1:
                current_index = sample_point_df.iloc[0].name
                self.work_df.loc[current_index, '是否完成'] = True # type: ignore
                self.work_df.loc[current_index, '采样仪器'] = instrument # type: ignore
                self.work_df.loc[current_index, '次序'] = order + 1 # type: ignore
                self.work_df.loc[current_index, '小组'] = group # type: ignore
                # self.work_df.loc[current_index, '启动时间'] = sample_time # type: ignore
                self.work_df.loc[current_index, '端口'] = None # type: ignore
            else:
                ports: List[int] = self.instruments.loc[instrument, '端口'] # type: ignore
                for i, j in enumerate(ports):
                    current_index = sample_point_df.iloc[i].name
                    self.work_df.loc[current_index, '是否完成'] = True # type: ignore
                    self.work_df.loc[current_index, '采样仪器'] = instrument # type: ignore
                    self.work_df.loc[current_index, '次序'] = order + 1 # type: ignore
                    self.work_df.loc[current_index, '小组'] = group # type: ignore
                    # self.work_df.loc[current_index, '启动时间'] = sample_time # type: ignore
                    self.work_df.loc[current_index, '端口'] = j # type: ignore
            # all_time_span: int = 15 + self.time_span
            # self.instruments.loc[instrument, '启动时间'] = (
            #     sample_time
            #     + timedelta(minutes=all_time_span)
            # ) # type: ignore

    def group_sample(self, group: int):
        '''小组仪器采样'''
        # 小组的所有仪器
        group_instruments_df: DataFrame = (
            self.instruments
            .query(f'小组 == "{group}"')
            # .reset_index(drop=True)
        )
        # 小组的仪器可以采样的类型
        gather_types: list[str] = (
            group_instruments_df['收集方式']
            .drop_duplicates()
            .tolist()
        )
        # 判断仪器能否继续采样
        self.judge_is_sample_group(group)
        is_finished: bool = group_instruments_df.iloc[0]['是否完成']
        if not is_finished:
            # 如果可以采样
            # 仪器采样次数加1
            order: int = group_instruments_df.iloc[0]['采样次数']

            self.instruments.loc[self.instruments['小组'] == group, '采样次数'] = order + 1
            # 选择下一个采样点
            instrument: str = group_instruments_df.iloc[0].name
            sample_point_num: int = self.select_sample_point(instrument)
            


    def sample_work(self) -> None:
        '''开始采样工作'''
        while self.instruments.query('是否完成 == False').shape[0] > 0:
            instrument_list: List[str] = self.instruments.index.tolist()
            for instrument in instrument_list:
                self.instrument_sample(instrument)
        # self.merge_sample_order_and_time()

    # def merge_sample_order_and_time(self):
    #     '''合并采样点位顺序和时间'''
    #     sample_schedule_df: DataFrame = (
    #         self.work_df
    #         .assign(
    #             采样识别值 = lambda df: df['小组'] * 0.1 + df['次序'] + (df['采样日程'] - 1) * 100,
    #             偏移次数 = lambda df: df['次序'].max()
    #             )
    #     )
        # for i in range(3):
            # sample_schedule_df[f'开始{i + 1}'] =  sample_schedule_df['次序'] + sample_schedule_df['偏移次数']
            # self.sample_schedule_df: DataFrame = pd.merge(
            #     left=self.work_df,
            #     right=self.sample_time_df,
            #     on=['小组', '次序'],
            #     how='left'
            # )
    
    # def get_sample_order_time(self, order: float):
    #     '''获取和采样次序匹配的采样时间'''
    #     time = self.sample_time_df.query('采样识别值 == @order').loc[0, '采样时间']
    #     return time
