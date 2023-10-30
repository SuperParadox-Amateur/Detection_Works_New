from io import BytesIO
from typing import Union
import pandas as pd
from nptyping import DataFrame  # , Structure as S
import numpy as np
from pandas.core.series import Series
from decimal import Decimal, ROUND_HALF_UP

class OccupationalNoise():
    def __init__(
        self,
        noise_df: DataFrame,
        scale_value: float = 1.,
        size: int = 3,
        error_range: float = 0.,
        ) -> None:
        self.noise_df: DataFrame = noise_df
        self.scale_value: float = scale_value
        self.size: int = size
        self.error_range: float = error_range
    
    def generate_random_noise_value(self) -> DataFrame:
        error_value: float = round(self.error_range * np.random.uniform(-1, 1), 1)
        self.noise_df["校准值"] = self.noise_df["基准值"] + error_value
        noise_cols: list[str] = [f"第{i + 1}次" for i in list(range(self.size))]
        prev_noise_cols: list[str] = noise_cols[: -1]
        last_noise_col: str = noise_cols[-1]

        for col in prev_noise_cols:
            self.noise_df[col] = self.noise_df["校准值"].apply(lambda v: np.random.normal(v, self.scale_value))

        self.noise_df[last_noise_col] = (
            self.noise_df["校准值"] * self.size
            - self.noise_df[prev_noise_cols].apply(np.sum, axis=1)
        )
        self.noise_df[noise_cols] = self.noise_df[noise_cols].applymap(lambda x: np.round(x, 1))
        self.noise_df["平均值"] = self.noise_df[noise_cols].apply(np.mean, axis=1)
        self.noise_df["平均值"] = self.noise_df["平均值"].apply(lambda x: np.round(x, 1))

        all_col_names = self.noise_df.columns
        available_cols: list = [c for c in all_col_names if c not in ["基准值", "校准值"]]
        return self.noise_df[available_cols]

    def calculate_l_a_eq_8h(
        self,
        _noise_value: float,
        _duration: float
    ) -> float:
        '''
        目的:
            根据噪声采样信息dataframe中的噪声值和日接触时间计算噪声值的8小时等效声级
        参数:
            _noise_value: 噪声值
            _duration:    日接触时间
        返回:
            带有噪声值的8小时等效声级的噪声采样信息dataframe
        '''
        _la: float = 10 * np.log10(10 ** (_noise_value / 10) * _duration / 8)

        return _la
    #%% [markdown]

    # ### 计算8小时等效声级

    #%%
    def get_8h_equivalent_acoustical_level(self, _df: DataFrame) -> Series:
        '''
        目的:
            根据噪声采样信息dataframe中的噪声值的8小时等效声级
            要求五舍六入
        参数:
            _df: 噪声采样信息dataframe
        返回:
            带有噪声值的8小时等效声级的噪声采样信息dataframe
        '''
        if _df["日接触时间"] < 0.5:
            return None # pd.Series(pd.NA)
        elif _df["每周工作天数"] == 5.0:
            _equivalent_noise_value: Series = (
                pd.Series(self.calculate_l_a_eq_8h(_df["平均值"], _df["日接触时间"]))
                .apply(self.round_five_six, _scale_int=1)
            )
            return _equivalent_noise_value
        else:
            return None # pd.Series(pd.NA)
    #%% [markdown]

    # 计算40小时等效声级

    #%%
    def get_40h_equivalent_acoustical_value(self, _df: DataFrame) -> Series:
        '''
        目的:
          根据噪声采样信息dataframe中的噪声值的40小时等效声级
          要求五舍六入
        参数:
          _df:          噪声采样信息dataframe
        返回:
          带有噪声值的40小时等效声级的噪声采样信息dataframe
        '''
        if _df["日接触时间"] < 0.5:
            return None # pd.Series(pd.NA)
        elif _df["每周工作天数"] != 5.0:
            _assist_val: np.ndarray = (
                np.log10(
                    10 ** (0.1 * calculate_l_a_eq_8h(_df["平均值"], _df["日接触时间"]))
                    * (_df["每周工作天数"] / 5)
                )
                * 10
            )
            _equivalent_noise_value: Series = (
                pd.Series(_assist_val)
                .apply(self.round_five_six, _scale_int=1)
            )
            return _equivalent_noise_value
        else:
            return None # pd.Series(pd.NA)

    # %% 将数据保存到excel文件中

    def get_df_to_xlsx(self, _noise_df: DataFrame) -> bytes:
        '''
        目的:
          将噪声数据dataframe存放到byte数据中，以便之后下载
        参数:
          _noise_df: 噪声数据dataframe
        返回:
          存放噪声数据的byte数据
        '''
        _xlsxbyte: BytesIO = BytesIO()
        _noise_df.to_excel(_xlsxbyte, sheet_name="噪声", index=False)
        _xlsxbyte.seek(0, 0)
        return _xlsxbyte.read()

    def round_five_six(self, _input_num: Union[float, int], _scale_int: int = 1) -> float:
        '''
        参数
            _num:       要五舍六入的数
            _scale_int: 保留小数的位数，默认为1
        返回:
            经过五舍六入的数
        '''
        _num: float = float(_input_num)
        _corrected_value: float = 10 ** -(_scale_int + 1)
        _target_number: float = _num - _corrected_value
        _d_num: Decimal = (
            Decimal(f'{_target_number}')
            .quantize(Decimal(f'{10 ** -_scale_int}'), rounding=ROUND_HALF_UP)
        )
        return float(_d_num)
