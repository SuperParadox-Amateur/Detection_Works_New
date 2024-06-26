from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from typing import Union, List, Optional
import pandas as pd
from nptyping import DataFrame  # , Structure as S
import numpy as np
from pandas.core.series import Series


class OccupationalNoiseInfo():
    '''职卫定点噪声的随机噪声值和等效噪声值'''

    def __init__(
        self,
        noise_df: DataFrame,
        scale_value: float = 1.,
        size: int = 3,
        # error_range: float = 0.,
    ) -> None:
        self.noise_df: DataFrame = noise_df
        self.scale_value: float = scale_value
        self.size: int = size
        self.new_noise_df: DataFrame = self.generate_random_noise_value()
        # self.error_range: float = error_range

    def generate_random_noise_value(
        self,
        # noise_value: float,
    ) -> DataFrame:
        '''生成随机噪声值'''
        noise_cols: List[str] = [f"第{i + 1}次" for i in range(self.size)]
        prev_noise_cols: List[str] = noise_cols[: -1]
        last_noise_col: str = noise_cols[-1]
        for col in prev_noise_cols:
            self.noise_df[col] = (
                self.noise_df["基准值"]
                .apply(lambda v: np.random.normal(v, self.scale_value))
                )
        self.noise_df[last_noise_col] = (
            self.noise_df["基准值"] * self.size
            - self.noise_df[prev_noise_cols].apply(np.sum, axis=1)
        )
        self.noise_df[noise_cols] = self.noise_df[noise_cols].applymap(lambda x: np.round(x, 1))
        self.noise_df["平均值"] = self.noise_df[noise_cols].apply(np.mean, axis=1)
        self.noise_df["平均值"] = self.noise_df["平均值"].apply(lambda x: np.round(x, 1))

        all_col_names = self.noise_df.columns
        available_cols: list = [c for c in all_col_names if c not in ["基准值"]]
        new_noise_df: DataFrame = self.noise_df[available_cols]

        new_noise_df["8小时等效"] = new_noise_df.apply( # type: ignore
            lambda df: self.get_8h_equivalent_acoustical_level( # type: ignore
                df['平均值'],
                df['日接触时间'],
                df['每周工作天数']
            ),
            axis=1
        )
        new_noise_df["40小时等效"] = new_noise_df.apply( # type: ignore
            lambda df: self.get_40h_equivalent_acoustical_value( # type: ignore
                df['平均值'],
                df['日接触时间'],
                df['每周工作天数']
            ),
            axis=1
        )
        return new_noise_df


        # prev_noise_list: List[float] = (
        #     np.random.normal(
        #         noise_value,
        #         self.scale_value,
        #         size=(self.size - 1)
        #     )
        #     .round(1)
        #     .tolist()
        # )
        # last_value: List[float] = [
        #     round(noise_value * self.size - sum(prev_noise_list), 1)]
        # random_noise_list: List[float] = prev_noise_list + last_value
        # return random_noise_list

    def calculate_l_a_eq_8h(
        self,
        noise_value: float,
        duration: float
    ) -> float:
        '''
        目的:
            根据噪声采样信息dataframe中的噪声值和日接触时间计算噪声值的8小时等效声级
        参数:
            noise_value: 噪声值
            duration:    日接触时间
        返回:
            带有噪声值的8小时等效声级的噪声采样信息dataframe
        '''
        la_value: float = 10 * \
            np.log10(10 ** (noise_value / 10) * duration / 8)

        return la_value
    # %% [markdown]

    # ### 计算8小时等效声级

    # %%
    def get_8h_equivalent_acoustical_level(self, noise_value: float, duration: float, workweek: float) -> Optional[float]:
        '''
        目的:
            根据噪声采样信息dataframe中的噪声值的8小时等效声级
            要求五舍六入
        参数:
            df: 噪声采样信息dataframe
        返回:
            带有噪声值的8小时等效声级的噪声采样信息dataframe
        '''
        if duration < 0.5:
            return None
        elif workweek == 5.0:
            equivalent_noise_value: float = self.round_five_six(
                self.calculate_l_a_eq_8h(noise_value, duration), 1)
            return equivalent_noise_value
        else:
            return None
    # %% [markdown]

    # 计算40小时等效声级

    # %%
    def get_40h_equivalent_acoustical_value(self, noise_value: float, duration: float, workweek: float) -> Optional[float]:
        '''
        目的:
          根据噪声采样信息dataframe中的噪声值的40小时等效声级
          要求五舍六入
        参数:
          df:          噪声采样信息dataframe
        返回:
          带有噪声值的40小时等效声级的噪声采样信息dataframe
        '''
        if duration < 0.5:
            return None
        elif workweek != 5.0:
            assist_value: float = (
                np.log10(
                    10 ** (0.1 * self.calculate_l_a_eq_8h(noise_value, duration))
                    * (workweek / 5)
                )
                * 10
            )
            equivalent_noise_value: float = self.round_five_six(
                assist_value, 1)
            return equivalent_noise_value
        else:
            return None

    # %% 将数据保存到excel文件中

    def get_df_to_xlsx(self, noise_df: DataFrame) -> bytes:
        '''
        目的:
          将噪声数据dataframe存放到byte数据中，以便之后下载
        参数:
          noise_df: 噪声数据dataframe
        返回:
          存放噪声数据的byte数据
        '''
        xlsxbyte: BytesIO = BytesIO()
        noise_df.to_excel(xlsxbyte, sheet_name="噪声", index=False)
        xlsxbyte.seek(0, 0)
        return xlsxbyte.read()

    def round_five_six(self, input_num: Union[float, int], scale_int: int = 1) -> float:
        '''
        参数
            num:       要五舍六入的数
            scale_int: 保留小数的位数，默认为1
        返回:
            经过五舍六入的数
        '''
        num: float = float(input_num)
        corrected_value: float = 10 ** -(scale_int + 1)
        target_number: float = num - corrected_value
        d_num: Decimal = (
            Decimal(f'{target_number}')
            .quantize(Decimal(f'{10 ** - scale_int}'), rounding=ROUND_HALF_UP)
        )
        return float(d_num)


# 测试
file_path: str = './templates/噪声值模板.csv'

df = pd.read_csv(file_path)
