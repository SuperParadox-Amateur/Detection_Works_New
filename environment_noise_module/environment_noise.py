from datetime import datetime, time
from typing import List, Dict, Any, Callable, Optional

import numpy as np
import pandas as pd
from nptyping import DataFrame, NDArray, Shape, Float

noise_degree_standard: Dict[str, Dict[str, int]] = {
    '0类': {
        '昼间': 50,
        '夜间': 40,
    },
    '1类': {
        '昼间': 55,
        '夜间': 45,
    },
    '2类': {
        '昼间': 60,
        '夜间': 50,
    },
    '3类': {
        '昼间': 65,
        '夜间': 55,
    },
    '4类': {
        '昼间': 70,
        '夜间': 55,
    },
}

distr_dict: Dict[str, Callable] = {
    "正态分布": np.random.normal,
    "拉普拉斯分布": np.random.laplace,
    "逻辑分布": np.random.logistic,
    "耿贝尔分布": np.random.gumbel,
}

class EnvNoise():
    '''随机环境噪声'''
    def __init__(
        self,
        in_df_distr: DataFrame,
        in_df_random: DataFrame,
        tag: str,
        noise_degree: str,
        in_noise_degree_standard: Dict[str, Dict[str, int]] = noise_degree_standard,
        in_distr_dict: Dict[str, Callable] = distr_dict,
        t_minute: int = 10,
        freq_weighting: str = 'A',
        count_ps: int = 10,
        time_weighting: str = 'F',
        leq_generator: str = 'original', # 有默认的“original”、“approximate”和“integral”
        ) -> None:
        '''定义'''
        self.noise_degree_standard: Dict[str, Dict[str, int]] = in_noise_degree_standard # 噪声等级标准
        self.distr_dict: Dict[str, Callable] = in_distr_dict # 分布方法字典
        self.in_df_distr: DataFrame = in_df_distr # 分布噪声数值df
        self.in_df_random: DataFrame = in_df_random # 随机噪声数值df
        self.tag: str = tag # 数据标签
        self.noise_degree: str = noise_degree # 噪声类别
        self.night_noise_limit: int = self.noise_degree_standard[self.noise_degree]['夜间'] # 夜间噪声限值
        self.t_minute: int = t_minute # 监测时长(min)
        self.freq_weighting: Optional[str] = freq_weighting # 频率计权方式
        self.count_ps: int = count_ps # 每秒监测次数
        self.size: int = self.t_minute * self.count_ps * 60 # 监测时长的所有监测次数
        self.time_weighting: str = time_weighting # 时间计权方式
        self.leq_generator: str = leq_generator # L50值生成方式，默认为近似值


    def create_distr_noise_info_df(self, distr_name: str):
        '''从分布噪声数值df生成随机分布噪声数值相关信息df'''
        distr_noise_info_list: List[Any] = []
        for i in range(self.in_df_distr.shape[0]):
            current_datetime = self.in_df_distr.loc[i, '日期时间']
            r1: int = int(self.in_df_distr.loc[i, '监测范围1']) # type: ignore
            r2: int = int(self.in_df_distr.loc[i, '监测范围2']) # type: ignore
            leq: float = float(self.in_df_distr.loc[i, '等效连续声级Leq']) # type: ignore
            sd_val: float = float(self.in_df_distr.loc[i, '标准差SD']) # type: ignore
            row_noise_info = self.create_distr_noise_info_dict(
                current_datetime, # type: ignore
                r1,
                r2,
                leq,
                sd_val,
                distr_name
            )
            distr_noise_info_list.append(row_noise_info)
        distr_noise_info_df: DataFrame = pd.DataFrame(data=distr_noise_info_list)
        return distr_noise_info_df


    def create_distr_noise_info_dict(
        self,
        current_datetime: datetime,
        r1: int,
        r2: int,
        leq: float,
        sd_val: float,
        distr_name: str
    ) -> Dict[str, Any]:
        '''从分布噪声数值生成随机分布噪声数值相关信息'''
        noise_array: NDArray[Shape[self.size], Float] = (
            self
            .distr_dict[distr_name]
            (leq, sd_val, self.size)
        )
        noise_info: Dict[str, Any] = self.create_noise_info_dict(noise_array)
        diff_val: float = float(10.0 * np.log10(self.t_minute * 60)) # 暴露声级和等效声级之间的差
        sel: float = leq + diff_val
        noise_info['r1'] = r1
        noise_info['r2'] = r2
        noise_info['dt'] = current_datetime
        noise_info['leq'] = leq
        noise_info['sel'] = sel
        noise_info['approximate_l50'] = self.get_approximate_l50(leq, noise_array)
        # 夜间值是否超标
        is_night: bool = (
            current_datetime >= datetime.combine(current_datetime.date(), time(22, 0, 0))
            or
            current_datetime <= datetime.combine(current_datetime.date(), time(6, 0, 0))
        )
        if is_night:
            if noise_info['lmax'] <= self.night_noise_limit + 15:
                noise_info['超过限值'] = False
            else:
                noise_info['超过限值'] = True
        else:
            noise_info['超过限值'] = None
        # 积分计算LEQ值知否小于L50
        is_integral_leq_gt_l50: bool = (
            noise_info['integral_leq'] >= noise_info['l50']
        )
        if is_integral_leq_gt_l50:
            noise_info['积分leq符合'] = True
        else:
            noise_info['积分leq符合'] = False

        return noise_info



    def create_random_noise_info_df(self):
        '''从分布噪声数值df生成随机分布噪声数值相关信息df'''
        random_noise_info_list: List[Any] = []
        for i in range(self.in_df_random.shape[0]):
            current_datetime = self.in_df_distr.loc[i, '日期时间']
            r1: int = int(self.in_df_random.loc[i, '监测范围1']) # type: ignore
            r2: int = int(self.in_df_random.loc[i, '监测范围2']) # type: ignore
            min_val: int = int(self.in_df_random.loc[i, '随机值下限']) # type: ignore
            max_val: int = int(self.in_df_random.loc[i, '随机值上限']) # type: ignore
            row_noise_info = self.create_random_noise_info_dict(
                current_datetime, # type: ignore
                min_val,
                max_val,
                r1,
                r2,
            )
            random_noise_info_list.append(row_noise_info)
        random_noise_info_df: DataFrame = pd.DataFrame(data=random_noise_info_list)
        return random_noise_info_df


    def create_random_noise_info_dict(
        self,
        current_datetime: datetime,
        min_val: int,
        max_val: int,
        r1: int,
        r2: int,
    ) -> Dict[str, Any]:
        '''从分布噪声数值生成随机分布噪声数值相关信息'''
        noise_array: NDArray[Shape[self.size], Float] = (
            np.random.randint(min_val * 10, max_val * 10, self.size) / 10
        )
        noise_info: Dict[str, Any] = self.create_noise_info_dict(noise_array)
        diff_val: float = float(10.0 * np.log10(self.t_minute * 60)) # 暴露声级和等效声级之间的差
        leq: float = float(np.mean(noise_array))
        sel: float = leq + diff_val
        noise_info['r1'] = r1
        noise_info['r2'] = r2
        noise_info['dt'] = current_datetime
        noise_info['leq'] = leq
        noise_info['sel'] = sel
        noise_info['approximate_l50'] = self.get_approximate_l50(leq, noise_array)
        is_night: bool = (
            current_datetime >= datetime.combine(current_datetime.date(), time(22, 0, 0))
            or
            current_datetime <= datetime.combine(current_datetime.date(), time(6, 0, 0))
        )
        if is_night:
            if noise_info['lmax'] <= self.night_noise_limit + 15:
                noise_info['超过限值'] = False
            else:
                noise_info['超过限值'] = True
        else:
            noise_info['超过限值'] = None

        # 积分计算LEQ值知否小于L50
        is_integral_leq_gt_l50: bool = (
            noise_info['integral_leq'] >= noise_info['l50']
        )
        if is_integral_leq_gt_l50:
            noise_info['积分leq符合'] = True
        else:
            noise_info['积分leq符合'] = False

        return noise_info


    def create_noise_info_dict(self, noise_array: NDArray) -> Dict[str, float]:
        '''从噪声数值信息生成噪声数值相关信息'''
        noise_info_dict: Dict[str, float] = {
            'lmax': float(noise_array.max()),
            'lmin': float(noise_array.min()), # 最小值和最大值
            'l5': float(np.percentile(noise_array, 95)),
            'l10': float(np.percentile(noise_array, 90)),
            'l50': float(np.percentile(noise_array, 50)),
            'l90': float(np.percentile(noise_array, 10)),
            'l95': float(np.percentile(noise_array, 5)),
            'sd': float(noise_array.std()),
            'approximate_leq': self.get_approximate_leq(noise_array),
            'integral_leq': self.get_integral_leq(noise_array),
        }
        return noise_info_dict


    def get_approximate_leq(self, noise_array: NDArray) -> float:
        '''求得噪声数值信息的近似LEQ值'''
        approximate_leq = float(
            np.percentile(noise_array, 50)
            + (np.percentile(noise_array, 90)
            - np.percentile(noise_array, 10)) ** 2 / 60
        )
        return approximate_leq


    def get_integral_leq(self, noise_array: NDArray) -> float:
        '''求得噪声数值信息的近似LEQ值'''
        integral_leq = float(
        10 * np.log10(
            np.sum(np.apply_along_axis(lambda x: 10 ** (0.1 * x), 0, noise_array)
                   * (1 / self.count_ps)) / (10 * 60)
            )
        )
        return integral_leq


    def get_approximate_l50(self, leq: float, noise_array: NDArray) -> float:
        '''求得噪声数值信息的近似L50值'''
        approximate_l50 = float(
            leq
            - (np.percentile(noise_array, 90) - np.percentile(noise_array, 10))
            ** 2 / 60
        )
        return approximate_l50


    def generate_noise_info_str(
            self,
            noise_info_dict: Dict[str, Any],
            # output_leq: str = 'original', # 有默认的“original”、“approximate”和“integral”
    ) -> str:
        '''从噪声信息dict生成噪声信息文本，并根据输出Leq类型输出其他对应信息'''
        # Leq和L50输出选择
        if self.leq_generator == 'approximate':
            leq_str: str = 'approximate_leq'
            l50_str: str = 'l50'
        elif self.leq_generator == 'integral':
            leq_str: str = 'integral_leq'
            l50_str: str = 'l50'
        else:
            leq_str: str = 'leq'
            l50_str: str = 'approximate_l50'

        current_info_list: List[str] = [
            f'{noise_info_dict["dt"]}',
            'Stat.-One',
            f'R: {noise_info_dict["r1"]}dB~{noise_info_dict["r2"]}dB Ts=00h{self.t_minute}m00s',
            f'Statistics: {self.freq_weighting} {self.time_weighting}',
            f'Leq,T= {noise_info_dict[leq_str]:.1f}dB SEL  = {noise_info_dict["sel"]:.1f}dB',
            f'Lmax = {noise_info_dict["lmax"]:.1f}dB Lmin = {noise_info_dict["lmin"]:.1f}dB',
            f'L5   = {noise_info_dict["l5"]:.1f}dB L10  = {noise_info_dict["l10"]:.1f}dB',
            f'L50  = {noise_info_dict[l50_str]:.1f}dB L90  = {noise_info_dict["l90"]:.1f}dB',
            f'L95  = {noise_info_dict["l95"]:.1f}dB SD   = {noise_info_dict["sd"]:>4.1f}dB',
            "\r\n"
        ]
        noise_string: str = "\r\n".join(current_info_list)
        return noise_string


    def generate_all_noise_info_str(self, in_df: DataFrame) -> str:
        '''从噪声数值df生成噪声信息文本'''
        noise_string_list: List[str] = []
        for i in range(in_df.shape[0]):
            noise_info_dict: Dict[str, Any] = in_df.loc[i].to_dict()
            noise_string: str = self.generate_noise_info_str(
                noise_info_dict = noise_info_dict,
            )
            noise_string_list.append(noise_string)

        return "\r\n".join(noise_string_list)
