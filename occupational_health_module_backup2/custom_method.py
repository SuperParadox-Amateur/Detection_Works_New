'''自定义方法'''
import os
from typing import Any, List, Dict
from decimal import Decimal, ROUND_HALF_DOWN

from new_occupational_health import NewOccupationalHealthItemInfo

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

def get_template_abs_path(self, templates_path_dict: Dict[str, str]) -> Dict[str, str]:
    '''获得模板的绝对路径'''
    templates_path_abs_dict: Dict[str, str] = {}
    for i, j in templates_path_dict.items():
        abs_path: str = os.path.join(
            os.path.abspath(os.path.join(os.getcwd(), "..")),
            j
        )
        templates_path_abs_dict[i] = abs_path
    return templates_path_abs_dict

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
