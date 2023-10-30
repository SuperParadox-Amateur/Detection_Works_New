# import io
# import os
# import math
# from copy import deepcopy
# from docx import Document
# import openpyxl
import pandas as pd
from nptyping import DataFrame
# from pandas.api.types import CategoricalDtype
from my_modules.occupational_health import OccupationalHealthItemInfo#, refresh_engaged_num
from my_modules.other_infos import templates_info

company_name: str = 'MSCN'
project_name: str = '23ZXP0000'

file_path: str = r'./templates/项目信息试验模板.xlsx'
point_info_df: DataFrame = pd.read_excel(file_path, sheet_name='定点') # type: ignore
personnel_info_df: DataFrame = pd.read_excel(file_path, sheet_name='个体') # type: ignore

new_project = OccupationalHealthItemInfo(company_name, project_name, point_info_df, personnel_info_df, templates_info)

new_project.get_dfs_num(new_project.default_types_order)