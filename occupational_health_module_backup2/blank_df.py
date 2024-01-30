'''
空白dataframe类
'''
from typing import Dict, Any

from nptyping import DataFrame

from new_occupational_health import NewOccupationalHealthItemInfo

class BlankDataFrame(NewOccupationalHealthItemInfo):
    '''空白'''
    def __init__(
            self,
            project_number: str,
            company_name: str,
            df: DataFrame,
            templates_info: Dict[str, Dict[str, Any]],
    ) -> None:
        super(BlankDataFrame, self).__init__(project_number, company_name, df, templates_info)
        self.blank_df: DataFrame = self.initialize_blank_df()

    def initialize_blank_df(self) -> DataFrame:
        '''初始化空白信息'''
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
