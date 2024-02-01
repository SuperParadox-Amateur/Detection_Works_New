from io import StringIO
import pathlib
from typing import List, Dict, Tuple, Any
from tkinter.filedialog import askopenfile

import ttkbootstrap as ttk
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.dialogs import Messagebox
# from ttkbootstrap.constants import BOTH, YES

import pandas as pd
from nptyping import DataFrame
from win32com.client import Dispatch

from occupational_health_module.new_occupational_health import NewOccupationalHealthItemInfo


class NewOccupationalHealth(ttk.Frame):
    '''职业卫生编号整理'''
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.pack(fill='both', expand=1)

        # 应用常量
        self.COLS: List[str] = [
            'ID',
            '委托编号',
            '样品类型',
            '样品编号',
            '送样编号',
            '样品名称',
            '检测参数',
            '采样/送样日期',
            '收样日期',
            '样品描述',
            '样品状态',
            '代表时长/h',
            '单元',
            '工种/岗位',
            '检测地点',
            '测点编号',
            '第几天',
            '第几个频次',
            '采样方式',
            '作业人数',
            '日接触时长/h',
            '周工作天数/d'
        ]

        # 应用变量
        _path: str = pathlib.Path().absolute().as_posix()
        # _path: str = 'E:/fxxk YiSaiTong/WT23ZDQ0063系统生成编号.xlsx'
        self.company_name = ttk.StringVar(value='')
        self.project_number = ttk.StringVar(value='')
        self.xlsx_path = ttk.StringVar(value=_path)
        self.rowdata: List[List[str]] = [['-'] * 22]

        # 职业卫生编号类
        self.raw_df: DataFrame = pd.DataFrame(columns=self.COLS)

        # 信息表单框架
        option_txt: str = '输入采样信息'
        self.option_lf = ttk.Labelframe(self, text=option_txt, padding=15)
        self.option_lf.pack(fill='x', expand=1, anchor='n')

        self.create_info_row()
        self.create_path_row()
        self.create_run_button_row()

        # 数据显示框架
        self.create_data_view()

        # 进度条
        self.progressbar = ttk.Progressbar(
            master=self,
            mode='indeterminate',
            bootstyle=('striped', 'success')
        )
        self.progressbar.pack(fill='x', expand=1)

        # 处理按钮
        self.create_handle_info_button()

    # 创建组件相关
    def create_info_row(self):
        '''向标题框架添加项目信息'''
        info_row = ttk.Frame(self.option_lf)
        info_row.pack(fill='x', expand=1, pady=15)
        # 公司名称输入
        company_name_lbl = ttk.Label(info_row, text='公司名称', width=15)
        company_name_lbl.pack(side='left', padx=(15, 0))
        company_name_ent = ttk.Entry(info_row, textvariable=self.company_name)
        company_name_ent.pack(side='left', fill='x', expand=1, padx=5)
        # 项目编号输入
        project_num_lbl = ttk.Label(info_row, text='项目编号', width=15)
        project_num_lbl.pack(side='left', padx=(15, 0))
        project_num_ent = ttk.Entry(info_row, textvariable=self.project_number)
        project_num_ent.pack(side='left', fill='x', expand=1, padx=5)

    def create_path_row(self):
        '''向标题框架添加文件路径'''
        path_row = ttk.Frame(self.option_lf)
        path_row.pack(fill='x', expand=1)
        path_lbl = ttk.Label(path_row, text='文件路径', width=15)
        path_lbl.pack(side='left', padx=(15, 0))
        path_ent = ttk.Entry(path_row, textvariable=self.xlsx_path)
        path_ent.pack(side='left', fill='x', expand=1, padx=5)
        browse_btn = ttk.Button(
            master=path_row,
            command=self.on_browse,
            text='选择',
            width=8
        )
        browse_btn.pack(side='left', padx=5)

    def create_run_button_row(self):
        '''向标题框架添加运行按钮'''
        run_btn_row = ttk.Frame(self.option_lf)
        run_btn_row.pack(fill='x', expand=1)
        run_btn = ttk.Button(
            master=run_btn_row,
            command=self.on_run,
            text='运行',
            width=8
        )
        run_btn.pack(side='left', padx=5)


    def create_data_view(self):
        '''表格数据显示'''
        self.datatable = Tableview(
            master=self,
            coldata=self.COLS,
            rowdata=self.rowdata,
            paginated=True,
            searchable=True,
            pagesize=20
        )
        self.datatable.pack(fill='both', expand=1, padx=10, pady=10)

    def create_handle_info_button(self):
        '''处理职业卫生编号按钮，默认不可见'''
        self.handle_info_btn = ttk.Button(
            master=self,
            command=self.handle_info,
            text='处理',
            width=8
        )

    # 按钮方法
    def on_browse(self):
        '''回调或者选择文件路径'''
        file = askopenfile(title='选择文件')
        if file:
            path: str = pathlib.Path(file.name).absolute().as_posix()
            self.xlsx_path.set(path)

    def on_run(self):
        '''运行按钮执行读取表格文件并显示数据'''
        # 路径是否是Excel文件
        is_excel_file: bool = (
            pathlib.Path(self.xlsx_path.get()).suffix in  ['.xls', '.xlsx']
        )
        # 路径是否是文件夹
        is_dir: bool = pathlib.Path(self.xlsx_path.get()).is_dir()

        if is_excel_file:
            self.progressbar.start(50)
            data: List[List[str]] = self.read_excel_file_pywin32(self.xlsx_path.get())
            # 选择的Excel文件的列名称是否符合
            is_cols_match: bool = data[0] == self.COLS
            if is_cols_match:
                self.rowdata = data[1:]
                self.datatable.build_table_data(rowdata=self.rowdata, coldata=self.COLS)
                copy_df: DataFrame = pd.DataFrame(data=data[1:], columns=data[0])
                self.raw_df: DataFrame = self.initialize_raw_df(copy_df)
                self.handle_info_btn.pack(side='right', padx=5)
            else:
                Messagebox.show_error('选择的Excel文件的列名称不符合', '错误')
            self.progressbar.stop()
        elif is_dir:
            Messagebox.show_error('路径为文件夹', '错误')
        else:
            Messagebox.show_error('路径不是Excel文件', '错误')

    def handle_info(self):
        '''将所有样品编号信息整理并写入模板'''
        self.occupational_health_info: NewOccupationalHealthItemInfo = NewOccupationalHealthItemInfo(
            self.project_number.get(),
            self.company_name.get(),
            self.raw_df
        )
        # print(self.raw_df.head())
        self.occupational_health_info.write_to_templates()

    # 自定义函数
    def str_nested_list(self, nested_list: List[Tuple[Any]]) -> List[List[str]]:
        '''将嵌套列表的所有项目转换为str样式'''
        details_value: List[List[str]] = []
        for raw_value in nested_list:
            value: List[Any] = [*raw_value]
            value_str: List[str] = list(map(str, value))
            clean_value_str: List[str] = [self.clean_str(s) for s in value_str]
            details_value.append(clean_value_str)
        return details_value

    def clean_str(self, string: str) -> str:
        '''清理数据'''
        na_list: List[str] = ['None', 'NA', 'Na', 'No']
        if string in na_list:
            return ''
        else:
            return string


    def read_excel_file_pywin32(self, file_path: str, app_name: str = 'ket') -> List[List[str]]:
        '''使用pywin32库读取excel文件'''
        xlapp = Dispatch(f'{app_name}.Application')
        xlapp.Visible = False
        target_wb = xlapp.Workbooks.Open(file_path)
        target_ws = target_wb.Worksheets(1)
        row = target_ws.UsedRange.Rows.Count
        col = target_ws.UsedRange.Columns.Count
        raw_details_value = list(
            target_ws.Range(target_ws.Cells(1, 1), target_ws.Cells(row, col)).Value
            )
        details_value: List[List[str]] = self.str_nested_list(raw_details_value)
        # details_df = pd.DataFrame(data=details_value[1:], columns=details_value[0])
        target_wb.Close()
        xlapp.Quit()

        # return details_df
        return details_value

    def nested_list_to_str(self, nested_list: List[List[str]]):
        '''嵌套列表转StringIO'''
        str_io = StringIO()
        for lst in nested_list:
            text = ','.join(lst)
            str_io.write(f'{text}\n')
        return str_io

    def initialize_raw_df(self, idf: DataFrame) -> DataFrame:
        '''初始化DataFrame的数据格式'''
        target_cols: Dict[str, Any] = {
            '测点编号': lambda x: int(float(x)),
            '第几天': lambda x: int(float(x)),
            '第几个频次': lambda x: int(float(x)),
            '日接触时长/h': float,
            '周工作天数/d': float
        }
        raw_df: DataFrame = idf.copy()

        for col, value in target_cols.items():
            raw_df[col] = raw_df[col].replace('', '0')
            raw_df[col] = raw_df[col].apply(value)
        return raw_df




if __name__ == '__main__':
    app = ttk.Window('Test App')
    app.state('zoomed')
    NewOccupationalHealth(app)
    app.mainloop()
