from datetime import time

import streamlit as st
import pandas as pd
from nptyping import DataFrame

from schedule_manage_module.schedule_manage import SampleScheduleManage

st.set_page_config(layout="wide", initial_sidebar_state="auto")

tab1, tab2,tab3 = st.tabs(['采样信息', '仪器信息', '休息时间'])

with tab1:
    i_raw_work_df: DataFrame = st.data_editor(
        pd.DataFrame([{
            '采样点编号': 1,
            '单元': '单元1',
            '检测地点': '检测地点1',
            '工种': '工种1',
            '日接触时间': 8,
            '检测因素': '监测因素1',
            '采样数量/天': 3,
            '收集方式': '大气',
            '定点采样时间': 15,
            '采样日程': 1,
        }]),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            '采样点编号': st.column_config.NumberColumn(format='%d'),
            # '单元': '单元1',
            # '检测地点': '检测地点1',
            # '工种': '工种1',
            '日接触时间': st.column_config.NumberColumn(format='%.2f'),
            '检测因素': '监测因素1',
            '采样数量/天': st.column_config.NumberColumn(format='%d'),
            '收集方式': st.column_config.SelectboxColumn('收集方式', options=['大气', '粉尘', '收集', '其他']),
            '定点采样时间': st.column_config.NumberColumn(format='%d'),
            '采样日程': st.column_config.NumberColumn(format='%d'),
        }
    )

with tab2:
    i_instrument_df: DataFrame = st.data_editor(
        pd.DataFrame([{
            '收集方式': '大气',
            '代号': 'Q1',
            '端口数': 2,
            '小组': 1,
            '采样日程': 1,
            '启动时间': time(8, 0, 0),
        }]),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            '收集方式': st.column_config.SelectboxColumn('收集方式', options=['大气', '粉尘', '收集', '其他']),
            '代号': st.column_config.TextColumn('代号'),
            '端口数': st.column_config.NumberColumn(format='%d'),
            '小组': st.column_config.NumberColumn(format='%d'),
            '采样日程': st.column_config.NumberColumn(format='%d'),
            '启动时间': st.column_config.TimeColumn(format='HH:mm:ss'),
        }
    )

with tab3:
    i_break_time_df: DataFrame = st.data_editor(
        pd.DataFrame([{
            '开始时间': time(12, 0, 0),
            '结束时间': time(13, 0 ,0)
        }]),
        num_rows='dynamic',
        use_container_width=True,
        column_config={
            '开始时间': st.column_config.TimeColumn('开始时间', format='HH:mm:ss'),
            '结束时间': st.column_config.TimeColumn('结束时间', format='HH:mm:ss')
        }
    )

i_time_span: float = st.number_input('采样间隔', min_value=0, value=3)

run: bool = st.button('运行', key='run')

if run:
    sample_schedule_manage: SampleScheduleManage = SampleScheduleManage(
        raw_work_df=i_raw_work_df,
        instrument_df=i_instrument_df,
        break_time_df=i_break_time_df,
        time_span=int(i_time_span)
    )
    sample_schedule_manage.sample_work()
    result_tab1, result_tab2, result_tab3 = st.tabs(['采样次序安排', '采样时间安排', '采样点位和时间安排'])
    with result_tab1:
        st.dataframe(sample_schedule_manage.work_df)
    with result_tab2:
        st.dataframe(sample_schedule_manage.sample_time_df)
    with result_tab3:
        st.dataframe(sample_schedule_manage.sample_schedule_df)
    # is_process: bool = st.button('处理', key='process')
    # if is_process:
    #     try:
    #         sample_schedule_manage.sample_work()
    #         st.success('完成')
    #         st.dataframe(sample_schedule_manage.work_df)
    #     except Exception:
    #         st.error('出现错误，无法进行')
