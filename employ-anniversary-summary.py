# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ==============================
# 页面初始化与状态管理
# ==============================
def init_session():
    """初始化所有session_state变量"""
    if 'uploaded_file' not in st.session_state:
        st.session_state.uploaded_file = None
    if 'filtered_df' not in st.session_state:
        st.session_state.filtered_df = None
    if 'result_df' not in st.session_state:
        st.session_state.result_df = None
    if 'error_msg' not in st.session_state:
        st.session_state.error_msg = None
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    if 'special_exclusions' not in st.session_state:
        st.session_state.special_exclusions = 0
    if 'livewater_employees' not in st.session_state:
        st.session_state.livewater_employees = []

init_session()

# ==============================
# 页面静态内容
# ==============================
st.title("员工入职周年分析系统")
st.markdown("""**本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系**""")

# ==============================
# 核心处理函数
# ==============================
def process_data(uploaded_file, current_year, current_month):
    """主处理流程"""
    try:
        # 读取数据
        df = pd.read_excel(uploaded_file)
        
        # 校验关键字段
        required_columns = {'入职日期', '三级组织', '四级组织', '员工二级类别', '员工一级类别'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"缺失关键字段: {', '.join(missing)}")

        # 步骤1：数据预处理
        df['入职日期'] = pd.to_datetime(df['入职日期'])
        df['入职月份'] = df['入职日期'].dt.month

        # 步骤2：数据筛选
        exclusion_mask = (
            (df['入职月份'] == current_month)
            & (~((df['三级组织'] == '财务中心') & (df['四级组织'] == '证照支持部')))
            & (df['员工二级类别'] == '正式员工')
        )
        filtered_df = df[exclusion_mask].copy()

        # 计算特殊剔除人数
        st.session_state.special_exclusions = len(df) - len(filtered_df)

        # 步骤3：周年计算
        filtered_df['周年数'] = current_year - filtered_df['入职日期'].dt.year
        filtered_df = filtered_df[filtered_df['周年数'] >= 1]

        # 步骤4：备注处理
        filtered_df['备注'] = filtered_df['员工一级类别'].apply(
            lambda x: '注意外包人员' if x == '外包' else '')

        # 活水员工检测（假设存在'员工类型'字段）
        if '员工类型' in filtered_df.columns:
            livewater = filtered_df[filtered_df['员工类型'] == '活水']
            st.session_state.livewater_employees = livewater['姓名'].tolist()

        # 步骤5：信息格式化
        def format_info(row):
            base = f"{row['三级组织']}-{row['姓名']}"
            return f"{base}（{row['花名']}）" if pd.notna(row['花名']) else base

        filtered_df['人员信息'] = filtered_df.apply(format_info, axis=1)

        # 步骤6：统计汇总
        result_df = (
            filtered_df.groupby('周年数', as_index=False)
            .agg(人员信息=('人员信息', lambda x: '、'.join(x)))
            .sort_values('周年数', ascending=False)
        )
        result_df['周年标签'] = result_df['周年数'].astype(str) + '周年'

        return filtered_df, result_df[['周年数', '周年标签', '人员信息']]

    except Exception as e:
        st.session_state.error_msg = str(e)
        return None, None

# ==============================
# 交互组件区域
# ==============================
with st.container():
    # 文件上传组件
    uploaded_file = st.file_uploader("上传花名册(仅支持xlsx格式)", 
                                    type=['xlsx'],
                                    accept_multiple_files=False)

    # 时间选择组件
    col1, col2 = st.columns(2)
    with col1:
        year_range = range(2021, datetime.now().year + 1)
        selected_year = st.selectbox("当前年份", options=year_range, index=len(year_range)-1)
    with col2:
        selected_month = st.selectbox("当前月份", options=range(1,13), format_func=lambda x: f"{x}月")

    # 功能按钮组
    btn_col1, btn_col2 = st.columns([1,2])
    with btn_col1:
        if st.button("🚀 开始分析", use_container_width=True):
            if uploaded_file is None:
                st.session_state.error_msg = "请先上传数据文件"
            else:
                st.session_state.filtered_df, st.session_state.result_df = process_data(
                    uploaded_file, selected_year, selected_month)
                st.session_state.processed = True

    with btn_col2:
        if st.button("🔄 重新开始", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            init_session()

# ==============================
# 结果展示与下载
# ==============================
if st.session_state.processed and st.session_state.filtered_df is not None:
    # 下载功能实现
    buffer1 = io.BytesIO()
    st.session_state.filtered_df.to_excel(buffer1, index=False)
    
    buffer2 = io.BytesIO()
    st.session_state.result_df.to_excel(buffer2, index=False)

    st.success("分析完成！请下载结果文件")
    st.download_button(label="📥 下载符合条件人员列表",
                       data=buffer1,
                       file_name='符合条件人员列表.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    st.download_button(label="📥 下载入职周年统计表",
                       data=buffer2,
                       file_name='入职周年统计表.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # 特别提醒显示
    st.warning("请人工检查以下情况：")
    st.write(f"1. 已排除特殊部门人员: {st.session_state.special_exclusions}人")
    if st.session_state.livewater_employees:
        st.write(f"2. 需要关注活水员工: {', '.join(st.session_state.livewater_employees)}")

# ==============================
# 错误处理显示
# ==============================
if st.session_state.error_msg:
    st.error(f"错误发生: {st.session_state.error_msg}")
