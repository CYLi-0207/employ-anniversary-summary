# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np  # 新增numpy库用于条件处理
from datetime import datetime
import io

def init_session():
    """初始化session_state"""
    keys = ['uploaded_file', 'filtered_df', 'result_df', 
           'error_msg', 'processed', 'special_exclusions',
           'livewater_employees']
    for key in keys:
        if key not in st.session_state:
            st.session_state[key] = None if key != 'processed' else False

# ================== 新增逻辑 ==================
def process_dates(df):
    """处理日期字段的核心逻辑"""
    # 转换日期格式
    df['入职日期'] = pd.to_datetime(df['入职日期'])
    df['司龄开始日期'] = pd.to_datetime(df['司龄开始日期'], errors='coerce')
    
    # 创建实际计算日期列（关键修改点）
    df['实际司龄日期'] = np.where(
        df['司龄开始日期'].notna(),
        df['司龄开始日期'],
        df['入职日期']
    )
    return df

def calculate_metrics(df, current_year, current_month):
    """计算关键指标"""
    # 计算入职月份（使用实际司龄日期）
    df['入职月份'] = df['实际司龄日期'].dt.month
    df['周年数'] = current_year - df['实际司龄日期'].dt.year
    return df

# ================== 修改后的处理流程 ==================
def process_data(uploaded_file, current_year, current_month):
    try:
        # 读取数据
        df = pd.read_excel(uploaded_file)
        
        # 增加字段校验（新增司龄开始日期字段）
        required_columns = {'入职日期', '三级组织', '四级组织', 
                          '员工二级类别', '员工一级类别', '司龄开始日期'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"缺失关键字段: {', '.join(missing)}")
        
        # 处理日期（新增处理步骤）
        df = process_dates(df)
        
        # 计算指标（修改为使用新日期字段）
        df = calculate_metrics(df, current_year, current_month)
        
        # 数据筛选
        exclusion_mask = (
            (df['入职月份'] == current_month)
            & (~((df['三级组织'] == '财务中心') & (df['四级组织'] == '证照支持部')))
            & (df['员工二级类别'] == '正式员工')
        )
        filtered_df = df[exclusion_mask].copy()
        
        # 存储特殊排除人数
        st.session_state.special_exclusions = len(df) - len(filtered_df)
        
        # 周年数筛选
        filtered_df = filtered_df[filtered_df['周年数'] >= 1]
        
        # 外包人员检测
        filtered_df['备注'] = filtered_df['员工一级类别'].apply(
            lambda x: '注意外包人员' if x == '外包' else '')
        
        # 活水员工检测（假设有员工类型字段）
        if '员工类型' in filtered_df.columns:
            livewater = filtered_df[filtered_df['员工类型'] == '活水']
            st.session_state.livewater_employees = livewater['姓名'].tolist()
        
        # 格式化展示信息（使用实际司龄日期）
        def format_info(row):
            base = f"{row['三级组织']}-{row['姓名']}"
            if pd.notna(row['花名']) and str(row['花名']).strip():
                return f"{base}（{row['花名']}）"
            return base
        
        filtered_df['人员信息'] = filtered_df.apply(format_info, axis=1)
        
        # 生成统计结果
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

# ================== 界面部分保持不变 ==================
# ... (保持原有界面代码不变，仅修改后台处理逻辑)

if __name__ == "__main__":
    init_session()
    st.title("员工入职周年分析系统")
    st.markdown("""**本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系**""")
    
    with st.container():
        uploaded_file = st.file_uploader("上传花名册(仅支持xlsx格式)", 
                                        type=['xlsx'],
                                        accept_multiple_files=False)
        
        col1, col2 = st.columns(2)
        with col1:
            year_range = range(2021, datetime.now().year + 1)
            selected_year = st.selectbox("当前年份", options=year_range, index=len(year_range)-1)
        with col2:
            selected_month = st.selectbox("当前月份", options=range(1,13), format_func=lambda x: f"{x}月")
        
        # 按钮组
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
    
    # 结果展示
    if st.session_state.processed and st.session_state.filtered_df is not None:
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
        
        st.warning("请人工检查以下情况：")
        st.write(f"1. 已排除特殊部门人员: {st.session_state.special_exclusions}人")
        if st.session_state.livewater_employees:
            st.write(f"2. 需要关注活水员工: {', '.join(st.session_state.livewater_employees)}")
    
    if st.session_state.error_msg:
        st.error(f"错误发生: {st.session_state.error_msg}")
