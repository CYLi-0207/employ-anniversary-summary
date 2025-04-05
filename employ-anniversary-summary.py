import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 初始化session状态
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'output1' not in st.session_state:
    st.session_state.output1 = None
if 'output2' not in st.session_state:
    st.session_state.output2 = None
if 'alerts' not in st.session_state:
    st.session_state.alerts = []
if 'excluded' not in st.session_state:
    st.session_state.excluded = []
if 'special_cases' not in st.session_state:
    st.session_state.special_cases = []

# 页面说明容器
with st.container():
    st.title("入职周年分析系统")
    st.warning("""
    ​**重要说明**  
    本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系
    """)

# 文件上传容器
with st.container():
    st.header("数据准备")
    uploaded_file = st.file_uploader("上传花名册（Excel格式）", type=["xlsx"], accept_multiple_files=False)

# 参数选择容器
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        year_range = range(2021, datetime.now().year+1)
        selected_year = st.selectbox("当前年份", options=year_range, index=len(year_range)-1)
    with col2:
        selected_month = st.selectbox("目标月份", options=range(1,13), format_func=lambda x: f"{x}月")

# 处理函数
def process_data():
    try:
        # 重置状态
        st.session_state.alerts = []
        st.session_state.excluded = []
        st.session_state.special_cases = []
        
        # 读取数据
        df = pd.read_excel(uploaded_file)
        
        # 字段校验
        required_columns = {'入职日期', '三级组织', '四级组织', '员工二级类别', '员工一级类别', '姓名', '花名'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"缺少必要字段：{', '.join(missing)}")
        
        # 处理日期字段
        df['司龄开始日期'] = pd.to_datetime(df.get('司龄开始日期', pd.NaT))
        df['入职日期'] = pd.to_datetime(df['入职日期'])
        df['实际入职日期'] = df['司龄开始日期'].fillna(df['入职日期'])
        
        # 计算月份
        df['入职月份'] = df['实际入职日期'].dt.month
        
        # 筛选逻辑
        filtered = df[
            (df['入职月份'] == selected_month) &
            (~((df['三级组织'] == '财务中心') & (df['四级组织'] == '证照支持部'))) &
            (df['员工二级类别'] == '正式员工')
        ].copy()
        
        # 记录被排除人员
        excluded = df[~df.index.isin(filtered.index)]
        st.session_state.excluded = excluded[['姓名', '员工二级类别', '三级组织']].to_dict('records')
        
        # 计算周年数
        filtered['周年数'] = selected_year - filtered['实际入职日期'].dt.year
        filtered = filtered[filtered['周年数'] >= 1]
        
        # 添加备注
        filtered['备注'] = filtered['员工一级类别'].apply(
            lambda x: '注意外包人员' if x == '外包' else '')
        
        # 识别活水人员
        if '异动类型' in filtered.columns:
            st.session_state.special_cases = filtered[
                filtered['异动类型'] == '活水']['姓名'].tolist()
        
        # 生成明细表
        output1 = io.BytesIO()
        filtered.to_excel(output1, index=False)
        st.session_state.output1 = output1.getvalue()
        
        # 生成统计表
        filtered['人员信息'] = filtered.apply(
            lambda row: f"{row['三级组织']}-{row['姓名']}" + 
            (f"（{row['花名']}）" if pd.notna(row['花名']) else ""), axis=1)
        
        result = filtered.groupby('周年数').agg(人员列表=('人员信息', lambda x: '、'.join(x)))
        result = result.sort_values('周年数', ascending=False).reset_index()
        result['周年标签'] = result['周年数'].astype(str) + '周年'
        
        output2 = io.BytesIO()
        result.to_excel(output2, index=False)
        st.session_state.output2 = output2.getvalue()
        
        # 处理完成
        st.session_state.processed = True
        st.session_state.alerts.append("分析完成")
        
    except Exception as e:
        st.error(f"处理失败：{str(e)}")
        st.session_state.processed = False

# 按钮容器
with st.container():
    col1, col2 = st.columns([1,2])
    with col1:
        if st.button("开始分析", disabled=not uploaded_file):
            process_data()
    with col2:
        if st.button("重新开始"):
            st.session_state.processed = False
            st.session_state.output1 = None
            st.session_state.output2 = None
            st.session_state.alerts = []
            st.experimental_rerun()

# 结果展示容器
if st.session_state.processed:
    with st.container():
        st.success("#### 处理结果")
        
        # 下载区域
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="下载人员明细表",
                data=st.session_state.output1,
                file_name="符合条件人员列表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.download_button(
                label="下载周年统计表",
                data=st.session_state.output2,
                file_name="入职周年统计表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # 提醒信息
        if st.session_state.excluded:
            st.warning(f"已排除 {len(st.session_state.excluded)} 名不符合条件人员")
        if st.session_state.special_cases:
            st.info(f"需人工核查活水人员：{', '.join(st.session_state.special_cases)}")
        if any('注意外包人员' in row for row in st.session_state.output1):
            st.error("发现外包人员，请核查备注列")

# 持久化显示提醒
for msg in st.session_state.alerts:
    st.success(msg)
