import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 初始化所有session状态
def init_session():
    states = {
        'processed': False,
        'output1': b'',
        'output2': b'',
        'alerts': [],
        'excluded': [],
        'special_cases': [],
        'file_uploaded': False
    }
    for key, value in states.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session()  # 确保初始化

# 页面结构
with st.container():
    st.title("入职周年分析系统")
    st.warning("""**重要说明**\n本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系""")

# 文件上传
uploaded_file = st.file_uploader("上传花名册（Excel格式）", type=["xlsx"], 
                               accept_multiple_files=False, key="file_upload")

# 参数选择
col1, col2 = st.columns(2)
with col1:
    year_range = list(range(2021, datetime.now().year+1))
    selected_year = st.selectbox("当前年份", options=year_range, 
                                index=len(year_range)-1, key="year_select")
with col2:
    selected_month = st.selectbox("目标月份", options=range(1,13), 
                                 format_func=lambda x: f"{x}月", key="month_select")

# 核心处理函数
def process_data():
    try:
        # 重置输出状态
        st.session_state.output1 = b''
        st.session_state.output2 = b''
        
        # 读取校验
        df = pd.read_excel(uploaded_file)
        required_cols = {'入职日期', '三级组织', '四级组织', '员工二级类别', 
                        '员工一级类别', '姓名', '花名'}
        if not required_cols.issubset(df.columns):
            missing = required_cols - set(df.columns)
            raise ValueError(f"缺少必要字段：{', '.join(missing)}")

        # 日期处理
        df['司龄开始日期'] = pd.to_datetime(df.get('司龄开始日期', pd.NaT))
        df['入职日期'] = pd.to_datetime(df['入职日期'])
        df['实际入职日期'] = df['司龄开始日期'].fillna(df['入职日期'])
        df['入职月份'] = df['实际入职日期'].dt.month

        # 执行筛选
        mask = (
            (df['入职月份'] == selected_month) &
            (~((df['三级组织'] == '财务中心') & 
             (df['四级组织'] == '证照支持部'))) &
            (df['员工二级类别'] == '正式员工')
        )
        filtered = df[mask].copy()
        
        # 记录排除人员
        st.session_state.excluded = df[~mask][['姓名', '员工二级类别', '三级组织']].to_dict('records')
        
        # 计算周年
        filtered['周年数'] = selected_year - filtered['实际入职日期'].dt.year
        filtered = filtered[filtered['周年数'] >= 1]
        
        # 生成明细表
        with io.BytesIO() as output:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered.to_excel(writer, index=False)
            st.session_state.output1 = output.getvalue()

        # 生成统计表
        filtered['人员信息'] = filtered.apply(
            lambda row: f"{row['三级组织']}-{row['姓名']}" + 
            (f"（{row['花名']}）" if pd.notna(row['花名']) else ""), axis=1)
        
        result = filtered.groupby('周年数').agg(人员列表=('人员信息', lambda x: '、'.join(x)))
        result = result.sort_values('周年数', ascending=False).reset_index()
        result['周年标签'] = result['周年数'].astype(str) + '周年'
        
        with io.BytesIO() as output:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result.to_excel(writer, index=False)
            st.session_state.output2 = output.getvalue()

        # 标记处理完成
        st.session_state.processed = True
        st.session_state.alerts = ["分析完成"]
        
    except Exception as e:
        st.error(f"处理失败：{str(e)}")
        st.session_state.processed = False

# 按钮操作区
col_btn1, col_btn2 = st.columns([1, 2])
with col_btn1:
    if st.button("▶️ 开始分析", disabled=not uploaded_file):
        process_data()
with col_btn2:
    if st.button("🔄 重新开始"):
        for key in ['processed', 'output1', 'output2', 'alerts', 'excluded', 'special_cases']:
            st.session_state[key] = init_session()[key]
        st.success("系统状态已重置")

# 结果展示区
if st.session_state.processed:
    with st.container():
        st.success("### 处理结果")
        
        # 下载功能
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="⬇️ 下载人员明细表",
                data=st.session_state.output1,
                file_name="符合条件人员列表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=len(st.session_state.output1) == 0  # 新增保护
            )
        with col_dl2:
            st.download_button(
                label="⬇️ 下载周年统计表",
                data=st.session_state.output2,
                file_name="入职周年统计表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=len(st.session_state.output2) == 0  # 新增保护
            )
        
        # 信息提示
        if st.session_state.excluded:
            st.warning(f"已排除 {len(st.session_state.excluded)} 名不符合条件人员")
        if '异动类型' in df.columns:  # 添加保护判断
            st.session_state.special_cases = filtered[filtered['异动类型'] == '活水']['姓名'].tolist()
            if st.session_state.special_cases:
                st.info(f"需人工核查活水人员：{', '.join(st.session_state.special_cases)}")
        if filtered['备注'].str.contains('注意外包人员').any():
            st.error("发现外包人员标记，请核查备注列")

# 持久化显示提示
for msg in st.session_state.alerts:
    st.success(msg)
