import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 初始化session状态
def init_session():
    session_defaults = {
        'processed': False,
        'output1': b'',
        'output2': b'',
        'alert_msgs': [],
        'excluded_count': 0,
        'outsource_count': 0,
        'huoshui_count': 0
    }
    for key, val in session_defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

init_session()

# 页面说明
st.title("入职周年分析系统")
st.warning("""**版本说明**\n本系统根据2025.4.4版业务规则开发，数据格式变更可能导致分析错误，请联系系统管理员""")

# 文件上传
uploaded_file = st.file_uploader("上传最新花名册", type=["xlsx"], 
                               help="请确保包含员工基础信息、入职日期、组织架构字段")

# 参数配置
col1, col2 = st.columns(2)
with col1:
    current_year = st.selectbox("基准年份", options=range(2021, datetime.now().year+1), 
                              index=3, format_func=lambda x: f"{x}年度")
with col2:
    target_month = st.selectbox("目标月份", options=range(1,13), 
                              format_func=lambda x: f"{x}月")

# 核心处理逻辑
def execute_analysis():
    try:
        # 重置状态
        st.session_state.update({
            'processed': False,
            'alert_msgs': [],
            'outsource_count': 0,
            'huoshui_count': 0
        })
        
        # 读取数据
        raw_df = pd.read_excel(uploaded_file)
        
        # 字段验证
        mandatory_fields = {'入职日期', '三级组织', '四级组织', '员工二级类别', '姓名'}
        if not mandatory_fields.issubset(raw_df.columns):
            missing = mandatory_fields - set(raw_df.columns)
            raise ValueError(f"缺少必要字段: {', '.join(missing)}")

        # 日期处理
        raw_df['计算基准日期'] = pd.to_datetime(
            raw_df.get('司龄开始日期', raw_df['入职日期'])
        )
        raw_df['入职月份'] = raw_df['计算基准日期'].dt.month
        
        # 基础筛选
        filtered = raw_df[
            (raw_df['入职月份'] == target_month) &
            (~((raw_df['三级组织'] == '财务中心') & 
             (raw_df['四级组织'] == '证照支持部'))) &
            (raw_df['员工二级类别'] == '正式员工')
        ].copy()
        
        # 计算周年
        filtered['周年数'] = current_year - filtered['计算基准日期'].dt.year
        valid_data = filtered[filtered['周年数'] >= 1]
        
        # 生成提醒信息
        st.session_state.outsource_count = valid_data['员工一级类别'].eq('外包').sum()
        st.session_state.huoshui_count = valid_data.get('异动类型', '').eq('活水').sum()
        
        # 生成输出文件
        with io.BytesIO() as buffer:
            with pd.ExcelWriter(buffer) as writer:
                valid_data.to_excel(writer, sheet_name='合格人员', index=False)
            st.session_state.output1 = buffer.getvalue()
        
        # 生成统计报表
        report_df = valid_data.groupby('周年数', as_index=False).agg(
            人数=('姓名', 'count'),
            人员列表=('姓名', lambda x: '、'.join(x))
        ).sort_values('周年数', ascending=False)
        report_df['周年标识'] = report_df['周年数'].astype(str) + '周年'
        
        with io.BytesIO() as buffer:
            with pd.ExcelWriter(buffer) as writer:
                report_df.to_excel(writer, sheet_name='周年统计', index=False)
            st.session_state.output2 = buffer.getvalue()
        
        # 更新状态
        st.session_state.processed = True
        st.session_state.alert_msgs.append("分析完成")
        
    except Exception as e:
        st.error(f"处理异常: {str(e)}")

# 操作按钮
col_act1, col_act2 = st.columns([1,3])
with col_act1:
    if st.button("开始分析", type="primary", disabled=not uploaded_file):
        execute_analysis()
with col_act2:
    if st.button("重置系统"):
        init_session()
        st.rerun()

# 结果展示
if st.session_state.processed:
    st.success("### 分析结果")
    
    # 文件下载
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="下载人员明细",
            data=st.session_state.output1,
            file_name="周年人员明细.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_dl2:
        st.download_button(
            label="下载统计报表",
            data=st.session_state.output2,
            file_name="周年统计报表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # 提醒信息
    st.markdown("---")
    st.caption("注意：系统不会自动过滤外包和活水人员，请人工检查下载文件中的相关记录")
