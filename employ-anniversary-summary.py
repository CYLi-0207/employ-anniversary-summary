import streamlit as st
import pandas as pd
from io import BytesIO

# 初始化session状态
def init_session():
    reset_keys = ['processed', 'result1', 'result2', 'messages', 'excluded']
    for key in reset_keys:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state.setdefault('messages', [])
    st.session_state.setdefault('excluded', [])

# 页面配置
st.set_page_config(
    page_title="员工入职周年分析系统",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 页面说明
st.markdown("""
**版本说明**  
本系统根据2025.4.4版本的花名册数据规范开发，数据格式变更可能导致分析错误，请联系系统管理员。
""")

# 重新开始按钮
if st.button("🔄 重新开始"):
    init_session()
    st.experimental_rerun()

# 文件上传
uploaded_file = st.file_uploader(
    "📤 上传最新花名册Excel文件",
    type=["xlsx"],
    key='file_uploader',
    help="请确保文件包含正确的员工信息字段"
)

# 参数选择
col1, col2 = st.columns(2)
with col1:
    selected_year = st.selectbox(
        "🗓️ 当前年份",
        options=range(2021, pd.Timestamp.now().year + 1),
        format_func=lambda x: f"{x}年"
    )
with col2:
    selected_month = st.selectbox(
        "📆 当前月份",
        options=range(1, 13),
        format_func=lambda x: f"{x}月",
        index=2  # 默认选择3月
    )

# 数据处理函数
def process_data(df, year, month):
    try:
        # 字段验证
        required_columns = ['入职日期', '三级组织', '四级组织', 
                          '员工二级类别', '员工一级类别', '姓名', '花名']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"缺失必要字段: {', '.join(missing)}")

        # 数据处理
        df['入职日期'] = pd.to_datetime(df['入职日期'], errors='coerce')
        if df['入职日期'].isnull().any():
            raise ValueError("存在无效的日期格式")
            
        df['入职月份'] = df['入职日期'].dt.month
        
        # 筛选逻辑
        condition = (
            (df['入职月份'] == month) &
            (~((df['三级组织'] == '财务中心') & (df['四级组织'] == '证照支持部'))) &
            (df['员工二级类别'] == '正式员工')
        )
        filtered_df = df[condition].copy()
        
        # 记录排除人员
        excluded = df[~condition][['姓名', '三级组织', '四级组织']]
        st.session_state.excluded = excluded.to_dict('records')
        
        # 计算周年
        filtered_df['周年数'] = year - filtered_df['入职日期'].dt.year
        filtered_df = filtered_df[filtered_df['周年数'] >= 1]
        
        # 生成备注
        filtered_df['备注'] = filtered_df['员工一级类别'].apply(
            lambda x: '注意外包人员' if x == '外包' else '')
            
        # 格式化显示
        filtered_df['人员信息'] = filtered_df.apply(
            lambda r: f"{r['三级组织']}-{r['姓名']}（{r['花名']}）" 
            if pd.notna(r['花名']) else f"{r['三级组织']}-{r['姓名']}", axis=1)

        # 生成统计
        summary = (
            filtered_df.groupby('周年数', as_index=False)
            .agg(人员信息=('人员信息', '、'.join))
            .sort_values('周年数', ascending=False)
        )
        summary['周年标签'] = summary['周年数'].astype(str) + '周年'
        
        return filtered_df, summary[['周年数', '周年标签', '人员信息']]

    except Exception as e:
        st.error(f"数据处理错误: {str(e)}")
        raise

# 执行分析
if st.button("🚀 开始分析", type="primary") and uploaded_file:
    init_session()
    try:
        df = pd.read_excel(uploaded_file)
        filtered_df, summary_df = process_data(df, selected_year, selected_month)
        
        # 保存结果到内存
        output1 = BytesIO()
        with pd.ExcelWriter(output1, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False)
        output1.seek(0)
        
        output2 = BytesIO()
        summary_df.to_excel(output2, index=False)
        output2.seek(0)
        
        # 更新状态
        st.session_state.update({
            'processed': True,
            'result1': output1,
            'result2': output2,
            'messages': [
                f"✅ 原始数据记录数: {len(df)}",
                f"✅ 符合条件人员数: {len(filtered_df)}",
                f"✅ 排除人员数: {len(st.session_state.excluded)}",
                "⚠️ 存在需注意的外包人员" if any(filtered_df['备注'] != '') else ""
            ]
        })
        
    except Exception as e:
        st.error(f"分析失败: {str(e)}")
        st.session_state.processed = False

# 结果显示
if st.session_state.get('processed'):
    st.success("## 分析完成")
    
    # 统计信息
    st.markdown("### 处理摘要")
    for msg in st.session_state.messages:
        st.write(msg)
    
    # 排除人员
    if st.session_state.excluded:
        st.markdown("### 特殊排除人员")
        st.dataframe(pd.DataFrame(st.session_state.excluded))
    
    # 下载按钮
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📥 下载人员清单",
            data=st.session_state.result1,
            file_name="周年人员清单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.download_button(
            label="📊 下载统计报表",
            data=st.session_state.result2,
            file_name="周年统计报表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
