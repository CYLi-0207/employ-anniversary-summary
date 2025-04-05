import streamlit as st
import pandas as pd
from datetime import datetime
import io
from io import BytesIO

# 初始化session状态
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'result1' not in st.session_state:
    st.session_state.result1 = None
if 'result2' not in st.session_state:
    st.session_state.result2 = None
if 'messages' not in st.session_state:
    st.session_state.messages = []
if 'excluded' not in st.session_state:
    st.session_state.excluded = []

# 页面布局设置
st.set_page_config(page_title="入职周年分析系统", layout="wide")

# ========================
# 页面说明部分
# ========================
st.markdown("""
**说明：​**  
本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系
""")

# ========================
# 重新开始按钮
# ========================
if st.button("重新开始"):
    st.session_state.processed = False
    st.session_state.result1 = None
    st.session_state.result2 = None
    st.session_state.messages = []
    st.session_state.excluded = []
    st.experimental_rerun()

# ========================
# 文件上传区域
# ========================
uploaded_file = st.file_uploader("上传花名册文件", type=["xlsx"], 
                                help="每次只能上传一份Excel文件")

# ========================
# 参数选择区域
# ========================
col1, col2 = st.columns(2)
with col1:
    selected_year = st.selectbox("当前年份", options=range(2021, datetime.now().year+1), 
                               format_func=lambda x: f"{x}年")
with col2:
    selected_month = st.selectbox("当前月份", options=range(1,13), 
                                format_func=lambda x: f"{x}月")

# ========================
# 处理逻辑函数
# ========================
def process_data(df, year, month):
    try:
        # 校验必要字段
        required_columns = ['入职日期', '三级组织', '四级组织', 
                          '员工二级类别', '员工一级类别', '姓名', '花名']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"缺失必要字段: {', '.join(missing)}")

        # 处理日期字段
        df['入职日期'] = pd.to_datetime(df['入职日期'])
        df['入职月份'] = df['入职日期'].dt.month

        # 原始数据备份
        original_count = len(df)

        # 筛选条件应用
        condition = (
            (df['入职月份'] == month) &
            (~((df['三级组织'] == '财务中心') & (df['四级组织'] == '证照支持部'))) &
            (df['员工二级类别'] == '正式员工')
        )
        filtered_df = df[condition].copy()
        
        # 记录被排除人员
        excluded_df = df[~condition]
        st.session_state.excluded = excluded_df[['姓名', '三级组织', '四级组织']].to_dict('records')

        # 计算周年数
        filtered_df['周年数'] = year - filtered_df['入职日期'].dt.year
        filtered_df = filtered_df[filtered_df['周年数'] >= 1]

        # 添加备注
        filtered_df['备注'] = filtered_df['员工一级类别'].apply(
            lambda x: '注意外包人员' if x == '外包' else '')

        # 生成人员信息
        def format_info(row):
            base = f"{row['三级组织']}-{row['姓名']}"
            return f"{base}（{row['花名']}）" if pd.notna(row['花名']) else base
            
        filtered_df['人员信息'] = filtered_df.apply(format_info, axis=1)

        # 生成统计结果
        result = (
            filtered_df.groupby('周年数', as_index=False)
            .agg(人员信息=('人员信息', lambda x: '、'.join(x)))
            .sort_values('周年数', ascending=False)
        )
        result['周年标签'] = result['周年数'].astype(str) + '周年'
        
        return filtered_df, result[['周年数', '周年标签', '人员信息']]

    except Exception as e:
        st.error(f"数据处理出错: {str(e)}")
        raise

# ========================
# 分析执行按钮
# ========================
if st.button("开始分析") and uploaded_file is not None:
    try:
        # 读取数据
        df = pd.read_excel(uploaded_file)
        
        # 执行处理
        filtered_df, summary_df = process_data(df, selected_year, selected_month)
        
        # 保存结果到内存
        output1 = BytesIO()
        filtered_df.to_excel(output1, index=False)
        output1.seek(0)
        
        output2 = BytesIO()
        summary_df.to_excel(output2, index=False) 
        output2.seek(0)
        
        # 更新session状态
        st.session_state.result1 = output1
        st.session_state.result2 = output2
        st.session_state.processed = True
        
        # 收集统计信息
        st.session_state.messages.append(f"✅ 原始数据记录数: {len(df)}")
        st.session_state.messages.append(f"✅ 符合条件人员数: {len(filtered_df)}")
        st.session_state.messages.append(f"✅ 排除人员数: {len(st.session_state.excluded)}")
        if any(filtered_df['备注'] != ''):
            st.session_state.messages.append("⚠️ 存在外包人员，请人工复核")

    except Exception as e:
        st.error(f"分析过程出错: {str(e)}")

# ========================
# 结果展示区域
# ========================
if st.session_state.processed:
    st.success("分析完成！请下载结果文件")
    
    # 显示统计信息
    st.subheader("处理结果摘要")
    for msg in st.session_state.messages:
        st.write(msg)
    
    # 显示排除人员
    if st.session_state.excluded:
        st.subheader("被排除人员清单")
        st.dataframe(pd.DataFrame(st.session_state.excluded))

# ========================
# 文件下载区域
# ========================
if st.session_state.result1:
    st.download_button(
        label="下载符合条件人员列表",
        data=st.session_state.result1,
        file_name="符合条件人员列表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.session_state.result2:
    st.download_button(
        label="下载入职周年统计表",
        data=st.session_state.result2,
        file_name="入职周年统计表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ========================
# 错误提示处理
# ========================
if uploaded_file is None and st.session_state.processed:
    st.warning("请先上传数据文件")
elif not st.session_state.processed:
    st.info("请先上传文件并点击开始分析按钮")
