"""
Excel数据统计工具
版本: 1.1.0
更新日期: 2026-03-09
功能: 支持14日和30日统计周期切换的Excel批量数据分析工具
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
from typing import List, Dict, Tuple
导入警告

# 版本信息
__version__ = "1.1.0"
__update_date__ = "2026-03-09"

# 性能优化设置
警告。filterwarnings('忽略')
pd.options.mode.chained_assignment = None

# 页面配置
st.set_page_config(
    page_title="Excel数据统计工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS - 马卡龙淡紫色+白色+蓝色点缀
st.markdown("""
<style>
    /* 主背景 */
    .stApp {
        background: linear-gradient(135deg, #f5f3ff 0%, #ffffff 100%);
    }
    
    /* 侧边栏 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #e9d5ff 0%, #f3e8ff 100%);
    }
    
    /* 标题样式 */
    h1 {
        color: #7c3aed;
        font-weight: 700;
        text-shadow: 2px 2px 4px rgba(124, 58, 237, 0.1);
    }
    
    h2, h3 {
        color: #8b5cf6;
    }
    
    /* 文件上传区域 */
    [data-testid="stFileUploader"] {
        background: white;
        border: 2px dashed #c4b5fd;
        border-radius: 15px;
        padding: 20px;
    }
    
    /* 按钮样式 */
    .stButton > button {
        background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 10px 30px;
        font-weight: 600;
        box-shadow: 0 4px 6px rgba(139, 92, 246, 0.3);
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(139, 92, 246, 0.4);
    }
    
    /* 数据表格 */
    [data-testid="stDataFrame"] {
        background: white;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(124, 58, 237, 0.1);
    }
    
    /* 成功消息 */
    .stSuccess {
        background: #ddd6fe;
        color: #5b21b6;
        border-left: 4px solid #8b5cf6;
    }
    
    /* 信息卡片 */
    .info-card {
        background: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(124, 58, 237, 0.15);
        border-left: 5px solid #8b5cf6;
        margin: 10px 0;
    }
    
    /* 蓝色点缀 */
    .accent-blue {
        color: #3b82f6;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


def extract_batch_number(filename: str) -> int:
    """
    从文件名中提取批次号
    例如: "第1001批xxx.xlsx" -> 1001
    """
    match = re.search(r'第(\d+)批', filename)
    if match:
        return int(match.group(1))
    return 0


def process_excel_file(file, days: int = 14) -> Dict:
    """
    处理单个Excel文件，统计所需数据
    参数:
        file: Excel文件对象
        days: 统计天数，14或30
    返回: {"batch": 批次号, "b": 登录数, "c": 消费数, "d": 总金额, "e": 套餐数}
    """
    try:
        # 提取批次号
        batch_num = extract_batch_number(file.name)
        
        # 读取Excel
        try:
            # 重置文件指针到开始位置
            file.seek(0)
            df = pd.read_excel(file)
        except Exception as read_error:
            return {
                'batch': batch_num,
                'filename': file.name,
                'status': 'error',
                'error_msg': f'读取文件失败: {str(read_error)}',
                'error_type': 'file_read'
            }
        
        # 验证数据不为空
        if df.empty:
            return {
                'batch': batch_num,
                'filename': file.name,
                'status': 'error',
                'error_msg': '文件为空，没有数据',
                'error_type': 'empty_file'
            }
        
        # 检查必需的列是否存在
        required_columns = [f'{days}日内是否登录', f'{days}日内是否消费']
        available_columns = list(df.columns)
        missing_columns = [col for col in required_columns if col not in available_columns]
        
        if missing_columns:
            return {
                'batch': batch_num,
                'filename': file.name,
                'status': 'error',
                'error_msg': f'❌ 缺少必需的列: {", ".join(missing_columns)}',
                'error_type': 'missing_columns',
                'available_columns': available_columns,
                'missing_columns': missing_columns
            }
        
        # b列: "X日内是否登录"中值为1的数量
        b_count = 0
        login_column = f'{days}日内是否登录'
        if login_column in df.columns:
            try:
                # 转换为数值类型，处理可能的字符串
                login_col = pd.to_numeric(df[login_column], errors='coerce')
                b_count = int((login_col == 1).sum())
            except Exception as e:
                pass  # 如果转换失败，保持为0
        
        # c列: "X日内是否消费"中值为1的数量
        c_count = 0
        consume_column = f'{days}日内是否消费'
        if consume_column in df.columns:
            try:
                consume_col = pd.to_numeric(df[consume_column], errors='coerce')
                c_count = int((consume_col == 1).sum())
            except Exception as e:
                pass
        
        # d列: 三个消费金额字段求和
        d_sum = 0
        amount_columns = [
            f'{days}日内会员消费金额',
            f'{days}日内声音包消费金额',
            f'{days}日内通用会员消费金额'
        ]
        for col in amount_columns:
            if col in df.columns:
                try:
                    # 转换为数值类型，处理可能的非数值
                    amount = pd.to_numeric(df[col], errors='coerce').fillna(0).sum()
                    d_sum += float(amount)
                except Exception as e:
                    pass
        
        # e列: "X日内购买记录"中特定套餐出现次数
        e_count = 0
        purchase_column = f'{days}日内购买记录'
        if purchase_column in df.columns:
            try:
                purchase_records = df[purchase_column].fillna('').astype(str)
                e_count = int(purchase_records.str.contains('两年年套餐499元|UNKNOWN499元', regex=True).sum())
            except Exception as e:
                pass
        
        result = {
            'batch': batch_num,
            'b': b_count,
            'c': c_count,
            'd': round(d_sum, 2),
            'e': e_count,
            'filename': file.name,
            'status': 'success',
            'rows': len(df)
        }
        
        # 清理DataFrame释放内存
        del df
        return result
    
    except Exception as e:
        import traceback
        return {
            'batch': extract_batch_number(file.name),
            'filename': file.name,
            'status': 'error',
            'error_msg': f'{type(e).__name__}: {str(e)}',
            'error_type': 'unknown',
            'traceback': traceback.format_exc()
        }


def create_summary_dataframe(results: List[Dict]) -> pd.DataFrame:
    """
    创建汇总数据表格
    """
    # 过滤成功的结果
    success_results = [r for r in results if r['status'] == 'success']
    
    if not success_results:
        return pd.DataFrame()
    
    # 创建DataFrame
    df = pd.DataFrame(success_results)
    
    # 选择需要的列并排序
    df = df[['batch', 'b', 'c', 'd', 'e']].sort_values('batch')
    
    # 重命名列
    df.columns = ['批次号', 'b列(登录数)', 'c列(消费数)', 'd列(总金额)', 'e列(套餐数)']
    
    # 重置索引
    df = df.reset_index(drop=True)
    
    return df


def create_visualizations(df: pd.DataFrame):
    """
    创建数据可视化图表
    """
    if df.empty:
        return
    
    # 准备数据
    batch_col = df.columns[0]  # 批次号
    
    # 创建四个子图
    col1, col2 = st.columns(2)
    
    with col1:
        # 登录数和消费数对比
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            x=df[batch_col],
            y=df['b列(登录数)'],
            name='登录数',
            marker_color='#8b5cf6'
        ))
        fig1.add_trace(go.Bar(
            x=df[batch_col],
            y=df['c列(消费数)'],
            name='消费数',
            marker_color='#3b82f6'
        ))
        fig1.update_layout(
            title='登录数 vs 消费数',
            xaxis_title='批次号',
            yaxis_title='数量',
            barmode='group',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # 总金额趋势
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=df[batch_col],
            y=df['d列(总金额)'],
            mode='lines+markers',
            name='总金额',
            line=dict(color='#8b5cf6', width=3),
            marker=dict(size=10, color='#6366f1')
        ))
        fig2.update_layout(
            title='总消费金额趋势',
            xaxis_title='批次号',
            yaxis_title='金额(元)',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig2, use_container_width=True)
    
    # 套餐购买数
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        x=df[batch_col],
        y=df['e列(套餐数)'],
        marker_color='#c4b5fd',
        marker_line_color='#8b5cf6',
        marker_line_width=2
    ))
    fig3.update_layout(
        title='499元套餐购买数量',
        xaxis_title='批次号',
        yaxis_title='购买次数',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)'
    )
    st.plotly_chart(fig3, use_container_width=True)


def validate_file(file, days: int = 14) -> Dict:
    """
    预验证文件是否可以正常处理
    参数:
        file: Excel文件对象
        days: 统计天数，14或30
    返回验证结果
    """
    try:
        batch_num = extract_batch_number(file.name)
        
        # 重置文件指针到开始位置
        file.seek(0)
        
        # 尝试读取文件
        df = pd.read_excel(file)
        
        if df.empty:
            return {
                'filename': file.name,
                'batch': batch_num,
                'valid': False,
                'issue': '文件为空',
                'rows': 0,
                'columns': 0
            }
        
        # 检查必需列
        required_columns = [f'{days}日内是否登录', f'{days}日内是否消费']
        available_columns = list(df.columns)
        missing = [col for col in required_columns if col not in available_columns]
        
        if missing:
            result = {
                'filename': file.name,
                'batch': batch_num,
                'valid': False,
                'issue': f'缺少列: {", ".join(missing)}',
                'rows': len(df),
                'columns': len(df.columns),
                'missing_columns': missing,
                'available_columns': available_columns
            }
            del df
            return result
        
        result = {
            'filename': file.name,
            'batch': batch_num,
            'valid': True,
            'rows': len(df),
            'columns': len(df.columns)
        }
        del df
        return result
    
    except Exception as e:
        return {
            'filename': file.name,
            'batch': extract_batch_number(file.name),
            'valid': False,
            'issue': f'读取失败: {str(e)}',
            'rows': 0,
            'columns': 0
        }


def export_to_excel(df: pd.DataFrame) -> BytesIO:
    """
    导出数据为Excel文件
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='统计结果')
    output.seek(0)
    return output


# 主应用
def main():
    # 标题
    st.markdown('<h1 style="text-align: center;">📊 Excel数据统计工具</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #8b5cf6; font-size: 18px;">订单销售数据批量分析系统</p>', unsafe_allow_html=True)
    
    st.markdown('---')
    
    # 侧边栏说明
    with st.sidebar:
        st.markdown('### 📖 使用说明')
        
        # 时间范围选择
        st.markdown('#### 📅 统计时间范围')
        time_range = st.radio(
            "选择统计周期",
            options=["14日", "30日"],
            horizontal=True,
            help="切换统计14日或30日的数据"
        )
        days = 14 if time_range == "14日" else 30
        
        st.markdown(f"""
        <div class="info-card">
        <h4>统计规则（{time_range}）</h4>
        <ul>
            <li><span class="accent-blue">b列</span>: {days}日内登录用户数</li>
            <li><span class="accent-blue">c列</span>: {days}日内消费用户数</li>
            <li><span class="accent-blue">d列</span>: 三类消费金额总和</li>
            <li><span class="accent-blue">e列</span>: 499元套餐购买次数</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="info-card">
        <h4>操作步骤</h4>
        <ol>
            <li>上传10个Excel文件</li>
            <li>点击"开始统计"按钮</li>
            <li>查看统计结果和图表</li>
            <li>下载Excel报告</li>
        </ol>
        </div>
        """, unsafe_allow_html=True)
    
    # 文件上传区
    st.markdown('### 📁 上传Excel文件')
    uploaded_files = st.file_uploader(
        "请选择Excel文件（支持批量上传）",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="文件命名格式：第xxx批..."
    )
    
    if uploaded_files:
        st.success(f'✅ 已上传 {len(uploaded_files)} 个文件')
        
        # 显示文件列表
        with st.expander('📋 查看上传的文件'):
            for idx, file in enumerate(uploaded_files, 1):
                batch = extract_batch_number(file.name)
                st.write(f"{idx}. {file.name} (批次: {batch})")
        
        st.markdown('---')
        
        # 操作按钮区域
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            if st.button('🗑️ 清除数据', help='清除当前数据，重新开始分析', type='secondary'):
                # 清除所有session state
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        with col2:
            if st.button('🔄 刷新', help='刷新当前页面', type='secondary'):
                # 清除所有session state（包括上传的文件）
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        
        st.markdown('---')
        
        # 统计按钮
        if st.button('🚀 开始统计', type='primary'):
            with st.spinner('正在处理数据...'):
                # 处理所有文件
                results = []
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, file in enumerate(uploaded_files):
                    # 显示当前处理的文件
                    status_text.info(f'📄 正在处理: {file.name} ({idx + 1}/{len(uploaded_files)})')
                    
                    # 确保文件指针在开始位置
                    try:
                        file.seek(0)
                    except:
                        pass
                    
                    # 处理文件，传入天数参数
                    result = process_excel_file(file, days=days)
                    if result['status'] == 'error':
                        status_text.warning(f'⚠️ 重试: {file.name}')
                        try:
                            file.seek(0)
                            result = process_excel_file(file, days=days)
                        except:
                            pass
                    
                    results.append(result)
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                
                progress_bar.empty()
                status_text.empty()
                
                # 检查错误
                errors = [r for r in results if r['status'] == 'error']
                if errors:
                    st.warning(f'⚠️ {len(errors)} 个文件处理失败')
                    with st.expander('🔍 查看错误详情', expanded=True):
                        for err in errors:
                            st.error(f"**文件**: {err['filename']}")
                            st.write(f"**批次号**: {err.get('batch', '未知')}")
                            st.write(f"**错误类型**: {err.get('error_type', 'unknown')}")
                            st.write(f"**错误信息**: {err['error_msg']}")
                            
                            # 如果是缺失列的错误，显示详细信息
                            if err.get('error_type') == 'missing_columns':
                                st.info('💡 **诊断建议**')
                                if 'missing_columns' in err:
                                    st.write(f"❌ **缺失的列**: {', '.join(err['missing_columns'])}")
                                if 'available_columns' in err:
                                    st.write(f"✅ **文件中实际的列** ({len(err['available_columns'])}个):")
                                    cols_display = ', '.join(err['available_columns'][:10])
                                    if len(err['available_columns']) > 10:
                                        cols_display += f" ... (还有{len(err['available_columns']) - 10}个)"
                                    st.code(cols_display)
                                st.warning(f'🔧 **解决方法**: 请检查Excel文件，确保包含"{days}日内是否登录"和"{days}日内是否消费"这两列')
                            
                            # 如果是文件读取错误，提供建议
                            elif err.get('error_type') == 'file_read':
                                st.info('💡 **可能的原因**')
                                st.write('- 文件可能已损坏')
                                st.write('- 文件格式不正确（不是标准的Excel格式）')
                                st.write('- 文件可能正在被其他程序使用')
                                st.warning('🔧 **解决方法**: 尝试重新保存文件，或使用Excel打开后另存为新文件')
                            
                            # 如果有traceback，显示在代码块中
                            如果错误中包含回溯：
                                with st.expander('📋 详细错误堆栈（技术信息）'):
                                    st.code(err['traceback'], language='python')
                            
                            st.markdown('---')
                
                # 创建汇总表格
                summary_df = create_summary_dataframe(results)
                
                if not summary_df.empty:
                    st.success('✅ 统计完成！')
                    
                    # 显示汇总表格
                    st.markdown('### 📊 统计结果汇总')
                    st.dataframe(
                        summary_df,
                        use_container_width=真,
                        height=400
                    )
                    
                    # 显示总计
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric('总登录数', f"{summary_df['b列(登录数)'].sum():,}")
                    with col2:
                        st.metric('总消费数', f"{summary_df['c列(消费数)'].sum():,}")
                    与col3:
                        st.metric'总金额', f"¥{summary_df['d列(总金额)'].sum():,.2f}")
                    与col4:
                        st.metric('总套餐数', f"{summary_df['e列(套餐数)'].sum():,}")
                    
                    st.markdown('---')
                    
                    # 数据可视化
                    st.markdown('### 📈 数据可视化')
                    create_visualizations(summary_df)
                    
                    st.markdown('---')
                    
                    # 导出按钮
                    st.markdown('### 💾 导出报告')
                    excel_data = export_to_excel(summary_df)
                    st.download_button(
                        label='📥 下载Excel报告',
                        data=excel_data,
                        file_name='统计结果.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                :
                    st.error('❌ 没有成功处理的文件')
    
    else:
        st.info('👆 请上传Excel文件开始统计')
    
    # 页脚
    st.markdown('---')
    st.markdown(
        f'<p style="text-align: center; color: #a78bfa; font-size: 14px;">'
        f'Excel数据统计工具 v{__version__} | 更新于 {__update_date__} | Powered by Streamlit'
        f'</p>',
        unsafe_allow_html=True
    )


if __name__ == '__main__':
    main()
