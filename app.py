import streamlit as st
import pandas as pd
from src.core.bom_generator import BomGenerator
import io
import zipfile

# 设置页面标题和布局
st.set_page_config(page_title="BOM表自动生成工具", layout="wide")

st.title("🚀 BOM表自动生成工具 (Web版)")
st.write("---")

# --- UI 交互部分 ---
st.header("1. 上传源文件")
uploaded_file = st.file_uploader(
    "请上传包含'明细表'的《新品研发明细表-最终版.xlsx》文件",
    type=["xlsx"]
)

# --- 逻辑处理部分 ---
if uploaded_file is not None:
    try:
        # Streamlit上传的文件是内存中的字节流，我们需要将其读入BomGenerator
        # BomGenerator的__init__可以直接接收这种字节流对象
        
        st.info("正在读取文件并分析内容...")
        
        # 将上传的文件内容读入BytesIO对象
        file_buffer = io.BytesIO(uploaded_file.getvalue())
        
        # 实例化我们的核心引擎
        generator = BomGenerator(file_buffer)
        
        st.success("文件读取成功！")
        
        # 从BomGenerator实例中获取所有款式编码
        style_codes = generator.get_all_style_codes()
        
        st.write("---")
        st.header("2. 文件内容预览")
        st.write(f"在文件中找到了 **{len(style_codes)}** 个有效的款式编码：")
        
        # 以多列形式展示款式编码，更美观
        num_columns = 5
        columns = st.columns(num_columns)
        for i, code in enumerate(style_codes):
            with columns[i % num_columns]:
                st.info(code)
        
        st.write("---")
        st.header("3. 选择要生成的款式")
        
        # 全选/全不选的逻辑
        select_all = st.checkbox("全选所有款式")
        
        if select_all:
            selected_codes = st.multiselect(
                "或取消选择不需要的款式:",
                options=style_codes,
                default=style_codes  # 默认全部选中
            )
        else:
            selected_codes = st.multiselect(
                "请选择一个或多个款式编码:",
                options=style_codes
            )
        
        st.write("---")
        st.header("4. 批量生成BOM表")
        
        # 显示已选择的款式数量
        if selected_codes:
            st.success(f"已选择 {len(selected_codes)} 个款式编码进行生成")
        else:
            st.warning("请至少选择一个款式编码")
        
        # 生成按钮
        if st.button("🚀 开始生成BOM表", type="primary", disabled=len(selected_codes) == 0):
            if not selected_codes:
                st.warning("⚠️ 请至少选择一个款式编码再进行生成")
            else:
                # 批量生成逻辑
                with st.spinner(f"正在生成 {len(selected_codes)} 个BOM文件，请稍候..."):
                    try:
                        # 创建内存中的ZIP文件
                        zip_buffer = io.BytesIO()
                        
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            # 进度条
                            progress_bar = st.progress(0)
                            
                            for i, code in enumerate(selected_codes):
                                # 生成单个BOM文件到内存
                                excel_bytes = generator.generate_bom_file_to_buffer(code)
                                
                                # 添加到ZIP文件
                                zip_file.writestr(f"{code}.xlsx", excel_bytes)
                                
                                # 更新进度条
                                progress_bar.progress((i + 1) / len(selected_codes))
                        
                        # 将ZIP缓冲区指针移到开头
                        zip_buffer.seek(0)
                        
                        st.success(f"✅ 成功生成 {len(selected_codes)} 个BOM文件！")
                        
                        # 提供下载按钮
                        st.download_button(
                            label="📥 点击下载BOM压缩包 (.zip)",
                            data=zip_buffer.getvalue(),
                            file_name="BOM_files.zip",
                            mime="application/zip",
                            type="primary"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ 生成BOM文件时发生错误：{str(e)}")
                
    except Exception as e:
        st.error(f"处理文件时发生错误：{e}")
