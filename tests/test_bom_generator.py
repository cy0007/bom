# 核心逻辑的测试文件

import pytest
from src.core.bom_generator import BomGenerator


def test_find_style_info_by_code():
    """测试根据款式编码查找样式信息的功能"""
    
    # 定义测试用的源文件路径
    source_file = 'tests/fixtures/新品研发明细表-最终版.xlsx'
    
    # 定义要查找的款式编码（使用实际存在的编码）
    style_code = 'H5A123416'
    
    # 实例化 BomGenerator 类
    generator = BomGenerator(source_file)
    
    # 调用一个尚不存在的方法
    style_info = generator.find_style_info(style_code)
    
    # 设置断言来验证返回数据的正确性
    # 我们期望返回一个字典，至少包含以下键值对（使用实际存在的数据）
    assert style_info['波段'] == '秋四波'
    assert style_info['品类'] == '长袖T恤'
    assert style_info['开发颜色'] == '黑色/红色'
