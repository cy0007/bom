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


def test_generate_skus():
    """测试根据款式编码、开发颜色和尺码生成SKU列表的功能"""
    
    # 定义测试用的源文件路径
    source_file = 'tests/fixtures/新品研发明细表-最终版.xlsx'
    
    # 实例化 BomGenerator 类
    generator = BomGenerator(source_file)
    
    # 定义输入数据
    style_code = 'H5A413492'
    dev_colors = '灰色/黑色/杏色'
    # 预期的尺码列表
    sizes = ['S', 'M', 'L', 'XL']
    
    # 调用一个尚不存在的方法
    sku_list = generator.generate_skus(style_code, dev_colors, sizes)
    
    # 预期结果
    expected_skus = [
        {
            'color': '灰色',
            'skus': {
                'S': 'H5A41349215S',
                'M': 'H5A41349215M',
                'L': 'H5A41349215L',
                'XL': 'H5A41349215XL'
            }
        },
        {
            'color': '黑色',
            'skus': {
                'S': 'H5A41349210S',
                'M': 'H5A41349210M',
                'L': 'H5A41349210L',
                'XL': 'H5A41349210XL'
            }
        },
        {
            'color': '杏色',
            'skus': {
                'S': 'H5A41349270S',
                'M': 'H5A41349270M',
                'L': 'H5A41349270L',
                'XL': 'H5A41349270XL'
            }
        }
    ]
    
    # 设置断言来验证返回结果的正确性
    assert sku_list == expected_skus
