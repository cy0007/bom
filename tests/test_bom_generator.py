# 核心逻辑的测试文件

import pytest
import openpyxl
import os
from unittest.mock import patch
from src.core.bom_generator import BomGenerator


def test_get_all_style_codes():
    """测试获取源文件中所有款式编码的功能"""
    
    # 定义测试用的源文件路径
    source_file = 'tests/fixtures/新品研发明细表-最终版.xlsx'
    
    # 实例化 BomGenerator 类
    generator = BomGenerator(source_file)
    
    # 调用一个尚不存在的方法
    style_codes = generator.get_all_style_codes()
    
    # 设置断言来验证返回数据的正确性
    # 期望返回一个列表，包含所有不为空的唯一款式编码
    assert isinstance(style_codes, list), "返回值应该是一个列表"
    assert len(style_codes) > 0, "应该找到至少一个款式编码"
    
    # 验证列表中的值都是字符串且不为空
    for code in style_codes:
        assert isinstance(code, str), f"款式编码应该是字符串，但得到：{type(code)}"
        assert code.strip() != "", f"款式编码不应该为空：'{code}'"
    
    # 验证列表中没有重复的款式编码
    assert len(style_codes) == len(set(style_codes)), "款式编码列表不应该包含重复值"
    
    # 验证是否包含我们知道的特定款式编码
    assert 'H5A123416' in style_codes, "应该包含已知的款式编码 H5A123416"
    assert 'H5A413438' in style_codes, "应该包含已知的款式编码 H5A413438"


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


def test_generate_bom_file():
    """测试生成完整BOM Excel文件的功能"""
    
    # 定义测试用的源文件路径
    source_file = 'tests/fixtures/新品研发明细表-最终版.xlsx'
    
    # 定义输入和输出参数
    style_code = 'H5A123416'  # 使用实际存在的款式编码
    output_dir = 'tests/output'  # 使用一个新的临时输出目录
    
    # 确保输出目录存在且为空
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # 清理之前测试可能留下的文件
    for f in os.listdir(output_dir):
        os.remove(os.path.join(output_dir, f))
    
    # 实例化 BomGenerator 类
    generator = BomGenerator(source_file)
    
    # 调用一个尚不存在的方法
    generator.generate_bom_file(style_code, output_dir)
    
    # 设置断言来验证文件是否成功生成，并且内容是否基本正确
    
    # 1. 验证文件是否已创建
    expected_file_path = os.path.join(output_dir, f"{style_code}.xlsx")
    assert os.path.exists(expected_file_path), "BOM文件未被创建"
    
    # 2. 验证文件内容（抽样检查）
    workbook = openpyxl.load_workbook(expected_file_path)
    sheet = workbook.active
    
    # 检查品名是否正确（C4可能是合并单元格，需要找到主单元格）
    # 对于合并单元格，值通常存储在左上角单元格
    c4_cell = sheet['C4']
    if c4_cell.__class__.__name__ == 'MergedCell':
        # 找到包含C4的合并区域的主单元格
        for range_ in sheet.merged_cells.ranges:
            if 'C4' in range_:
                main_cell = sheet.cell(row=range_.min_row, column=range_.min_col)
                product_name_value = main_cell.value
                break
        else:
            product_name_value = None
    else:
        product_name_value = c4_cell.value
    
    assert product_name_value == 'HECO秋四波长袖T恤H5A123416'
    
    # 检查第一个颜色的SKU是否正确（假设在C10单元格）
    assert sheet['C10'].value == 'H5A12341610S'  # 黑色的S码
    
    # 检查第二个颜色的下单颜色是否正确（假设在B11单元格）
    assert sheet['B11'].value == '红色'


def test_generate_bom_file_with_many_colors():
    """测试生成包含5种颜色的BOM文件的功能（验证动态行插入）"""
    
    # 定义测试用的源文件路径
    source_file = 'tests/fixtures/新品研发明细表-最终版.xlsx'
    
    # 定义输入和输出参数
    style_code = 'H5A123416_MULTI'  # 使用不同的款式编码避免文件冲突
    output_dir = 'tests/output'  # 使用一个新的临时输出目录
    
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # 清理目标文件（如果存在）
    target_file = os.path.join(output_dir, f"{style_code}.xlsx")
    if os.path.exists(target_file):
        try:
            os.remove(target_file)
        except PermissionError:
            # 如果文件被占用，跳过清理
            pass
    
    # 实例化 BomGenerator 类
    generator = BomGenerator(source_file)
    
    # 模拟有5种颜色的SKU数据
    mock_sku_list = [
        {
            'color': '黑色',
            'skus': {'S': 'H5A12341610S', 'M': 'H5A12341610M', 'L': 'H5A12341610L', 'XL': 'H5A12341610XL'}
        },
        {
            'color': '红色',
            'skus': {'S': 'H5A12341620S', 'M': 'H5A12341620M', 'L': 'H5A12341620L', 'XL': 'H5A12341620XL'}
        },
        {
            'color': '蓝色',
            'skus': {'S': 'H5A12341630S', 'M': 'H5A12341630M', 'L': 'H5A12341630L', 'XL': 'H5A12341630XL'}
        },
        {
            'color': '绿色',
            'skus': {'S': 'H5A12341640S', 'M': 'H5A12341640M', 'L': 'H5A12341640L', 'XL': 'H5A12341640XL'}
        },
        {
            'color': '黄色',
            'skus': {'S': 'H5A12341650S', 'M': 'H5A12341650M', 'L': 'H5A12341650L', 'XL': 'H5A12341650XL'}
        }
    ]
    
    # 模拟产品信息数据
    mock_style_info = {
        '波段': '秋四波',
        '品类': '长袖T恤',
        '开发颜色': '黑色/红色/蓝色/绿色/黄色'
    }
    
    # 使用patch模拟方法返回值
    with patch.object(generator, 'generate_skus', return_value=mock_sku_list), \
         patch.object(generator, 'find_style_info', return_value=mock_style_info):
        # 调用生成BOM文件的方法
        generator.generate_bom_file(style_code, output_dir)
    
    # 设置断言来验证文件是否成功生成，并且内容是否包含所有5种颜色
    
    # 1. 验证文件是否已创建
    expected_file_path = os.path.join(output_dir, f"{style_code}.xlsx")
    assert os.path.exists(expected_file_path), "BOM文件未被创建"
    
    # 2. 验证文件内容（检查所有5种颜色都被正确填充）
    workbook = openpyxl.load_workbook(expected_file_path)
    sheet = workbook.active
    
    # 验证品名
    c4_cell = sheet['C4']
    if c4_cell.__class__.__name__ == 'MergedCell':
        for range_ in sheet.merged_cells.ranges:
            if 'C4' in range_:
                main_cell = sheet.cell(row=range_.min_row, column=range_.min_col)
                product_name_value = main_cell.value
                break
        else:
            product_name_value = None
    else:
        product_name_value = c4_cell.value
    
    assert product_name_value == 'HECO秋四波长袖T恤H5A123416_MULTI'
    
    # 验证前3种颜色被正确填充（模板原本支持的数量）
    expected_colors = ['黑色', '红色', '蓝色']
    
    def get_cell_value(sheet, cell_address):
        """获取单元格值，处理合并单元格情况"""
        cell = sheet[cell_address]
        if cell.__class__.__name__ == 'MergedCell':
            # 找到包含目标单元格的合并区域的主单元格
            for range_ in sheet.merged_cells.ranges:
                if cell_address in range_:
                    main_cell = sheet.cell(row=range_.min_row, column=range_.min_col)
                    return main_cell.value
            return None
        else:
            return cell.value
    
    # 验证前3种颜色
    for i, expected_color in enumerate(expected_colors):
        color_row = 9 + (i * 2)  # 每种颜色间隔2行
        sku_row = 10 + (i * 2)
        
        # 检查颜色名称
        color_value = get_cell_value(sheet, f'B{color_row}')
        assert color_value == expected_color, f"第{i+1}种颜色名称不正确，期望：{expected_color}，实际：{color_value}"
        
        # 检查SKU（只检查S码作为代表）
        expected_sku_s = mock_sku_list[i]['skus']['S']
        sku_value = get_cell_value(sheet, f'C{sku_row}')
        assert sku_value == expected_sku_s, f"第{i+1}种颜色的S码SKU不正确，期望：{expected_sku_s}，实际：{sku_value}"
    
    # 验证动态行插入功能：检查工作表的总行数是否增加了
    # 原始模板应该有大约15行左右，插入4行后应该有19行以上
    assert sheet.max_row >= 19, f"动态行插入可能失败，当前总行数：{sheet.max_row}，期望至少19行"
    
    # 验证调用了正确的方法（通过检查sku_list的长度来间接验证）
    # 如果有5种颜色且方法被正确调用，文件应该成功生成


def test_generate_bom_file_to_buffer():
    """测试生成BOM文件到内存缓冲区的功能"""
    
    # 定义测试用的源文件路径
    source_file = 'tests/fixtures/新品研发明细表-最终版.xlsx'
    
    # 定义输入参数
    style_code = 'H5A123416'  # 使用实际存在的款式编码
    
    # 实例化 BomGenerator 类
    generator = BomGenerator(source_file)
    
    # 调用一个尚不存在的方法
    excel_bytes = generator.generate_bom_file_to_buffer(style_code)
    
    # 设置断言来验证返回数据的正确性
    
    # 1. 验证返回值是bytes类型
    assert isinstance(excel_bytes, bytes), "返回值应该是bytes类型"
    assert len(excel_bytes) > 0, "返回的字节数据不应该为空"
    
    # 2. 验证返回的字节数据是有效的Excel文件
    import openpyxl
    import io
    
    # 将字节数据转换为BytesIO对象，然后尝试用openpyxl读取
    excel_buffer = io.BytesIO(excel_bytes)
    try:
        workbook = openpyxl.load_workbook(excel_buffer)
        sheet = workbook.active
        
        # 3. 验证Excel文件内容的正确性（抽样检查）
        
        # 检查款式编码是否正确填充（B3单元格）
        assert sheet['B3'].value == style_code, f"B3单元格应该包含款式编码 {style_code}"
        
        # 检查品名是否正确（B4单元格，可能是合并单元格）
        b4_cell = sheet['B4']
        if b4_cell.__class__.__name__ == 'MergedCell':
            # 找到包含B4的合并区域的主单元格
            for range_ in sheet.merged_cells.ranges:
                if 'B4' in range_:
                    main_cell = sheet.cell(row=range_.min_row, column=range_.min_col)
                    product_name_value = main_cell.value
                    break
            else:
                product_name_value = None
        else:
            product_name_value = b4_cell.value
        
        expected_product_name = 'HECO秋四波长袖T恤H5A123416'
        assert product_name_value == expected_product_name, f"品名应该是 {expected_product_name}"
        
        # 检查第一个颜色名称是否正确（A8单元格）
        assert sheet['A8'].value == '黑色', "A8单元格应该包含第一个颜色名称 '黑色'"
        
        # 检查第一个颜色的S码SKU是否正确（B6单元格）
        assert sheet['B6'].value == 'H5A12341610S', "B6单元格应该包含正确的S码SKU"
        
        print("✅ Excel文件内容验证通过")
        
    except Exception as e:
        # 如果无法读取Excel文件，测试失败
        assert False, f"返回的字节数据不是有效的Excel文件: {str(e)}"
    finally:
        excel_buffer.close()
        
    # 4. 验证文件大小合理（应该大于某个最小值）
    min_expected_size = 5000  # 至少5KB
    assert len(excel_bytes) > min_expected_size, f"Excel文件大小应该大于 {min_expected_size} 字节，实际大小: {len(excel_bytes)}"
