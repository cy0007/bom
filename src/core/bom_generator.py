# 核心逻辑，处理并生成文件

from typing import Dict, Any, List
import pandas as pd
import json
import openpyxl
import os


class BomGenerator:
    """BOM生成器类，用于处理Excel文件并生成BOM表
    
    该类主要用于读取包含产品研发明细信息的Excel文件，
    并提供根据款式编码查找产品信息、生成SKU以及生成完整BOM文件的功能。
    
    支持的操作：
    - 读取Excel文件中的产品明细数据
    - 根据款式编码查找对应的产品信息
    - 返回包含波段、品类、开发颜色等关键信息的结构化数据
    - 根据款式编码、开发颜色和尺码生成对应的SKU列表
    - 基于BOM模板生成完整的Excel BOM文件，支持动态颜色数量
    
    Example:
        >>> generator = BomGenerator('product_details.xlsx')
        >>> info = generator.find_style_info('H5A123416')
        >>> print(info['品类'])  # 输出: 长袖T恤
        >>> skus = generator.generate_skus('H5A413492', '灰色/黑色', ['S', 'M'])
        >>> print(skus[0]['skus']['S'])  # 输出: H5A41349215S
    """
    
    # 类级别常量定义，避免魔术字符串
    SHEET_NAME = '明细表'
    STYLE_CODE_COL = '款式编码'
    WAVE_COL = '波段'
    CATEGORY_COL = '品类'
    DEV_COLOR_COL = '开发颜色'
    COLOR_CODES_PATH = 'src/resources/color_codes.json'
    TEMPLATE_PATH = 'src/resources/bom_template.xlsx'
    
    # BOM模板单元格位置配置
    CELL_CONFIG = {
        'product_name': 'C4',        # 品名单元格位置
        'color_start_row': 9,        # 第一个下单颜色的行号
        'sku_start_row': 10,         # 第一个规格码的行号
        'rows_per_color': 2,         # 每个颜色块占用的行数
        'sku_columns': {             # SKU对应的列位置
            'S': 'C',
            'M': 'D', 
            'L': 'E',
            'XL': 'F'
        },
        'color_column': 'B',         # 颜色名称所在列
        'insert_before_row': 15      # 需要插入新行时的基准行（工艺要求行）
    }
    
    def __init__(self, source_path: str) -> None:
        """初始化BomGenerator实例
        
        读取指定的Excel文件，解析产品明细数据并存储在内存中供后续查询使用。
        Excel文件需要包含"明细表"工作表，且具有特定的表头结构。
        
        Args:
            source_path (str): 源Excel文件的完整路径
            
        Raises:
            FileNotFoundError: 当指定的Excel文件不存在时
            ValueError: 当Excel文件格式不正确或缺少必要工作表时
            
        Note:
            Excel文件应包含复杂的多行表头结构，该方法会自动处理表头解析。
        """
        try:
            # 读取指定Excel文件中的"明细表"Sheet，手动处理复杂表头
            temp_df = pd.read_excel(source_path, sheet_name=self.SHEET_NAME, header=1)
            
            # 使用第0行（现在是DataFrame的第一行）作为列名
            new_columns = temp_df.iloc[0].values
            self.df = temp_df.iloc[1:].copy()
            self.df.columns = new_columns
            
            # 验证必要的列是否存在
            required_columns = [self.STYLE_CODE_COL, self.WAVE_COL, 
                              self.CATEGORY_COL, self.DEV_COLOR_COL]
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                raise ValueError(f"Excel文件缺少必要的列: {missing_columns}")
            
            # 加载颜色代码映射表
            try:
                with open(self.COLOR_CODES_PATH, 'r', encoding='utf-8') as f:
                    self.color_codes = json.load(f)
            except FileNotFoundError:
                raise FileNotFoundError(f"错误：颜色代码文件未找到，路径：{self.COLOR_CODES_PATH}")
            except json.JSONDecodeError as e:
                raise ValueError(f"错误：颜色代码文件格式不正确：{str(e)}")
                
        except FileNotFoundError as e:
            if "颜色代码文件" in str(e):
                raise e
            raise FileNotFoundError(f"错误：源文件未找到，路径：{source_path}")
        except Exception as e:
            if "No sheet named" in str(e):
                raise ValueError(f"错误：Excel文件中未找到工作表 '{self.SHEET_NAME}'")
            raise ValueError(f"读取Excel文件时发生错误: {str(e)}")
    
    def find_style_info(self, style_code: str) -> Dict[str, Any]:
        """根据款式编码查找对应的产品样式信息
        
        在已加载的产品数据中搜索指定的款式编码，返回该产品的关键信息。
        返回的信息包括波段、品类和开发颜色等核心产品属性。
        
        Args:
            style_code (str): 要查找的款式编码，如 'H5A123416'
            
        Returns:
            Dict[str, Any]: 包含产品信息的字典，格式如下：
                {
                    '波段': str,        # 产品所属波段，如 '秋四波'
                    '品类': str,        # 产品品类，如 '长袖T恤'
                    '开发颜色': str      # 开发颜色信息，如 '黑色/红色'
                }
                
        Raises:
            ValueError: 当指定的款式编码在数据中不存在时
            
        Example:
            >>> info = generator.find_style_info('H5A123416')
            >>> print(info)
            {'波段': '秋四波', '品类': '长袖T恤', '开发颜色': '黑色/红色'}
        """
        # 查找款式编码匹配的行
        matching_rows = self.df[self.df[self.STYLE_CODE_COL] == style_code]
        
        # 检查是否找到匹配的行
        if matching_rows.empty:
            raise ValueError(f"错误：未在源文件中找到款式编码 '{style_code}'。")
        
        # 获取第一个匹配的行
        row = matching_rows.iloc[0]
        
        # 构建并返回字典
        return {
            self.WAVE_COL: row[self.WAVE_COL],
            self.CATEGORY_COL: row[self.CATEGORY_COL],
            self.DEV_COLOR_COL: row[self.DEV_COLOR_COL]
        }
    
    def generate_skus(self, style_code: str, dev_colors_str: str, sizes: List[str]) -> List[Dict[str, Any]]:
        """根据款式编码、开发颜色和尺码生成SKU列表
        
        根据输入的款式编码、开发颜色字符串和尺码列表，生成对应的SKU信息。
        该方法会解析颜色字符串，查找每个颜色对应的数字代码，然后为每个颜色-尺码
        组合生成唯一的SKU编码。
        
        SKU生成规则：{款式编码}{颜色代码}{尺码}
        例如：H5A41349215S 表示款式H5A413492的灰色(15)S码产品
        
        Args:
            style_code (str): 产品款式编码，如 'H5A413492'
            dev_colors_str (str): 开发颜色字符串，多个颜色用'/'分隔，
                                如 '灰色/黑色/杏色'
            sizes (List[str]): 产品尺码列表，如 ['S', 'M', 'L', 'XL']
            
        Returns:
            List[Dict[str, Any]]: SKU信息列表，每个元素包含以下结构：
                [
                    {
                        'color': str,           # 颜色名称，如 '灰色'
                        'skus': Dict[str, str]  # 尺码到SKU的映射，如 {'S': 'H5A41349215S'}
                    },
                    ...
                ]
                
        Raises:
            ValueError: 当颜色名称在颜色代码字典中找不到时
            
        Example:
            >>> generator = BomGenerator('data.xlsx')
            >>> skus = generator.generate_skus('H5A413492', '灰色/黑色', ['S', 'M'])
            >>> print(skus)
            [
                {
                    'color': '灰色',
                    'skus': {'S': 'H5A41349215S', 'M': 'H5A41349215M'}
                },
                {
                    'color': '黑色', 
                    'skus': {'S': 'H5A41349210S', 'M': 'H5A41349210M'}
                }
            ]
        """
        # 初始化结果列表
        result_list = []
        
        # 根据 '/' 分割颜色字符串
        color_names = [color.strip() for color in dev_colors_str.split('/')]
        
        # 遍历每种颜色
        for color_name in color_names:
            try:
                # 从颜色代码字典中查找对应的代码
                color_code = self.color_codes[color_name]
            except KeyError:
                raise ValueError(f"错误：在颜色代码字典中未找到颜色 '{color_name}'。")
            
            # 创建该颜色的SKU字典
            skus_dict = {}
            
            # 遍历每个尺码
            for size in sizes:
                # 生成SKU字符串
                sku = self._create_sku(style_code, color_code, size)
                # 存入字典
                skus_dict[size] = sku
            
            # 将该颜色的完整信息添加到结果列表
            result_list.append({
                'color': color_name,
                'skus': skus_dict
            })
        
        return result_list
    
    def _create_sku(self, style_code: str, color_code: str, size: str) -> str:
        """创建单个SKU编码
        
        根据款式编码、颜色代码和尺码生成SKU字符串。
        这是一个私有辅助方法，用于保持SKU生成逻辑的一致性。
        
        Args:
            style_code (str): 款式编码
            color_code (str): 颜色数字代码
            size (str): 尺码
            
        Returns:
            str: 生成的SKU编码
            
        Example:
            >>> sku = self._create_sku('H5A413492', '15', 'S')
            >>> print(sku)  # 输出: H5A41349215S
        """
        return f"{style_code}{color_code}{size}"
    
    def generate_bom_file(self, style_code: str, output_dir: str) -> None:
        """生成完整的BOM Excel文件
        
        基于预定义的BOM模板，为指定的款式编码生成包含所有颜色和SKU信息的
        完整Excel BOM文件。该方法支持动态数量的颜色，当颜色超过3种时会
        自动插入新行来容纳额外的颜色信息。
        
        处理流程：
        1. 确保输出目录存在
        2. 加载BOM模板文件
        3. 查找并填充产品基本信息（品名等）
        4. 生成所有颜色的SKU信息
        5. 如果颜色数量超过3种，动态插入新行
        6. 填充所有颜色和SKU信息到对应位置
        7. 保存为新的Excel文件
        
        Args:
            style_code (str): 产品款式编码，如 'H5A123416'
            output_dir (str): 输出目录的完整路径。如果目录不存在会自动创建
            
        Raises:
            ValueError: 当款式编码在源数据中不存在时
            FileNotFoundError: 当BOM模板文件不存在时
            PermissionError: 当无法创建输出目录或写入文件时
            
        Example:
            >>> generator = BomGenerator('source.xlsx')
            >>> generator.generate_bom_file('H5A123416', './output')
            # 将在 ./output/ 目录下生成 H5A123416.xlsx 文件
            
        Note:
            - 生成的文件名格式为: {款式编码}.xlsx
            - 支持任意数量的颜色，超过3种会自动扩展表格行数
            - 品名格式为: HECO{波段}{品类}{款式编码}
            - SKU生成规则: {款式编码}{颜色代码}{尺码}
        """
        # 1. 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 2. 加载模板文件
        try:
            workbook = openpyxl.load_workbook(self.TEMPLATE_PATH)
            sheet = workbook.active
        except FileNotFoundError:
            raise FileNotFoundError(f"错误：BOM模板文件未找到，路径：{self.TEMPLATE_PATH}")
        
        # 3. 获取产品基本信息
        style_info = self.find_style_info(style_code)
        
        # 4. 生成品名（HECO + 波段 + 品类 + 款式编码）
        product_name = f"HECO{style_info[self.WAVE_COL]}{style_info[self.CATEGORY_COL]}{style_code}"
        
        # 5. 填充产品名称到工作表（处理合并单元格）
        self._write_to_cell(sheet, self.CELL_CONFIG['product_name'], product_name)
        
        # 6. 生成SKU列表
        dev_colors = style_info[self.DEV_COLOR_COL]
        sizes = ['S', 'M', 'L', 'XL']
        sku_list = self.generate_skus(style_code, dev_colors, sizes)
        
        # 7. 处理动态行插入（如果颜色超过3种）
        if len(sku_list) > 3:
            self._insert_additional_rows(sheet, len(sku_list))
        
        # 8. 填充所有颜色和SKU信息
        self._fill_color_and_sku_data(sheet, sku_list)
        
        # 9. 保存文件
        output_file_path = os.path.join(output_dir, f"{style_code}.xlsx")
        try:
            workbook.save(output_file_path)
        except PermissionError:
            raise PermissionError(f"错误：无法保存文件到 {output_file_path}，请检查目录权限。")
    
    def _write_to_cell(self, sheet, cell_address: str, value: str) -> None:
        """向Excel单元格写入值，处理合并单元格情况
        
        Args:
            sheet: openpyxl工作表对象
            cell_address (str): 单元格地址，如 'C4'
            value (str): 要写入的值
        """
        target_cell = sheet[cell_address]
        if target_cell.__class__.__name__ == 'MergedCell':
            # 找到包含目标单元格的合并区域的主单元格
            for range_ in sheet.merged_cells.ranges:
                if cell_address in range_:
                    # 写入到合并区域的左上角单元格
                    main_cell = sheet.cell(row=range_.min_row, column=range_.min_col)
                    main_cell.value = value
                    break
            else:
                # 如果找不到合并区域，直接写入目标单元格
                target_cell.value = value
        else:
            target_cell.value = value
    
    def _insert_additional_rows(self, sheet, color_count: int) -> None:
        """当颜色数量超过3种时，动态插入新行
        
        Args:
            sheet: openpyxl工作表对象
            color_count (int): 颜色总数
        """
        if color_count <= 3:
            return
        
        # 计算需要插入的行数：每超过一种颜色需要插入2行
        additional_colors = color_count - 3
        rows_to_insert = additional_colors * self.CELL_CONFIG['rows_per_color']
        
        # 在指定位置插入新行
        insert_row = self.CELL_CONFIG['insert_before_row']
        sheet.insert_rows(insert_row, rows_to_insert)
    
    def _fill_color_and_sku_data(self, sheet, sku_list: List[Dict[str, Any]]) -> None:
        """填充所有颜色和SKU信息到工作表
        
        Args:
            sheet: openpyxl工作表对象
            sku_list (List[Dict[str, Any]]): SKU信息列表
        """
        config = self.CELL_CONFIG
        color_column = config['color_column']
        sku_columns = config['sku_columns']
        
        for i, color_data in enumerate(sku_list):
            # 计算当前颜色块的行位置
            color_row = config['color_start_row'] + (i * config['rows_per_color'])
            sku_row = config['sku_start_row'] + (i * config['rows_per_color'])
            
            # 填充颜色名称 - 使用_write_to_cell处理合并单元格
            color_cell_address = f"{color_column}{color_row}"
            self._write_to_cell(sheet, color_cell_address, color_data['color'])
            
            # 填充SKU信息 - 使用_write_to_cell处理合并单元格
            for size, sku in color_data['skus'].items():
                if size in sku_columns:
                    column = sku_columns[size]
                    sku_cell_address = f"{column}{sku_row}"
                    self._write_to_cell(sheet, sku_cell_address, sku)
