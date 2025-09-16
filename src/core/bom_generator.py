# 核心逻辑，处理并生成文件

from typing import Dict, Any
import pandas as pd


class BomGenerator:
    """BOM生成器类，用于处理Excel文件并生成BOM表
    
    该类主要用于读取包含产品研发明细信息的Excel文件，
    并提供根据款式编码查找产品信息的功能。
    
    支持的操作：
    - 读取Excel文件中的产品明细数据
    - 根据款式编码查找对应的产品信息
    - 返回包含波段、品类、开发颜色等关键信息的结构化数据
    
    Example:
        >>> generator = BomGenerator('product_details.xlsx')
        >>> info = generator.find_style_info('H5A123416')
        >>> print(info['品类'])  # 输出: 长袖T恤
    """
    
    # 类级别常量定义，避免魔术字符串
    SHEET_NAME = '明细表'
    STYLE_CODE_COL = '款式编码'
    WAVE_COL = '波段'
    CATEGORY_COL = '品类'
    DEV_COLOR_COL = '开发颜色'
    
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
                
        except FileNotFoundError:
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
