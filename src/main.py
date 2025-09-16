# 程序入口，GUI逻辑

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# 添加正确的模块路径，兼容开发环境和PyInstaller打包环境
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # PyInstaller 打包后的环境
    module_path = sys._MEIPASS
else:
    # 开发环境，从当前文件所在目录找到src目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    module_path = current_dir

if module_path not in sys.path:
    sys.path.insert(0, module_path)

from core.bom_generator import BomGenerator


class Application(tk.Tk):
    """BOM表自动生成工具的主应用程序窗口
    
    该类创建并管理图形用户界面，提供用户友好的操作界面来：
    - 选择源Excel文件
    - 选择输出目录
    - 触发BOM表生成操作
    - 显示操作状态和结果
    """
    
    def __init__(self):
        """初始化应用程序窗口"""
        super().__init__()
        
        # 设置窗口属性
        self.title("BOM表自动生成工具")
        self.geometry("500x300")
        
        # 初始化状态变量
        self.source_file_path = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        self.status_text = tk.StringVar(value="准备就绪")
        
        # 创建界面元素
        self._create_widgets()
    
    def _create_widgets(self):
        """创建并布局界面元素"""
        # 主框架
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 源文件选择区域
        source_frame = tk.Frame(main_frame)
        source_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(source_frame, text="源文件路径:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        
        source_entry_frame = tk.Frame(source_frame)
        source_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.source_entry = tk.Entry(source_entry_frame, textvariable=self.source_file_path, 
                                   state="readonly", font=("Arial", 9))
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        tk.Button(source_entry_frame, text="选择...", width=8, 
                 command=self._select_source_file).pack(side=tk.RIGHT)
        
        # 输出目录选择区域
        output_frame = tk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(output_frame, text="输出文件夹路径:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        
        output_entry_frame = tk.Frame(output_frame)
        output_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.output_entry = tk.Entry(output_entry_frame, textvariable=self.output_dir_path,
                                   state="readonly", font=("Arial", 9))
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        tk.Button(output_entry_frame, text="选择...", width=8,
                 command=self._select_output_dir).pack(side=tk.RIGHT)
        
        # 操作按钮区域
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 20))
        
        self.generate_button = tk.Button(button_frame, text="开始生成", 
                                       font=("Arial", 11, "bold"),
                                       bg="#4CAF50", fg="white",
                                       height=2, command=self._start_generation)
        self.generate_button.pack(fill=tk.X)
        
        # 状态栏
        status_frame = tk.Frame(main_frame, relief=tk.SUNKEN, bd=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_label = tk.Label(status_frame, textvariable=self.status_text,
                                   font=("Arial", 9), anchor=tk.W, padx=5, pady=2)
        self.status_label.pack(fill=tk.X)
    
    def _select_source_file(self):
        """选择源文件的回调方法"""
        file_path = filedialog.askopenfilename(
            title="选择源Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            defaultextension=".xlsx"
        )
        
        if file_path:
            self.source_file_path.set(file_path)
            self.status_text.set(f"已选择源文件: {os.path.basename(file_path)}")
        else:
            self.status_text.set("准备就绪")
    
    def _select_output_dir(self):
        """选择输出目录的回调方法"""
        dir_path = filedialog.askdirectory(
            title="选择输出文件夹"
        )
        
        if dir_path:
            self.output_dir_path.set(dir_path)
            self.status_text.set(f"已选择输出目录: {os.path.basename(dir_path)}")
        else:
            self.status_text.set("准备就绪")
    
    def _start_generation(self):
        """开始生成BOM表的回调方法"""
        # 获取用户选择的路径
        source_path = self.source_file_path.get().strip()
        output_path = self.output_dir_path.get().strip()
        
        # 输入验证
        if not source_path or not output_path:
            messagebox.showerror("错误", "请先选择源文件和输出文件夹！")
            return
        
        # 验证源文件是否存在
        if not os.path.exists(source_path):
            messagebox.showerror("错误", f"源文件不存在：\n{source_path}")
            return
        
        # 验证输出目录是否存在
        if not os.path.exists(output_path):
            messagebox.showerror("错误", f"输出目录不存在：\n{output_path}")
            return
        
        try:
            # 更新状态栏
            self.status_text.set("正在处理中，请稍候...")
            self.update()  # 强制更新界面
            
            # 禁用生成按钮防止重复点击
            self.generate_button.config(state="disabled")
            
            # 创建 BomGenerator 实例
            generator = BomGenerator(source_path)
            
            # 获取所有款式编码
            style_codes = generator.df[generator.STYLE_CODE_COL].dropna().unique()
            
            if len(style_codes) == 0:
                messagebox.showwarning("警告", "源文件中没有找到任何款式编码！")
                return
            
            # 循环生成每个款式的BOM文件
            success_count = 0
            failed_items = []
            
            for i, style_code in enumerate(style_codes):
                try:
                    # 更新进度状态
                    self.status_text.set(f"正在处理: {style_code} ({i+1}/{len(style_codes)})")
                    self.update()  # 强制更新界面
                    
                    # 生成BOM文件
                    generator.generate_bom_file(style_code, output_path)
                    success_count += 1
                    
                except Exception as e:
                    # 记录失败的款式编码和错误信息
                    failed_items.append(f"{style_code}: {str(e)}")
                    continue
            
            # 构建结果消息
            if success_count == len(style_codes):
                # 全部成功
                messagebox.showinfo("成功", 
                    f"处理完成！\n\n"
                    f"成功生成 {success_count} 个BOM文件。\n"
                    f"输出目录: {output_path}")
            elif success_count > 0:
                # 部分成功
                failed_msg = "\n".join(failed_items[:5])  # 只显示前5个错误
                if len(failed_items) > 5:
                    failed_msg += f"\n... 和其他 {len(failed_items) - 5} 个错误"
                
                messagebox.showwarning("部分成功", 
                    f"处理完成（部分成功）！\n\n"
                    f"成功: {success_count} 个\n"
                    f"失败: {len(failed_items)} 个\n\n"
                    f"失败详情:\n{failed_msg}")
            else:
                # 全部失败
                failed_msg = "\n".join(failed_items[:3])  # 只显示前3个错误
                if len(failed_items) > 3:
                    failed_msg += f"\n... 和其他 {len(failed_items) - 3} 个错误"
                
                messagebox.showerror("失败", 
                    f"处理失败！\n\n"
                    f"所有 {len(style_codes)} 个款式编码都处理失败。\n\n"
                    f"错误详情:\n{failed_msg}")
                    
        except Exception as e:
            # 捕获整体处理过程中的错误
            messagebox.showerror("发生错误", f"处理失败：\n\n{str(e)}")
            
        finally:
            # 恢复状态栏和按钮
            self.status_text.set("准备就绪")
            self.generate_button.config(state="normal")


if __name__ == "__main__":
    app = Application()
    app.mainloop()
