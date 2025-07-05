# 导入必要的库
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os

# 定义主窗口
class JSONExcelTool:
    def __init__(self, root):
        self.root = root
        self.root.title("原子抽提字段输出解析")

        # 输入文件路径
        self.input_label = tk.Label(root, text="输入文件路径:")
        self.input_label.pack(pady=5)
        
        self.input_entry = tk.Entry(root, width=50)
        self.input_entry.pack(pady=5)
        
        self.input_button = tk.Button(root, text="选择文件", command=self.select_input_file)
        self.input_button.pack(pady=5)

        # 输出文件路径
        self.output_label = tk.Label(root, text="输出文件路径:")
        self.output_label.pack(pady=5)
        
        self.output_entry = tk.Entry(root, width=50)
        self.output_entry.pack(pady=5)
        
        self.output_button = tk.Button(root, text="选择输出位置并命名文件", command=self.select_output_path)
        self.output_button.pack(pady=5)

        # 执行按钮
        self.execute_button = tk.Button(root, text="执行转换", command=self.execute_conversion)
        self.execute_button.pack(pady=20)

    def select_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file_path)

    def select_output_path(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)

    def execute_conversion(self):
        input_file = self.input_entry.get()
        output_file = self.output_entry.get()

        if not input_file or not output_file:
            messagebox.showerror("错误", "请填写输入和输出路径")
            return

        try:
            # 定义输出列的顺序
            output_columns = [
                '公开(公告)号', '标题', '摘要', '权利要求', 'IPC分类号', '说明书',
                '技术标题', '改进及应用',
                '改进对象词集合', '改进对象词1', '改进对象词2', '改进对象词3',
                '应用产品词集合',
                '应用产品词1', '应用产品词1来源片段',
                '应用产品词2', '应用产品词2来源片段',
                '应用产品词3', '应用产品词3来源片段',
                '核心问题', '技术问题内容', '技术问题来源片段', '对象及缺陷', '技术问题对象词', '技术问题词', '技术问题条件短语',
                '技术手段', '技术手段内容', '技术手段片段来源', '技术手段对象词', '技术手段关键词', '技术手段类型',
                'Triz发明原理', '技术效果', '技术效果内容', '技术效果来源片段',
                '对象及效果',  '技术效果对象词', '技术效果词',
                '次要问题', '次要问题_技术问题内容', '次要问题_技术问题来源片段', '次要问题_对象及缺陷', '次要问题_技术问题条件短语',
                '次要问题_技术手段内容', '次要问题_技术手段片段来源', '次要问题_技术手段对象词', '次要问题_技术手段关键词', '次要问题_技术手段类型',
                '次要问题_技术效果内容', '次要问题_技术效果来源片段',
                '次要问题_对象及效果', '次要问题_技术效果对象词', '次要问题_技术效果词'
            ]

            def safe_get(data, keys, default=""):
                """安全获取嵌套字典的值"""
                if not isinstance(keys, (list, tuple)):
                    keys = [keys]
                current = data
                for key in keys:
                    if isinstance(current, dict) and key in current:
                        current = current[key]
                    elif isinstance(current, list) and isinstance(key, int) and key < len(current):
                        current = current[key]
                    else:
                        return default
                return current if isinstance(current, (str, int, float, list, dict)) else default

            def format_list(data, separator="、"):
                """格式化列表为字符串"""
                if isinstance(data, list):
                    return separator.join([str(item) for item in data if item])
                return str(data) if data else ""

            def format_source_fragments(fragments):
                """格式化来源片段"""
                if isinstance(fragments, list):
                    return "\n".join([str(frag) for frag in fragments if frag])
                return str(fragments) if fragments else ""

            # 读取原始数据
            df = pd.read_excel(input_file)
            
            # 查找所有表头包含“输出”的列
            output_columns_in_df = [col for col in df.columns if "输出" in col]
            
            processed_data = []
            for index, row in df.iterrows():
                # 初始化行数据，复制当前行的所有数据
                row_data = row.to_dict()
                # 确保输出列都存在，不存在的初始化为空字符串
                for col in output_columns:
                    if col not in row_data:
                        row_data[col] = ""
                
                # 依次解析所有包含“输出”的列
                for output_col in output_columns_in_df:
                    # 解析JSON数据
                    try:
                        json_data = json.loads(str(row[output_col]).strip())
                    except:
                        json_data = {}
                    
                    # 填充技术标题和改进及应用
                    row_data['技术标题'] = safe_get(json_data, "技术标题")
                    row_data['改进及应用'] = json.dumps(safe_get(json_data, ["改进及应用"], []), ensure_ascii=False)
                    
                    # 处理改进对象词
                    improvements = safe_get(json_data, ["改进及应用"], [])
                    improvement_objects = [safe_get(imp, "改进对象词") for imp in improvements]
                    row_data['改进对象词集合'] = format_list(improvement_objects)
                    for i in range(1, 4):
                        if i <= len(improvements):
                            row_data[f'改进对象词{i}'] = safe_get(improvements[i-1], "改进对象词")
                    
                    # 处理应用产品词
                    app_products = safe_get(json_data, ["改进及应用"], [])
                    row_data['应用产品词集合'] = format_list([format_list(safe_get(p, "应用产品词")) for p in app_products])
                    for i in range(1, 4):
                        if i <= len(app_products):
                            product = app_products[i-1]
                            row_data[f'应用产品词{i}'] = format_list(safe_get(product, "应用产品词"))
                            row_data[f'应用产品词{i}来源片段'] = safe_get(product, "应用产品来源片段")
                    
                    # 处理核心问题
                    core_issue = safe_get(json_data, ["核心问题"], {})
                    row_data['核心问题'] = json.dumps(core_issue, ensure_ascii=False)
                    row_data['技术问题内容'] = safe_get(core_issue, "技术问题内容")
                    row_data['技术问题来源片段'] = format_source_fragments(safe_get(core_issue, "技术问题来源片段"))
                    
                    # 处理对象及缺陷
                    obj_defects = safe_get(core_issue, ["对象及缺陷"], [{}])
                    if not isinstance(obj_defects, list):
                        obj_defects = [obj_defects]
                    obj_defect = obj_defects[0] if obj_defects else {}
                    row_data['对象及缺陷'] = json.dumps(obj_defects, ensure_ascii=False)
                    row_data['技术问题对象词'] = format_list(safe_get(obj_defect, "技术问题对象词"))
                    row_data['技术问题词'] = format_list(safe_get(obj_defect, "技术问题词"))
                    row_data['技术问题条件短语'] = safe_get(core_issue, "技术问题条件短语")
                    
                    # 处理技术手段
                    tech_methods = safe_get(core_issue, ["技术手段"], [{}])
                    if not isinstance(tech_methods, list):
                        tech_methods = [tech_methods]
                    tech_method = tech_methods[0] if tech_methods else {}
                    row_data['技术手段'] = json.dumps(tech_methods, ensure_ascii=False)
                    row_data['技术手段内容'] = safe_get(tech_method, "技术手段内容")
                    row_data['技术手段片段来源'] = format_source_fragments(safe_get(tech_method, "技术手段来源片段"))
                    row_data['技术手段对象词'] = format_list(safe_get(tech_method, "技术手段对象词"))
                    row_data['技术手段关键词'] = format_list(safe_get(tech_method, "技术手段关键词"))
                    row_data['技术手段类型'] = format_list(safe_get(tech_method, "技术手段类型"))
                    
                    # 处理Triz发明原理
                    triz_principles = safe_get(tech_method, ["triz发明原理"], [])
                    row_data['Triz发明原理'] = json.dumps(triz_principles, ensure_ascii=False)
                    
                    # 处理技术效果
                    tech_effects = safe_get(tech_method, ["技术效果"], [{}])
                    if not isinstance(tech_effects, list):
                        tech_effects = [tech_effects]
                    tech_effect = tech_effects[0] if tech_effects else {}
                    row_data['技术效果'] = json.dumps(tech_effects, ensure_ascii=False)
                    row_data['技术效果内容'] = safe_get(tech_effect, "技术效果内容")
                    row_data['技术效果来源片段'] = format_source_fragments(safe_get(tech_effect, "技术效果来源片段"))
                    
                    # 处理对象及效果
                    obj_effects = safe_get(tech_effect, ["对象及效果"], [])
                    row_data['对象及效果'] = json.dumps(obj_effects, ensure_ascii=False)  # 填充对象及效果字段
                    
                    # 只取第一组技术效果对象词和技术效果词
                    first_obj_effect = obj_effects[0] if obj_effects else {}
                    row_data['技术效果对象词'] = format_list(safe_get(first_obj_effect, "技术效果对象词"))
                    row_data['技术效果词'] = format_list(safe_get(first_obj_effect, "技术效果词"))
                    
                    # 处理次要问题
                    secondary_issues = safe_get(json_data, ["次要问题"], [{}])
                    if not isinstance(secondary_issues, list):
                        secondary_issues = [secondary_issues]
                    sec_issue = secondary_issues[0] if secondary_issues else {}
                    row_data['次要问题'] = json.dumps(secondary_issues, ensure_ascii=False)
                    row_data['次要问题_技术问题内容'] = safe_get(sec_issue, "技术问题内容")
                    row_data['次要问题_技术问题来源片段'] = format_source_fragments(safe_get(sec_issue, "技术问题来源片段"))
                    
                    # 次要问题对象及缺陷
                    sec_obj_defects = safe_get(sec_issue, ["对象及缺陷"], [{}])
                    if not isinstance(sec_obj_defects, list):
                        sec_obj_defects = [sec_obj_defects]
                    sec_obj_defect = sec_obj_defects[0] if sec_obj_defects else {}
                    row_data['次要问题_对象及缺陷'] = json.dumps(sec_obj_defects, ensure_ascii=False)
                    row_data['次要问题_技术问题条件短语'] = safe_get(sec_issue, "技术问题条件短语")
                    
                    # 次要问题技术手段
                    sec_tech_methods = safe_get(sec_issue, ["技术手段"], [{}])
                    if not isinstance(sec_tech_methods, list):
                        sec_tech_methods = [sec_tech_methods]
                    sec_tech_method = sec_tech_methods[0] if sec_tech_methods else {}
                    row_data['次要问题_技术手段内容'] = safe_get(sec_tech_method, "技术手段内容")
                    row_data['次要问题_技术手段片段来源'] = format_source_fragments(safe_get(sec_tech_method, "技术手段来源片段"))
                    row_data['次要问题_技术手段对象词'] = format_list(safe_get(sec_tech_method, "技术手段对象词"))
                    row_data['次要问题_技术手段关键词'] = format_list(safe_get(sec_tech_method, "技术手段关键词"))
                    row_data['次要问题_技术手段类型'] = format_list(safe_get(sec_tech_method, "技术手段类型"))
                    
                    # 次要问题技术效果
                    sec_tech_effects = safe_get(sec_tech_method, ["技术效果"], [{}])
                    if not isinstance(sec_tech_effects, list):
                        sec_tech_effects = [sec_tech_effects]
                    sec_tech_effect = sec_tech_effects[0] if sec_tech_effects else {}
                    row_data['次要问题_技术效果内容'] = safe_get(sec_tech_effect, "技术效果内容")
                    row_data['次要问题_技术效果来源片段'] = format_source_fragments(safe_get(sec_tech_effect, "技术效果来源片段"))
                    
                    # 次要问题对象及效果
                    sec_obj_effects = safe_get(sec_tech_effect, ["对象及效果"], [])
                    row_data['次要问题_对象及效果'] = json.dumps(sec_obj_effects, ensure_ascii=False)
                    
                    # 次要问题只取第一组技术效果对象词和技术效果词
                    sec_first_obj_effect = sec_obj_effects[0] if sec_obj_effects else {}
                    row_data['次要问题_技术效果对象词'] = format_list(safe_get(sec_first_obj_effect, "技术效果对象词"))
                    row_data['次要问题_技术效果词'] = format_list(safe_get(sec_first_obj_effect, "技术效果词"))
                    
                processed_data.append(row_data)

            # 生成并保存表格
            df_output = pd.DataFrame(processed_data, columns=output_columns)
            df_output.to_excel(output_file, index=False, engine='openpyxl')
            
            messagebox.showinfo("成功", f"处理完成！输出文件已保存为: {output_file}")
        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = JSONExcelTool(root)
    root.mainloop()