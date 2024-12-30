import os
import pandas as pd
import warnings
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tqdm import tqdm  # 引入 tqdm 用于显示进度条
from graphviz import Digraph

# 忽略特定的 UserWarning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def read_file(file_path):
    """根据文件扩展名读取CSV、XLSX或XLS文件"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件未找到: {file_path}")

    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith('.xlsx'):
        try:
            return pd.read_excel(file_path, engine='openpyxl')
        except ImportError:
            raise ImportError("缺少 'openpyxl' 库。请使用 'pip install openpyxl' 来安装它。")
    elif file_path.endswith('.xls'):
        try:
            return pd.read_excel(file_path, engine='xlrd')
        except ImportError:
            raise ImportError("缺少 'xlrd' 库。请使用 'pip install xlrd' 来安装它。")
    else:
        raise ValueError("不支持的文件格式。请提供CSV、XLSX或XLS格式的文件。")

def write_file(df, file_path):
    """根据文件扩展名写入CSV、XLSX或XLS文件"""
    base_path, extension = os.path.splitext(file_path)
    index = 1
    while os.path.exists(file_path):
        file_path = f"{base_path}_{index}{extension}"
        index += 1
    
    # 处理数字列长度问题
    for col in df.columns:
        if df[col].dtype in ['int64', 'float64']:  # 针对数字列进行处理
            df[col] = df[col].apply(lambda x: str(x) if isinstance(x, (int, float)) and not pd.isna(x) and len(str(int(x))) > 11 else x)
    
    if file_path.endswith('.csv'):
        df.to_csv(file_path, index=False)
    elif file_path.endswith('.xlsx'):
        try:
            wb = Workbook()
            ws = wb.active
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            
            # 设置列格式
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter  # 获取列字母
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width
            
            # 直接保存文件
            wb.save(file_path)
        except ImportError:
            raise ImportError("缺少 'openpyxl' 库。请使用 'pip install openpyxl' 来安装它。")
    elif file_path.endswith('.xls'):
        try:
            df.to_excel(file_path, index=False, engine='xlwt')
        except ImportError:
            raise ImportError("缺少 'xlwt' 库。请使用 'pip install xlwt' 来安装它。")
    else:
        raise ValueError("不支持的文件格式。请提供CSV、XLSX或XLS格式的文件。")

def generate_hierarchy_chart(data, display_fields, output_file="hierarchy_chart"):
    """生成层级图，支持动态展示字段"""
    # 创建一个有向图，格式为 SVG
    dot = Digraph(format="svg", engine="dot")
    
    # 构建节点和边
    for parent, children in data.items():
        # 构建节点标签，显示字段的具体数据
        if parent == "":  # 处理没有推荐人的根节点
            parent_label = "无推荐人"
        else:
            parent_data = children.get('data', {})
            parent_label = "\n".join([f"{field}: {parent_data.get(field, '')}" for field in display_fields])
        dot.node(parent, label=parent_label if parent != "" else "无推荐人", shape="box")
        
        for child in children.get("children", []):
            child_data = data.get(child, {}).get('data', {})
            child_label = "\n".join([f"{field}: {child_data.get(field, '')}" for field in display_fields])
            dot.node(child, label=child_label, shape="box")
            dot.edge(parent, child)
    
    # 保存为 SVG 格式
    dot.render(output_file, cleanup=True)
    print(f"SVG 图已生成：{output_file}.svg")

def calculate_membership_levels(input_file_path, member_id_col, referrer_id_col, output_file_path=None, display_fields=None):
    """计算会员层级，并生成层级图"""
    # 读取文件
    print(f"读取文件: {input_file_path}")
    df = read_file(input_file_path)
    
    # 检查输入的列名是否在数据中
    if member_id_col not in df.columns or referrer_id_col not in df.columns:
        raise ValueError("提供的列名在数据中不存在，请检查输入的列名。")
    
    # 初始化层级列、下游人数列、上游路径列和直接下游人数列
    df['Level'] = -1
    df['Downstream_Count'] = 0
    df['Direct_Downstream_Count'] = 0  # 新增直接下游人数列
    df['Upstream_Path'] = ''  # 新增上游路径列
    
    # 创建会员ID到推荐人ID的映射
    member_referrer_map = dict(zip(df[member_id_col], df[referrer_id_col]))
    
    # 计算层级和上游路径
    def get_level_and_path(member_id):
        if member_id not in member_referrer_map or pd.isna(member_referrer_map[member_id]):
            return 0, [member_id]
        referrer_id = member_referrer_map[member_id]
        if referrer_id == member_id:
            return 0, [member_id]
        level, path = get_level_and_path(referrer_id)
        return level + 1, path + [member_id]
    
    print("开始计算会员层级和上游路径...")
    for idx in tqdm(df.index, desc="计算层级和路径", unit="行"):
        member_id = df.at[idx, member_id_col]
        if df.at[idx, 'Level'] == -1:
            level, path = get_level_and_path(member_id)
            df.at[idx, 'Level'] = level
            df.at[idx, 'Upstream_Path'] = ' -> '.join(map(str, path))
    
    print("开始计算下游人数...")
    referrer_downstream_map = {}
    for member_id, referrer_id in member_referrer_map.items():
        if pd.notna(referrer_id):
            if referrer_id not in referrer_downstream_map:
                referrer_downstream_map[referrer_id] = []
            referrer_downstream_map[referrer_id].append(member_id)
    
    def calculate_downstream_count(member_id):
        total_downstream = 0
        stack = [member_id]
        seen = set()
        
        while stack:
            current_member = stack.pop()
            if current_member in seen:
                continue
            seen.add(current_member)
            
            if current_member in referrer_downstream_map:
                direct_downstream = len(referrer_downstream_map[current_member])
                df.loc[df[member_id_col] == current_member, 'Direct_Downstream_Count'] = direct_downstream
                total_downstream += direct_downstream
                stack.extend(referrer_downstream_map[current_member])
        
        return total_downstream
    
    for idx in tqdm(df.index, desc="计算下游人数", unit="行"):
        member_id = df.at[idx, member_id_col]
        df.at[idx, 'Downstream_Count'] = calculate_downstream_count(member_id)
    
    # 如果没有指定输出路径，使用默认路径和文件名
    if not output_file_path:
        file_name, file_extension = os.path.splitext(input_file_path)
        output_file_path = f"{file_name}_with_levels{file_extension}"
    
    # 处理文件名重复的情况
    write_file(df, output_file_path)
    print(f"会员层级、下游人数、直接下游人数和上游路径计算完成，结果已保存到 {output_file_path}")
    
    # 打印总层级数和最高层级
    max_level = df['Level'].max()
    total_levels = df['Level'].nunique()
    print(f"一共有 {total_levels} 个层级，最高层级是 {max_level}")
    
    # 如果用户指定了要生成层级图
    if display_fields:
        # 构建层级数据结构以供图形化
        hierarchy_data = {}
        
        for idx in df.index:
            member_id = str(df.at[idx, member_id_col])  # 确保ID为字符串
            referrer_id = str(df.at[idx, referrer_id_col]) if pd.notna(df.at[idx, referrer_id_col]) else ""
            
            # 获取需要显示的字段数据
            member_data = {}
            for field in display_fields:
                if field in df.columns:
                    member_data[field] = df.at[idx, field]
                else:
                    member_data[field] = ""
            
            # 添加到层级数据中
            hierarchy_data.setdefault(referrer_id, {"data": {}, "children": []})
            hierarchy_data.setdefault(member_id, {"data": member_data, "children": []})
            if referrer_id != "":
                hierarchy_data[referrer_id]["children"].append(member_id)
    
        # 生成层级图
        generate_hierarchy_chart(hierarchy_data, display_fields)
    else:
        print("未指定要在层级图中显示的字段，跳过生成层级图。")

# 示例使用
if __name__ == "__main__":
    print("欢迎使用层级架构分析程序！")
    print("本程序支持读取CSV、XLSX和XLS格式的会员数据文件，计算并输出会员层级、下游人数、直接下游人数和上游路径，并生成层级图。")
    print("请按照提示输入所需信息。")
    
    input_file_path = input("请输入输入文件的路径（支持CSV、XLSX、XLS格式）: ").strip()
    while not os.path.exists(input_file_path):
        print("文件不存在，请重新输入。")
        input_file_path = input("请输入输入文件的路径（支持CSV、XLSX、XLS格式）: ").strip()
    
    output_file_path = input("请输入输出文件的路径（支持CSV、XLSX、XLS格式，直接回车使用源文件路径和格式）: ").strip()
    if not output_file_path:
        output_file_path = None

    member_id_col = input("请输入会员ID字段名称: ").strip()
    while member_id_col not in read_file(input_file_path).columns:
        print(f"字段 '{member_id_col}' 不存在，请重新输入。")
        member_id_col = input("请输入会员ID字段名称: ").strip()
    
    referrer_id_col = input("请输入推荐人ID字段名称: ").strip()
    while referrer_id_col not in read_file(input_file_path).columns:
        print(f"字段 '{referrer_id_col}' 不存在，请重新输入。")
        referrer_id_col = input("请输入推荐人ID字段名称: ").strip()
    
    display_fields_input = input("请输入要在层级图中显示的字段名（多个字段用空格或逗号隔开，默认为会员ID字段）：").strip()
    
    # 处理用户输入的显示字段
    if display_fields_input:
        display_fields = [field.strip() for field in display_fields_input.replace(",", " ").split()]
    else:
        display_fields = [member_id_col]  # 默认字段为会员ID字段
    
    # 检查输入的显示字段是否存在于数据中
    df_sample = read_file(input_file_path)
    invalid_fields = [field for field in display_fields if field not in df_sample.columns]
    if invalid_fields:
        print(f"以下字段在数据中不存在，将被忽略: {', '.join(invalid_fields)}")
        display_fields = [field for field in display_fields if field in df_sample.columns]
    
    if not display_fields:
        print("没有有效的字段用于生成层级图，跳过生成层级图。")
    else:
        print(f"将使用以下字段生成层级图: {', '.join(display_fields)}")
    
    # 执行层级计算的主函数
    calculate_membership_levels(input_file_path, member_id_col, referrer_id_col, output_file_path, display_fields)
