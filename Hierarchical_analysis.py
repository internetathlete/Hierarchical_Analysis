import os
import pandas as pd
import warnings
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

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
                column = col[0].column_letter  # Get the column name
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

def calculate_membership_levels(input_file_path, member_id_col, referrer_id_col, output_file_path=None):
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
    def get_level_and_path(member_id):  # 初始层级从0开始
        if member_id not in member_referrer_map or pd.isna(member_referrer_map[member_id]):
            return 0, [member_id]  # 返回当前层级和路径
        referrer_id = member_referrer_map[member_id]
        if referrer_id == member_id:  # 自己推荐自己，避免无限循环
            return 0, [member_id]
        level, path = get_level_and_path(referrer_id)  # 递归获取上游推荐人的层级和路径
        return level + 1, path + [member_id]  # 层级加1，并在路径中追加当前会员ID
    
    # 提示计算会员层级开始
    print("开始计算会员层级和上游路径...")
    
    # 为每个会员计算层级和上游路径
    for idx, row in df.iterrows():
        member_id = row[member_id_col]
        if df.at[idx, 'Level'] == -1:  # 如果还没有计算过层级
            level, path = get_level_and_path(member_id)
            df.at[idx, 'Level'] = level
            df.at[idx, 'Upstream_Path'] = ' -> '.join(map(str, path))  # 将路径转换为字符串表示
    
    # 提示计算下游人数开始
    print("开始计算下游人数...")
    
    # 创建推荐人ID到下游会员列表的映射
    referrer_downstream_map = {}
    for member_id, referrer_id in member_referrer_map.items():
        if pd.notna(referrer_id):
            if referrer_id not in referrer_downstream_map:
                referrer_downstream_map[referrer_id] = []
            referrer_downstream_map[referrer_id].append(member_id)
    
    # 为每个会员计算直接下游人数和总下游人数
    def calculate_downstream_count(member_id):
        if member_id not in referrer_downstream_map:
            return 0
        direct_downstream = len(referrer_downstream_map[member_id])
        df.loc[df[member_id_col] == member_id, 'Direct_Downstream_Count'] = direct_downstream  # 设置直接下游人数
        total_downstream = direct_downstream
        for downstream_id in referrer_downstream_map[member_id]:
            total_downstream += calculate_downstream_count(downstream_id)
        return total_downstream
    
    # 为每个会员计算总下游人数
    for idx, row in df.iterrows():
        member_id = row[member_id_col]
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

# 示例使用
if __name__ == "__main__":
    # 功能介绍
    print("欢迎使用层级架构分析程序！")
    print("本程序支持读取CSV、XLSX和XLS格式的会员数据文件，计算并输出会员层级、下游人数、直接下游人数和上游路径。")
    print("请按照提示输入所需信息。")
    
    input_file_path = input("请输入输入文件的路径（支持CSV、XLSX、XLS格式）: ").strip()
    output_file_path = input("请输入输出文件的路径（支持CSV、XLSX、XLS格式，直接回车使用源文件路径和格式）: ").strip()
    if not output_file_path:
        output_file_path = None

    member_id_col = input("请输入会员ID字段名称: ")
    referrer_id_col = input("请输入推荐人ID字段名称: ")

    calculate_membership_levels(input_file_path, member_id_col, referrer_id_col, output_file_path)
