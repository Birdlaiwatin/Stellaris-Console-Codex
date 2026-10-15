import os
import re
import xlsxwriter
from pathlib import Path

# 导入配置路径
from A配置路径 import *


def clean_localisation_value(value):
    """清理本地化值，移除注释和多余引号"""
    # 移除注释（#后面的所有内容）
    if '#' in value:
        value = value.split('#')[0].strip()

    # 移除引号
    if value.startswith('"') and value.endswith('"'):
        value = value[1:-1]

    return value.strip()


def parse_localisation_files(localisation_dir):
    """解析本地化文件，构建翻译字典（递归搜索子文件夹）"""
    localisation_dict = {}

    for file_path in Path(localisation_dir).rglob("*.yml"):
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                for line in f:
                    line = line.strip()
                    if ':' in line and not line.startswith('#'):
                        # 分割键和值
                        parts = line.split(':', 1)
                        if len(parts) == 2:
                            key = parts[0].strip()
                            value = parts[1].strip()

                            # 清理值（移除注释和引号）
                            value = clean_localisation_value(value)

                            localisation_dict[key] = value
        except Exception as e:
            print(f"解析本地化文件时出错 {file_path}: {e}")

    return localisation_dict


def resolve_localisation(value, localisation_dict, depth=0, max_depth=10):
    """递归解析本地化值中的变量引用"""
    if depth > max_depth:
        return value  # 防止无限递归

    # 查找所有 $...$ 格式的变量引用
    pattern = r'\$([^$]+)\$'
    matches = re.findall(pattern, value)

    if not matches:
        return value  # 没有变量引用，直接返回

    # 替换所有变量引用
    for match in matches:
        # 递归解析引用的变量
        resolved_value = resolve_localisation(localisation_dict.get(match, match), localisation_dict, depth + 1,
                                              max_depth)
        # 替换原字符串中的引用
        value = value.replace(f'${match}$', resolved_value)

    return value


def parse_ship_size_files(ship_size_dir):
    """解析舰船类型文件，构建舰船类型字典"""
    ship_size_dict = {}

    for file_path in Path(ship_size_dir).glob("*.txt"):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

                # 查找所有舰船类型定义块
                pattern = r'(\w+)\s*=\s*\{'
                matches = list(re.finditer(pattern, content))

                for match in matches:
                    ship_name = match.group(1)
                    start_pos = match.start()

                    # 找到匹配的结束大括号
                    brace_count = 0
                    in_quotes = False
                    escape_next = False

                    for i in range(start_pos, len(content)):
                        char = content[i]

                        # 处理转义字符
                        if escape_next:
                            escape_next = False
                            continue

                        if char == '\\':
                            escape_next = True
                            continue

                        # 处理引号内的内容
                        if char == '"' and not escape_next:
                            in_quotes = not in_quotes
                            continue

                        if not in_quotes:
                            if char == '{':
                                brace_count += 1
                            elif char == '}':
                                brace_count -= 1
                                if brace_count == 0:
                                    # 找到了完整的定义块
                                    ship_block = content[start_pos:i + 1]
                                    break
                    else:
                        # 如果没有找到匹配的结束大括号，跳过这个定义
                        continue

                    # 在完整的ship_block中查找class
                    class_match = re.search(r'class\s*=\s*(\w+)', ship_block)
                    ship_class = class_match.group(1) if class_match else None

                    # 查找max_speed
                    speed_match = re.search(r'max_speed\s*=\s*([\d.]+)', ship_block)
                    max_speed = speed_match.group(1) if speed_match else None

                    # 只提取graphical_culture = { ... } 格式的内容
                    gc_match = re.search(r'graphical_culture\s*=\s*\{([^}]+)\}', ship_block)
                    if gc_match:
                        # 提取大括号内的内容，保留原始格式
                        graphical_culture = gc_match.group(1).strip()
                        # 只规范化空格，保留引号
                        graphical_culture = re.sub(r'\s+', ' ', graphical_culture).strip()
                    else:
                        # 如果没有找到graphical_culture = { ... }，则使用默认值root
                        graphical_culture = "root"

                    ship_size_dict[ship_name] = {
                        'class': ship_class,
                        'max_speed': max_speed,
                        'graphical_culture': graphical_culture
                    }

        except Exception as e:
            print(f"解析舰船类型文件时出错 {file_path}: {e}")

    return ship_size_dict


def parse_ship_design_files(designs_dir):
    """解析舰船设计文件，提取ship_design块"""
    ship_designs = []

    for file_path in Path(designs_dir).glob("*.txt"):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

                # 查找所有ship_design块
                design_blocks = re.findall(r'ship_design\s*=\s*\{([^}]+?(?:\{[^}]*\}[^}]*?)*)\}', content, re.DOTALL)

                for block in design_blocks:
                    # 提取name
                    name_match = re.search(r'name\s*=\s*"([^"]*)"', block)
                    name = name_match.group(1) if name_match else None

                    # 提取ship_size
                    size_match = re.search(r'ship_size\s*=\s*(\w+)', block)
                    ship_size = size_match.group(1) if size_match else None

                    if name and ship_size:
                        ship_designs.append({
                            'name': name,
                            'ship_size': ship_size
                        })
        except Exception as e:
            print(f"解析舰船设计文件时出错 {file_path}: {e}")

    return ship_designs


def classify_ship_type(ship_class, max_speed, localisation_dict):
    """根据舰船类型和最大速度对舰船进行分类"""
    if not ship_class:
        return "未知类型"

    # 获取舰船类型的本地化翻译
    class_translation = resolve_localisation(localisation_dict.get(ship_class, ship_class), localisation_dict)

    # 对防御站类型进行细化
    if ship_class == "shipclass_military_station":
        if max_speed and float(max_speed) == 0:
            return "防御站"
        else:
            return "防御平台"

    return class_translation


def categorize_ship_types(ship_types):
    """将舰船类型按照指定类别进行分类"""
    categories = {
        "军事舰船类": ["军事舰船", "巨像"],
        "恒星基地类": ["恒星基地"],
        "防御平台与防御站类": ["防御平台", "防御站"],
        "民用舰船类": ["科研船", "工程船", "殖民船"],
        "其他类": []
    }

    # 将舰船类型分类
    categorized = {category: [] for category in categories}

    for ship_type in ship_types:
        found = False
        for category, types in categories.items():
            if ship_type in types:
                categorized[category].append(ship_type)
                found = True
                break

        # 如果没有找到对应类别，放入其他类
        if not found:
            categorized["其他类"].append(ship_type)

    # 对每个类别内的舰船类型进行排序
    for category, types in categories.items():
        if category in categorized:
            # 按照categories中定义的顺序排序
            sorted_types = []
            for t in types:
                if t in categorized[category]:
                    sorted_types.append(t)
            categorized[category] = sorted_types

    # 对其他类按字母顺序排序
    categorized["其他类"].sort()

    return categorized


def generate_output(ship_designs, ship_size_dict, localisation_dict):
    """生成输出结果"""
    output_lines = []
    excel_data = []

    # 收集所有舰船类型
    ship_types = set()
    for design in ship_designs:
        ship_size_key = design['ship_size']
        ship_size_info = ship_size_dict.get(ship_size_key, {})
        ship_class = ship_size_info.get('class')
        max_speed = ship_size_info.get('max_speed')

        ship_type = classify_ship_type(ship_class, max_speed, localisation_dict)
        ship_types.add(ship_type)

    # 对舰船类型进行分类
    categorized_types = categorize_ship_types(ship_types)

    # 添加舰船类型汇总信息
    output_lines.append("已查找到以下舰船类型：")
    for category, types in categorized_types.items():
        if types:
            output_lines.append(f"  {category}：{', '.join(types)}")
    output_lines.append("")  # 空行分隔

    # 添加舰船设计详情
    for design in ship_designs:
        name_key = design['name']
        ship_size_key = design['ship_size']

        # 获取舰船名称翻译，并解析其中的变量引用
        ship_name_translation = resolve_localisation(localisation_dict.get(name_key, name_key), localisation_dict)

        # 获取舰船尺寸的本地化翻译，并解析其中的变量引用
        ship_size_translation = resolve_localisation(localisation_dict.get(ship_size_key, ship_size_key),
                                                     localisation_dict)

        # 获取舰船类型信息
        ship_size_info = ship_size_dict.get(ship_size_key, {})
        ship_class = ship_size_info.get('class')
        max_speed = ship_size_info.get('max_speed')
        graphical_culture = ship_size_info.get('graphical_culture', 'root')

        # 获取舰船类型的本地化翻译，并解析其中的变量引用
        ship_type = classify_ship_type(ship_class, max_speed, localisation_dict)

        # 构建完整的舰船名称
        full_ship_name = f"{ship_name_translation}级{ship_size_translation}"

        # 添加到输出
        output_lines.append(f"全局舰船KEY：\"{name_key}\"")
        output_lines.append(f"舰船名称：{full_ship_name}")
        output_lines.append(f"舰船类型：{ship_type}")
        output_lines.append(f"模型类别：{graphical_culture}")
        output_lines.append(f"舰船组件：{ship_size_key}")
        output_lines.append("")  # 空行分隔

        # 添加到Excel数据
        excel_data.append({
            "全局舰船KEY": name_key,
            "舰船名称": full_ship_name,
            "舰船类型": ship_type,
            "模型类别": graphical_culture,
            "舰船组件": ship_size_key
        })

    return output_lines, excel_data, categorized_types


def save_to_excel(excel_data, categorized_types, output_dir):
    """将数据保存到Excel文件，使用xlsxwriter库"""
    # 如果没有数据，直接返回
    if not excel_data:
        print("没有数据可导出到Excel")
        return

    # 创建分类排序映射
    category_order = {
        "军事舰船类": 1,
        "恒星基地类": 2,
        "防御平台与防御站类": 3,
        "民用舰船类": 4,
        "其他类": 5
    }

    # 定义每个大类内部的小类顺序
    subcategory_order = {
        "军事舰船类": ["军事舰船", "巨像"],
        "恒星基地类": ["恒星基地"],
        "防御平台与防御站类": ["防御平台", "防御站"],
        "民用舰船类": ["科研船", "工程船", "殖民船"],
        "其他类": []  # 其他类内部按原始顺序
    }

    # 为每个数据项添加分类和排序字段
    for i, item in enumerate(excel_data):
        ship_type = item["舰船类型"]
        category = "其他类"

        for cat, types in categorized_types.items():
            if ship_type in types:
                category = cat
                break

        item["分类"] = category
        item["分类排序"] = category_order.get(category, 99)

        # 添加小类排序
        if category in subcategory_order and ship_type in subcategory_order[category]:
            item["小类排序"] = subcategory_order[category].index(ship_type)
        else:
            item["小类排序"] = 99

        # 添加原始顺序索引，保持与TXT文件相同的顺序
        item["原始顺序"] = i

    # 按分类、小类和原始顺序排序
    excel_data.sort(key=lambda x: (x["分类排序"], x["小类排序"], x["原始顺序"]))

    # 创建Excel文件（使用版本号命名）
    excel_file = os.path.join(output_dir, f"{VERSION}全局舰船设计.xlsx")
    workbook = xlsxwriter.Workbook(excel_file)

    # 定义单元格格式
    header_format = workbook.add_format({
        'bold': True,
        'fg_color': '#D7E4BC',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    cell_format = workbook.add_format({
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'
    })

    # 保存完整数据到第一个工作表
    worksheet_all = workbook.add_worksheet('所有舰船')

    # 写入表头
    headers = ["全局舰船KEY", "舰船名称", "舰船类型", "模型类别", "舰船组件", "分类"]
    for col, header in enumerate(headers):
        worksheet_all.write(0, col, header, header_format)

    # 写入数据
    for row, item in enumerate(excel_data, start=1):
        worksheet_all.write(row, 0, item["全局舰船KEY"], cell_format)
        worksheet_all.write(row, 1, item["舰船名称"], cell_format)
        worksheet_all.write(row, 2, item["舰船类型"], cell_format)
        worksheet_all.write(row, 3, item["模型类别"], cell_format)
        worksheet_all.write(row, 4, item["舰船组件"], cell_format)
        worksheet_all.write(row, 5, item["分类"], cell_format)

    # 设置列宽
    worksheet_all.set_column(0, 0, 20)  # 全局舰船KEY
    worksheet_all.set_column(1, 1, 30)  # 舰船名称
    worksheet_all.set_column(2, 2, 15)  # 舰船类型
    worksheet_all.set_column(3, 3, 20)  # 模型类别
    worksheet_all.set_column(4, 4, 15)  # 舰船组件
    worksheet_all.set_column(5, 5, 15)  # 分类

    # 按分类保存到不同工作表
    for category in category_order.keys():
        if category in categorized_types and categorized_types[category]:
            # 筛选该分类的数据
            category_data = [item for item in excel_data if item["分类"] == category]

            if category_data:
                # 创建工作表
                worksheet = workbook.add_worksheet(category)

                # 写入表头（不包含分类列）
                filtered_headers = [h for h in headers if h != "分类"]
                for col, header in enumerate(filtered_headers):
                    worksheet.write(0, col, header, header_format)

                # 写入数据
                for row, item in enumerate(category_data, start=1):
                    worksheet.write(row, 0, item["全局舰船KEY"], cell_format)
                    worksheet.write(row, 1, item["舰船名称"], cell_format)
                    worksheet.write(row, 2, item["舰船类型"], cell_format)
                    worksheet.write(row, 3, item["模型类别"], cell_format)
                    worksheet.write(row, 4, item["舰船组件"], cell_format)

                # 设置列宽
                worksheet.set_column(0, 0, 20)  # 全局舰船KEY
                worksheet.set_column(1, 1, 30)  # 舰船名称
                worksheet.set_column(2, 2, 15)  # 舰船类型
                worksheet.set_column(3, 3, 20)  # 模型类别
                worksheet.set_column(4, 4, 15)  # 舰船组件

    # 关闭工作簿
    workbook.close()

    print(f"Excel文件已保存到: {excel_file}")


def main():
    # 创建输出目录
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("正在解析本地化文件...")
    localisation_dict = parse_localisation_files(LOCALISATION_DIR)
    print(f"已加载 {len(localisation_dict)} 个本地化条目")

    print("正在解析舰船类型文件...")
    ship_size_dict = parse_ship_size_files(SHIPSIZE_DIR)
    print(f"已加载 {len(ship_size_dict)} 个舰船类型")

    print("正在解析舰船设计文件...")
    ship_designs = parse_ship_design_files(DESIGNS_DIR)
    print(f"已找到 {len(ship_designs)} 个舰船设计")

    print("正在生成输出...")
    output_lines, excel_data, categorized_types = generate_output(ship_designs, ship_size_dict, localisation_dict)

    # 写入输出文件（使用版本号命名）
    output_file = os.path.join(OUTPUT_DIR, f"{VERSION}全局舰船设计.txt")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_lines))

    print(f"完成！TXT结果已保存到: {output_file}")

    # 保存到Excel
    save_to_excel(excel_data, categorized_types, OUTPUT_DIR)


if __name__ == "__main__":
    main()