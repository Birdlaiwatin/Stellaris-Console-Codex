import os
import re
import xlsxwriter
from A配置路径 import *


def extract_planet_models():
    # 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 创建带版本号的Excel文件名
    excel_filename = f"{VERSION}星球模型列表.xlsx"
    excel_filepath = os.path.join(OUTPUT_DIR, excel_filename)

    # 创建Excel工作簿
    workbook = xlsxwriter.Workbook(excel_filepath)
    worksheet = workbook.add_worksheet('星球模型')

    # 设置表头格式
    header_format = workbook.add_format({
        'bold': True,
        'fg_color': '#D7E4BC',
        'border': 1
    })

    # 写入表头
    worksheet.write('A1', '星球模型名称', header_format)
    worksheet.set_column('A:A', 50)

    row = 1  # 从第2行开始写入数据

    # 遍历PLANETS_MODELS_DIR目录及其子目录
    for root, dirs, files in os.walk(PLANETS_MODELS_DIR):
        for file in files:
            if file.endswith('.asset'):
                file_path = os.path.join(root, file)
                print(f"正在处理文件: {file_path}")

                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    # 使用正则表达式查找所有entity块
                    entity_blocks = re.findall(r'entity\s*=\s*\{([^}]*?(?:\{[^}]*\}[^}]*?)*)\}', content, re.DOTALL)

                    for block in entity_blocks:
                        # 在每个entity块中查找最外层的name
                        name_match = re.search(r'name\s*=\s*"([^"]*)"', block)
                        if name_match:
                            entity_name = name_match.group(1)
                            worksheet.write(row, 0, entity_name)
                            row += 1
                            print(f"  找到实体: {entity_name}")

                except Exception as e:
                    print(f"处理文件 {file_path} 时出错: {e}")

    # 关闭工作簿
    workbook.close()
    print(f"提取完成！共找到 {row - 1} 个星球模型，结果已保存到 '{excel_filepath}'")


if __name__ == "__main__":
    extract_planet_models()