import os
import re
import xlsxwriter
from A配置路径 import *


class ArmyExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.armies = []

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    def load_localization_files(self):
        """加载所有本地化文件"""
        print("正在加载本地化文件...")
        localization_files = []

        for root, dirs, files in os.walk(LOCALISATION_DIR):
            for file in files:
                if file.endswith('.yml'):
                    localization_files.append(os.path.join(root, file))

        for file_path in localization_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                lines = content.split('\n')
                current_key = None
                current_value = []

                for line in lines:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue

                    if ':' in line and not line.startswith(' '):
                        if current_key and current_value:
                            value_str = ' '.join(current_value).strip()
                            if '"' in value_str:
                                value_str = value_str.split('"')[1] if value_str.count('"') >= 2 else value_str
                            self.localization_dict[current_key] = value_str

                        parts = line.split(':', 1)
                        current_key = parts[0].strip()
                        current_value = [parts[1].strip()] if len(parts) > 1 else []
                    elif current_key and line:
                        current_value.append(line.strip())

                if current_key and current_value:
                    value_str = ' '.join(current_value).strip()
                    if '"' in value_str:
                        value_str = value_str.split('"')[1] if value_str.count('"') >= 2 else value_str
                    self.localization_dict[current_key] = value_str

            except Exception as e:
                print(f"加载本地化文件失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.localization_dict)} 个本地化条目")

    def process_dollar_replacement(self, text, max_depth=5, current_depth=0):
        """处理$xxx$替换符"""
        if current_depth >= max_depth:
            return text  # 返回文本

        # 查找所有$xxx$替换符
        dollar_matches = re.findall(r'\$([^$]+)\$', text)
        if not dollar_matches:
            return text

        for match in dollar_matches:
            # 查找替换符对应的本地化
            if match in self.localization_dict:
                replacement = self.localization_dict[match]
                # 递归处理替换后的文本
                replacement = self.process_dollar_replacement(replacement, max_depth, current_depth + 1)
                text = text.replace(f'${match}$', replacement)
            else:
                # 如果找不到对应的本地化，保留原样
                pass

        return text

    def extract_balanced_braces(self, text, start_pos):
        """提取平衡的花括号内容"""
        if start_pos >= len(text) or text[start_pos] != '{':
            return None, start_pos

        brace_count = 1
        pos = start_pos + 1
        start = start_pos

        while pos < len(text) and brace_count > 0:
            if text[pos] == '{':
                brace_count += 1
            elif text[pos] == '}':
                brace_count -= 1
            pos += 1

        if brace_count == 0:
            return text[start:pos], pos
        else:
            return None, start_pos

    def extract_complete_blocks(self, content):
        """提取完整的陆军定义块，正确处理嵌套花括号"""
        blocks = []
        lines = content.split('\n')

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 寻找陆军定义开始
            if line and '=' in line and '{' in line and not line.startswith('#'):
                parts = line.split('=', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    # 检查是否包含开括号
                    if '{' in parts[1]:
                        # 开始提取完整块
                        block_lines = []
                        brace_count = 0
                        in_block = False

                        # 收集整个块的内容
                        j = i
                        while j < len(lines):
                            current_line = lines[j]
                            block_lines.append(current_line)

                            # 统计花括号
                            for char in current_line:
                                if char == '{':
                                    brace_count += 1
                                    in_block = True
                                elif char == '}':
                                    brace_count -= 1

                            j += 1

                            # 当所有花括号匹配完成时，块结束
                            if in_block and brace_count == 0:
                                break

                        # 处理收集到的块
                        if in_block and brace_count == 0:
                            block_content = '\n'.join(block_lines)
                            # 提取键和内容
                            match = re.match(r'(\w+)\s*=\s*\{', block_content)
                            if match:
                                block_key = match.group(1)
                                # 提取内容部分（去掉外层花括号）
                                content_start = block_content.find('{') + 1
                                content_end = block_content.rfind('}')
                                block_content_clean = block_content[content_start:content_end].strip()

                                blocks.append((block_key, block_content_clean))

                            i = j  # 跳过已处理的块
                            continue
            i += 1

        return blocks

    def parse_army_file(self, file_path):
        """解析单个陆军文件"""
        file_name = os.path.basename(file_path)
        print(f"正在解析文件: {file_name}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 提取完整的陆军定义
            armies = self.extract_complete_blocks(content)

            for army_key, army_content in armies:
                print(f"  解析陆军: {army_key}")
                army_data = self.parse_single_army(army_key, army_content, file_name)
                if army_data:
                    self.armies.append(army_data)

        except Exception as e:
            print(f"解析文件失败: {file_path}, 错误: {e}")

    def parse_single_army(self, army_key, army_content, source_file):
        """解析单个陆军定义"""
        army_data = {
            'key': army_key,
            'name': '',
            'attributes': '',
            'source': source_file
        }

        # 获取陆军名称
        raw_name = self.localization_dict.get(army_key, army_key)
        # 处理$$替换符
        processed_name = self.process_dollar_replacement(raw_name)
        army_data['name'] = processed_name

        # 构建陆军属性字符串
        attributes = []

        # 检查是否为防御部队
        defensive_match = re.search(r'defensive\s*=\s*yes', army_content)
        if defensive_match:
            attributes.append("防御部队")

        # 检查是否无法生成
        pop_spawned_match = re.search(r'is_pop_spawned\s*=\s*yes', army_content)
        if pop_spawned_match:
            attributes.append("无法生成")

        # 将属性列表转换为字符串
        army_data['attributes'] = '，'.join(attributes)

        return army_data

    def parse_all_armies(self):
        """解析所有陆军文件"""
        print("正在解析陆军文件...")

        if not os.path.exists(ARMIES_DIR):
            print(f"陆军目录不存在: {ARMIES_DIR}")
            return

        for file_name in os.listdir(ARMIES_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(ARMIES_DIR, file_name)
                self.parse_army_file(file_path)

        print(f"已解析 {len(self.armies)} 个陆军")

    def generate_excel(self):
        """生成Excel文件"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}陆军信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('陆军信息')

        # 定义格式
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        cell_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter'
        })

        # 设置列宽
        worksheet.set_column('A:A', 30)  # 陆军名称
        worksheet.set_column('B:B', 30)  # 陆军KEY
        worksheet.set_column('C:C', 30)  # 陆军属性

        # 写入表头
        headers = ['陆军名称', '陆军KEY', '陆军属性']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for army in self.armies:
            worksheet.write(row, 0, army['name'], cell_format)
            worksheet.write(row, 1, army['key'], cell_format)
            worksheet.write(row, 2, army['attributes'], cell_format)
            row += 1

        workbook.close()
        print(f"Excel文件已生成: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取陆军信息...")

        self.load_localization_files()
        self.parse_all_armies()
        self.generate_excel()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = ArmyExtractor()
    extractor.run()