import os
import re
import xlsxwriter

# 导入配置文件
from A配置路径 import *


class DistrictExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_loc_dict = {}
        self.districts = []
        self.insensitive_localization_dict = {}

        # 调试信息
        self.debug_info = {
            'missing_district_name': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    def preprocess_localization_text(self, text):
        """预处理本地化文本"""
        if not text:
            return text

        # 1. 递归处理$$替换符
        text, _ = self.process_dollar_replacement(text)

        # 2. 清除颜色标识符 §Yxxx§!
        text = re.sub(r'§[A-Z](.*?)§!', r'\1', text)

        return text

    def clean_whitespace_text(self, text):
        """清理空白文本"""
        if not text:
            return text

        # 替换多个连续空格为单个空格
        text = re.sub(r'\s+', ' ', text)
        # 去除首尾空格
        text = text.strip()

        return text

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
                            value_str = self.preprocess_localization_text(value_str)
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
                    value_str = self.preprocess_localization_text(value_str)
                    self.localization_dict[current_key] = value_str

            except Exception as e:
                print(f"加载本地化文件失败: {file_path}, 错误: {e}")

        # 构建大小写不敏感的本地化字典
        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localization_dict.items()}
        print(f"已加载 {len(self.localization_dict)} 个本地化条目")

    def load_scripted_loc(self):
        """加载脚本本地化"""
        print("正在加载脚本本地化...")

        if not os.path.exists(SCRIPTED_LOC_DIR):
            print(f"脚本本地化目录不存在: {SCRIPTED_LOC_DIR}")
            return

        for file_name in os.listdir(SCRIPTED_LOC_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(SCRIPTED_LOC_DIR, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    # 使用更强大的方法提取defined_text块，处理嵌套花括号
                    self.extract_defined_text_blocks(content)

                except Exception as e:
                    print(f"加载脚本本地化文件失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.scripted_loc_dict)} 个脚本本地化条目")

    def extract_defined_text_blocks(self, content):
        """提取defined_text块，处理嵌套花括号"""
        lines = content.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith('defined_text'):
                # 找到defined_text块开始
                block_lines = []
                brace_count = 0
                in_block = False

                # 收集整个块的内容
                while i < len(lines):
                    current_line = lines[i]
                    block_lines.append(current_line)

                    # 统计花括号
                    for char in current_line:
                        if char == '{':
                            brace_count += 1
                            in_block = True
                        elif char == '}':
                            brace_count -= 1

                    i += 1

                    # 当所有花括号匹配完成时，块结束
                    if in_block and brace_count == 0:
                        break

                # 处理收集到的块
                block_content = '\n'.join(block_lines)
                self.parse_defined_text_block(block_content)
            else:
                i += 1

    def parse_defined_text_block(self, block_content):
        """解析单个defined_text块"""
        # 提取name
        name_match = re.search(r'name\s*=\s*(\w+)', block_content)
        if not name_match:
            return

        name = name_match.group(1)

        # 提取default值
        default_match = re.search(r'default\s*=\s*(\w+)', block_content)
        if default_match:
            self.scripted_loc_dict[name] = default_match.group(1)
        else:
            # 如果没有default，找第一个localization_key
            loc_key_match = re.search(r'localization_key\s*=\s*(\w+)', block_content)
            if loc_key_match:
                self.scripted_loc_dict[name] = loc_key_match.group(1)

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

    def process_dollar_replacement(self, text, max_depth=5, current_depth=0):
        """处理$xxx$替换符"""
        if current_depth >= max_depth:
            return text, True

        dollar_matches = re.findall(r'\$([^$]+)\$', text)
        if not dollar_matches:
            return text, False

        for match in dollar_matches:
            if match in self.localization_dict:
                replacement = self.localization_dict[match]
                replacement, exceeded = self.process_dollar_replacement(replacement, max_depth, current_depth + 1)
                text = text.replace(f'${match}$', replacement)

        if current_depth >= max_depth - 1 and re.search(r'\$[^$]+\$', text):
            return text, True

        return text, False

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符"""
        if not text or current_depth >= max_depth:
            return text

        def replace_brackets(match):
            inner_text = match.group(1)
            result = inner_text

            # 处理 [xxx.GetNamePlural] 或 [xxx.GetName]
            name_match = re.match(r'(\w+)\.GetName(Plural)?', inner_text)
            if name_match:
                prefix = name_match.group(1)
                is_plural = name_match.group(2) is not None

                job_key = f"job_{prefix}"
                if is_plural:
                    job_key_plural = f"{job_key}_plural"
                    if job_key_plural in self.localization_dict:
                        result = self.localization_dict[job_key_plural]
                    elif job_key in self.localization_dict:
                        result = self.localization_dict[job_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到工作本地化: {job_key} (来自 [{inner_text}])"
                        )
                else:
                    if job_key in self.localization_dict:
                        result = self.localization_dict[job_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到工作本地化: {job_key} (来自 [{inner_text}])"
                        )

            # 处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' not in inner_text:
                concept_key = inner_text[1:-1]
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            # 处理 ['xxx:yyy'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)
                if len(parts) == 2:
                    district_key = parts[1]
                    if district_key in self.localization_dict:
                        result = self.localization_dict[district_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到区划本地化: {district_key} (来自 [{inner_text}])"
                        )

            # 常规[]替换符处理
            elif inner_text in self.scripted_loc_dict:
                loc_key = self.scripted_loc_dict[inner_text]
                if loc_key in self.localization_dict:
                    result = self.localization_dict[loc_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"脚本本地化键 '{loc_key}' 未找到本地化文本 (来自 [{inner_text}])"
                    )
            else:
                self.debug_info['missing_bracket_replacement'].add(
                    f"脚本本地化 '{inner_text}' 未找到"
                )
                return match.group(0)

            if current_depth < max_depth - 1:
                result = self.process_bracket_replacement(result, max_depth, current_depth + 1)

            return result

        text = re.sub(r'\[([^\]]+)\]', replace_brackets, text)
        return text

    def process_pound_replacement(self, text):
        """处理£xxx£替换符"""
        if not text:
            return text

        def replace_pounds(match):
            inner_text = match.group(1)
            concept_key = f"concept_{inner_text}"

            if concept_key in self.localization_dict:
                return self.localization_dict[concept_key]
            else:
                if inner_text in self.localization_dict:
                    return self.localization_dict[inner_text]
                else:
                    self.debug_info['missing_pound_replacement'].add(
                        f"概念本地化 '{concept_key}' 和 '{inner_text}' 均未找到")
                    return ""

        text = re.sub(r'£([^£]+)£', replace_pounds, text)

        text, _ = self.process_dollar_replacement(text)
        text = self.process_bracket_replacement(text)

        return text

    def process_special_pound_replacement(self, text):
        """处理特殊的£xxx£替换符 - 固定翻译"""
        if not text:
            return text

        text = text.replace('£blocker£', '')

        # 固定£替换符翻译字典
        special_pound_translations = {
            'society': '社会学研究',
            'physics': '物理学研究',
            'engineering': '工程学研究',
            'trade_value': '贸易'
        }

        for pound_key, translation in special_pound_translations.items():
            text = text.replace(f'£{pound_key}£', translation)

        return text

    def process_district_name(self, text):
        """专门处理区划名称"""
        if not text:
            return text

        # 1. 预处理：处理$$替换符和清理颜色标识符
        text = self.preprocess_localization_text(text)

        # 2. 清理空白文本
        text = self.clean_whitespace_text(text)

        # 3. 处理[]替换符
        text = self.process_bracket_replacement(text)

        # 4. 处理特殊£xxx£替换（固定翻译）
        text = self.process_special_pound_replacement(text)

        # 5. 处理常规£xxx£替换
        text = self.process_pound_replacement(text)

        # 6. 处理换行符
        text = text.replace('\\n', '\n')

        # 7. 最终清理空白
        text = self.clean_whitespace_text(text)

        return text

    def parse_district_file(self, file_path):
        """解析单个区划文件"""
        file_name = os.path.basename(file_path)
        print(f"正在解析文件: {file_name}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 使用更强大的方法提取完整的区划定义
            districts = self.extract_complete_blocks(content)

            for district_key, district_content in districts:
                print(f"  解析区划: {district_key}")
                self.parse_single_district(district_key, district_content, file_name)

        except Exception as e:
            print(f"解析文件失败: {file_path}, 错误: {e}")

    def extract_complete_blocks(self, content):
        """提取完整的区划定义块，正确处理嵌套花括号"""
        blocks = []
        pos = 0

        while pos < len(content):
            # 查找任何键后跟等号和花括号的模式
            pattern = r'(\w+)\s*=\s*\{'
            match = re.search(pattern, content[pos:])

            if not match:
                break

            key = match.group(1)
            brace_start = pos + match.end() - 1
            block_content, new_pos = self.extract_balanced_braces(content, brace_start)

            if block_content:
                # 提取内容部分（去掉外层花括号）
                content_start = block_content.find('{') + 1
                content_end = block_content.rfind('}')
                block_content_clean = block_content[content_start:content_end].strip()

                blocks.append((key, block_content_clean))

            pos = new_pos

        return blocks

    def parse_single_district(self, district_key, district_content, source_file):
        """解析单个区划定义"""
        # 只处理以district_开头的键
        if not district_key.startswith('district_'):
            return

        district_data = {
            'key': district_key,
            'name': '',
            'zone_slots': [],
            'planet_types': [],
            'source': source_file,
        }

        # 获取区划名称
        raw_name = self.localization_dict.get(district_key, district_key)
        district_data['name'] = self.process_district_name(raw_name)

        if district_key not in self.localization_dict:
            self.debug_info['missing_district_name'].add(district_key)

        # 提取zone_slots
        zone_slots_match = re.search(r'zone_slots\s*=\s*\{', district_content)
        if zone_slots_match:
            brace_start = zone_slots_match.end() - 1
            zone_slots_block, _ = self.extract_balanced_braces(district_content, brace_start)

            if zone_slots_block:
                # 提取zone_slots中的所有值
                slots_matches = re.findall(r'\b(\w+)\b', zone_slots_block)
                district_data['zone_slots'] = [slot_name for slot_name in slots_matches
                                               if slot_name not in ['zone_slots', '=', '{', '}']]

        # 提取允许的星球类型（从potential块中提取uses_district_set）
        potential_match = re.search(r'potential\s*=\s*\{', district_content)
        if potential_match:
            brace_start = potential_match.end() - 1
            potential_block, _ = self.extract_balanced_braces(district_content, brace_start)

            if potential_block:
                # 提取所有uses_district_set后面的值
                district_set_matches = re.findall(r'uses_district_set\s*=\s*(\w+)', potential_block)
                district_data['planet_types'] = list(set(district_set_matches))  # 去重

        print(
            f"    区划 {district_key}: 特化类型={district_data['zone_slots']}, 星球类型={district_data['planet_types']}")
        self.districts.append(district_data)

    def parse_all_districts(self):
        """解析所有区划文件"""
        print("正在解析区划文件...")

        if not os.path.exists(DISTRICTS_DIR):
            print(f"区划目录不存在: {DISTRICTS_DIR}")
            return

        for file_name in os.listdir(DISTRICTS_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(DISTRICTS_DIR, file_name)
                self.parse_district_file(file_path)

        print(f"已解析 {len(self.districts)} 个区划")

    def generate_excel(self):
        """生成Excel文件"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}区划信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('区划信息')

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

        # 设置列宽和行高
        worksheet.set_column('A:A', 40)  # 区划名称
        worksheet.set_column('B:B', 40)  # 区划KEY
        worksheet.set_column('C:C', 30)  # 允许的特化类型
        worksheet.set_column('D:D', 30)  # 允许的星球类型
        worksheet.set_column('E:E', 20)  # 来源文件
        worksheet.set_default_row(25)  # 设置默认行高

        # 写入表头
        headers = ['区划名称', '区划KEY', '允许的特化类型', '允许的星球类型', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for district in self.districts:
            worksheet.write(row, 0, district['name'], cell_format)
            worksheet.write(row, 1, district['key'], cell_format)

            # 特化类型用逗号分隔
            zone_slots_text = ', '.join(district['zone_slots'])
            worksheet.write(row, 2, zone_slots_text, cell_format)

            # 星球类型用逗号分隔
            planet_types_text = ', '.join(district['planet_types'])
            worksheet.write(row, 3, planet_types_text, cell_format)

            worksheet.write(row, 4, district['source'], cell_format)
            row += 1

        workbook.close()

        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}区划信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 区划信息调试信息 ===\n\n")

            if self.debug_info['missing_district_name']:
                f.write("=== 缺失区划名称 ===\n")
                for info in sorted(self.debug_info['missing_district_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_bracket_replacement']:
                f.write("=== 缺失[]替换符 ===\n")
                for info in sorted(self.debug_info['missing_bracket_replacement']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_pound_replacement']:
                f.write("=== 缺失£xxx£替换符 ===\n")
                for info in sorted(self.debug_info['missing_pound_replacement']):
                    f.write(f"{info}\n")
                f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取区划信息...")

        self.load_localization_files()
        self.load_scripted_loc()
        self.parse_all_districts()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = DistrictExtractor()
    extractor.run()