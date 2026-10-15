from A配置路径 import *

import os
import re
import xlsxwriter


class ArchaeologicalSiteExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_variables_dict = {}
        self.scripted_loc_dict = {}
        self.archaeological_sites = []
        self.insensitive_localization_dict = {}

        # 调试信息
        self.debug_info = {
            'missing_site_name': set(),
            'missing_event_name': set(),
            'missing_scripted_variable': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
            'max_recursion_depth_exceeded': [],
            'debug_sites': {}
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    def debug_log(self, site_key, message):
        """记录特定考古遗址的调试信息"""
        if site_key in ['zroni_digsite_1', 'zroni_digsite_2', 'baol_digsite_1']:
            if site_key not in self.debug_info['debug_sites']:
                self.debug_info['debug_sites'][site_key] = []
            self.debug_info['debug_sites'][site_key].append(message)

    def preprocess_localization_text(self, text):
        """预处理本地化文本"""
        if not text:
            return text

        # 1. 递归处理$$替换符
        text, _ = self.process_dollar_replacement(text)

        # 2. 清除颜色标识符 §Yxxx§!
        text = re.sub(r'§[A-Z](.*?)§!', r'\1', text)

        return text

    def remove_duplicate_text(self, text):
        """清除连续重复文本，只处理文字，不处理数字"""
        if not text:
            return text

        # 清除所有空白字符（包括空格、制表符、换行符等）
        text = re.sub(r'\s+', '', text)

        # 查找连续重复的文字（2-10个字符），不处理单个字的重复
        pattern = r'([\u4e00-\u9fa5a-zA-Z]{2,10})\1+'

        # 循环处理直到没有重复
        changed = True
        while changed:
            changed = False
            matches = re.findall(pattern, text)
            if matches:
                for match in matches:
                    # 只保留一个副本
                    text = text.replace(match + match, match)
                    changed = True

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

    def process_special_at_replacement(self, text):
        """处理特殊@替换符结构：@[ xxx * 数值]"""
        if not text:
            return text

        # 查找 @[ xxx * 数值] 格式
        pattern = r'@\[\s*(\w+)\s*\*\s*([\d.-]+)\s*\]'

        def replace_special_at(match):
            var_name = match.group(1)
            multiplier = float(match.group(2))

            if var_name in self.scripted_variables_dict:
                result = self.scripted_variables_dict[var_name] * multiplier

                # 处理浮点数精度问题 - 清除浮点数，如果是整数则返回整数
                if abs(result - round(result)) < 1e-10:
                    result = round(result)

                # 如果是整数，返回整数形式
                if result == int(result):
                    return str(int(result))
                else:
                    return str(result)
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                return match.group(0)

        text = re.sub(pattern, replace_special_at, text)

        return text

    def process_special_dollar_at_replacement(self, text):
        """处理特殊替换符：$@xxx|任意数值%$"""
        if not text:
            return text

        # 查找 $@xxx|任意数值%$ 格式
        pattern = r'\$@(\w+)\|([^$]+)%?\$'

        def replace_special_dollar_at(match):
            var_name = match.group(1)
            default_value = match.group(2)
            has_percent = match.group(0).endswith('%$')

            if var_name in self.scripted_variables_dict:
                result = self.scripted_variables_dict[var_name]
                # 清除浮点数，如果是整数则返回整数
                if result == int(result):
                    result = int(result)

                if has_percent:
                    result = result * 100
                    if abs(result - round(result)) < 0.001:
                        return f"{int(result)}%"
                    else:
                        formatted = f"{result:.2f}%".rstrip('0').rstrip('.')
                        if formatted.endswith('.'):
                            formatted = formatted[:-1]
                        if not formatted.endswith('%'):
                            formatted += '%'
                        return formatted
                else:
                    return str(result)
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                if has_percent:
                    return f"{default_value}%"
                else:
                    return default_value

        text = re.sub(pattern, replace_special_dollar_at, text)

        return text

    def load_localization_files(self):
        """加载所有本地化文件"""
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
                pass

        # 构建大小写不敏感的本地化字典
        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localization_dict.items()}
        print(f"已加载 {len(self.localization_dict)} 个本地化条目")

    def load_scripted_variables(self):
        """加载脚本变量"""

        if os.path.exists(SCRIPTED_VARIABLES_DIR):
            for file_name in os.listdir(SCRIPTED_VARIABLES_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(SCRIPTED_VARIABLES_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                        matches = re.findall(pattern, content)

                        for var_name, var_value in matches:
                            # 清除浮点数，如果是整数则存储为整数
                            value = float(var_value)
                            if value == int(value):
                                self.scripted_variables_dict[var_name] = int(value)
                            else:
                                self.scripted_variables_dict[var_name] = value

                    except Exception as e:
                        pass

        if os.path.exists(ARCHAEOLOGICAL_SITES_DIR):
            for file_name in os.listdir(ARCHAEOLOGICAL_SITES_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(ARCHAEOLOGICAL_SITES_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        lines = content.split('\n')
                        brace_level = 0

                        for line in lines:
                            line = line.strip()
                            if not line or line.startswith('#'):
                                continue

                            for char in line:
                                if char == '{':
                                    brace_level += 1
                                elif char == '}':
                                    brace_level -= 1

                            if brace_level == 0:
                                pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                                match = re.match(pattern, line)
                                if match:
                                    var_name, var_value = match.groups()
                                    # 清除浮点数，如果是整数则存储为整数
                                    value = float(var_value)
                                    if value == int(value):
                                        self.scripted_variables_dict[var_name] = int(value)
                                    else:
                                        self.scripted_variables_dict[var_name] = value

                    except Exception as e:
                        pass

        print(f"已加载 {len(self.scripted_variables_dict)} 个脚本变量")

    def load_scripted_loc(self):
        """加载脚本本地化"""

        if not os.path.exists(SCRIPTED_LOC_DIR):
            return

        for file_name in os.listdir(SCRIPTED_LOC_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(SCRIPTED_LOC_DIR, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    self.extract_defined_text_blocks(content)

                except Exception as e:
                    pass

        print(f"已加载 {len(self.scripted_loc_dict)} 个脚本本地化条目")

    def extract_defined_text_blocks(self, content):
        """提取defined_text块，处理嵌套花括号"""
        lines = content.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith('defined_text'):
                block_lines = []
                brace_count = 0
                in_block = False

                while i < len(lines):
                    current_line = lines[i]
                    block_lines.append(current_line)

                    for char in current_line:
                        if char == '{':
                            brace_count += 1
                            in_block = True
                        elif char == '}':
                            brace_count -= 1

                    i += 1

                    if in_block and brace_count == 0:
                        break

                block_content = '\n'.join(block_lines)
                self.parse_defined_text_block(block_content)
            else:
                i += 1

    def parse_defined_text_block(self, block_content):
        """解析单个defined_text块"""
        name_match = re.search(r'name\s*=\s*(\w+)', block_content)
        if not name_match:
            return

        name = name_match.group(1)

        default_match = re.search(r'default\s*=\s*(\w+)', block_content)
        if default_match:
            self.scripted_loc_dict[name] = default_match.group(1)
        else:
            loc_key_match = re.search(r'localization_key\s*=\s*(\w+)', block_content)
            if loc_key_match:
                self.scripted_loc_dict[name] = loc_key_match.group(1)

    def extract_modifier_value(self, value_str):
        """提取修正值，处理@变量"""
        if isinstance(value_str, (int, float)):
            # 清除浮点数，如果是整数则返回整数
            if value_str == int(value_str):
                return int(value_str)
            return value_str

        if isinstance(value_str, str) and '#' in value_str:
            value_str = value_str.split('#')[0].strip()

        value_str = self.process_special_at_replacement(value_str)
        value_str = self.process_special_dollar_at_replacement(value_str)

        if isinstance(value_str, str) and value_str.startswith('@'):
            var_name = value_str[1:]
            if var_name in self.scripted_variables_dict:
                result = self.scripted_variables_dict[var_name]
                # 清除浮点数，如果是整数则返回整数
                if result == int(result):
                    return int(result)
                return result
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                return value_str

        if isinstance(value_str, str) and '$' in value_str:
            value_str, _ = self.process_dollar_replacement(value_str)

        # 尝试转换为数字并清除浮点数
        try:
            value_float = float(value_str)
            if value_float == int(value_float):
                return int(value_float)
            return value_float
        except (ValueError, TypeError):
            return value_str

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
            else:
                pass

        if current_depth >= max_depth - 1 and re.search(r'\$[^$]+\$', text):
            return text, True

        return text, False

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符"""
        if not text or current_depth >= max_depth:
            return text

        text = re.sub(r'\[[^\]]*\.GetIcon\]', '', text)

        def replace_brackets(match):
            inner_text = match.group(1)
            result = inner_text

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

            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' not in inner_text:
                concept_key = inner_text[1:-1]
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localization_dict:
                        result = self.localization_dict[building_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到建筑本地化: {building_key} (来自 [{inner_text}])"
                        )

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

    def process_site_name(self, text):
        """处理考古遗址名称文本"""
        if not text:
            return text

        # 1. 预处理：处理$$替换符和清理颜色标识符
        text = self.preprocess_localization_text(text)

        # 2. 清理空白文本
        text = self.clean_whitespace_text(text)

        # 3. 处理特殊$@xxx|任意数值%$替换符
        text = self.process_special_dollar_at_replacement(text)

        # 4. 处理[]替换符
        text = self.process_bracket_replacement(text)

        # 5. 处理常规£xxx£替换
        text = self.process_pound_replacement(text)

        # 6. 清除重复文本
        text = self.remove_duplicate_text(text)

        # 7. 处理换行符
        text = text.replace('\\n', '\n')

        # 8. 最终清理空白
        text = self.clean_whitespace_text(text)

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

    def parse_archaeological_site_file(self, file_path):
        """解析单个考古遗址文件"""
        file_name = os.path.basename(file_path)

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 提取完整的考古遗址定义块
            sites = self.extract_complete_blocks(content)

            for site_key, site_content in sites:
                self.parse_single_site(site_key, site_content, file_name)

        except Exception as e:
            pass

    def extract_complete_blocks(self, content):
        """提取完整的考古遗址定义块，正确处理嵌套花括号"""
        blocks = []
        lines = content.split('\n')

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 寻找考古遗址定义开始
            if line and '=' in line and '{' in line:
                parts = line.split('=', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    if '{' in parts[1]:
                        block_lines = []
                        brace_count = 0
                        in_block = False

                        j = i
                        while j < len(lines):
                            current_line = lines[j]
                            block_lines.append(current_line)

                            for char in current_line:
                                if char == '{':
                                    brace_count += 1
                                    in_block = True
                                elif char == '}':
                                    brace_count -= 1

                            j += 1

                            if in_block and brace_count == 0:
                                break

                        if in_block and brace_count == 0:
                            block_content = '\n'.join(block_lines)
                            match = re.match(r'(\w+)\s*=\s*\{', block_content)
                            if match:
                                block_key = match.group(1)
                                content_start = block_content.find('{') + 1
                                content_end = block_content.rfind('}')
                                block_content_clean = block_content[content_start:content_end].strip()

                                blocks.append((block_key, block_content_clean))

                            i = j
                            continue
            i += 1

        return blocks

    def parse_single_site(self, site_key, site_content, source_file):
        """解析单个考古遗址定义"""
        site_data = {
            'key': site_key,
            'name': '',
            'details': '',
            'source': source_file
        }

        # 考古遗址名称直接使用site_key查找本地化
        raw_name = self.localization_dict.get(site_key, site_key)
        site_data['name'] = self.process_site_name(raw_name)

        if site_key not in self.localization_dict:
            self.debug_info['missing_site_name'].add(site_key)

        # 提取所有stage信息
        stages = self.extract_stages(site_content, site_key)

        # 构建考古详情字符串
        details_parts = []
        for i, stage in enumerate(stages, 1):
            event_key = stage['event']
            difficulty = stage['difficulty']

            # 获取事件名称的本地化（event + ".name"）
            event_name_key = f"{event_key}.name"
            event_name = self.localization_dict.get(event_name_key, event_key)

            # 处理事件名称中的替换符
            event_name = self.process_site_name(event_name)

            details_parts.append(f"阶段{i}：{event_name}（难度{difficulty}）")

            if event_name_key not in self.localization_dict:
                self.debug_info['missing_event_name'].add(event_name_key)

        site_data['details'] = '\n'.join(details_parts)

        self.archaeological_sites.append(site_data)

    def extract_stages(self, site_content, site_key):
        """提取所有stage信息"""
        stages = []
        pos = 0

        while pos < len(site_content):
            # 查找stage块
            stage_match = re.search(r'stage\s*=\s*\{', site_content[pos:])
            if not stage_match:
                break

            brace_start = pos + stage_match.end() - 1
            stage_block, new_pos = self.extract_balanced_braces(site_content, brace_start)

            if stage_block:
                # 提取difficulty和event
                difficulty_match = re.search(r'difficulty\s*=\s*(\d+)', stage_block)
                event_match = re.search(r'event\s*=\s*([^\s#]+)', stage_block)

                if difficulty_match and event_match:
                    difficulty = difficulty_match.group(1)
                    event = event_match.group(1)

                    stages.append({
                        'difficulty': difficulty,
                        'event': event
                    })

                    self.debug_log(site_key, f"找到阶段: 难度={difficulty}, 事件={event}")

            pos = new_pos

        return stages

    def parse_all_archaeological_sites(self):
        """解析所有考古遗址文件"""

        if not os.path.exists(ARCHAEOLOGICAL_SITES_DIR):
            return

        for file_name in os.listdir(ARCHAEOLOGICAL_SITES_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(ARCHAEOLOGICAL_SITES_DIR, file_name)
                self.parse_archaeological_site_file(file_path)

        print(f"已解析 {len(self.archaeological_sites)} 个考古遗址")

    def generate_excel(self):
        """生成Excel文件"""

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}考古遗址信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('考古遗址信息')

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

        description_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'text_wrap': True
        })

        # 设置列宽
        worksheet.set_column('A:A', 50)  # 考古名称
        worksheet.set_column('B:B', 40)  # 考古KEY
        worksheet.set_column('C:C', 80)  # 考古详情
        worksheet.set_column('D:D', 30)  # 来源文件

        # 写入表头
        headers = ['考古名称', '考古KEY', '考古详情', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for site in self.archaeological_sites:
            worksheet.write(row, 0, site['name'], description_format)
            worksheet.write(row, 1, site['key'], cell_format)
            worksheet.write(row, 2, site['details'], description_format)
            worksheet.write(row, 3, site['source'], cell_format)
            row += 1

        workbook.close()
        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息"""

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}考古遗址调试信息.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 考古遗址调试信息 ===\n\n")

            if self.debug_info['missing_site_name']:
                f.write("=== 缺失考古遗址名称 ===\n")
                for info in sorted(self.debug_info['missing_site_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_event_name']:
                f.write("=== 缺失事件名称 ===\n")
                for info in sorted(self.debug_info['missing_event_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_scripted_variable']:
                f.write("=== 缺失脚本变量 ===\n")
                for info in sorted(self.debug_info['missing_scripted_variable']):
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

            if self.debug_info['max_recursion_depth_exceeded']:
                f.write("=== 超过最大递归深度 ===\n")
                for info in self.debug_info['max_recursion_depth_exceeded']:
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['debug_sites']:
                f.write("=== 特定考古遗址调试信息 ===\n\n")
                for site_key, messages in self.debug_info['debug_sites'].items():
                    f.write(f"考古遗址 {site_key}:\n")
                    for message in messages:
                        f.write(f"  - {message}\n")
                    f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""

        self.load_localization_files()
        self.load_scripted_variables()
        self.load_scripted_loc()
        self.parse_all_archaeological_sites()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = ArchaeologicalSiteExtractor()
    extractor.run()