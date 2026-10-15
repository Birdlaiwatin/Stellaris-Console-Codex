import os
import re
from pathlib import Path
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 修正翻译字典 import MODIFIER_TRANSLATIONS
from 颜色转换字典 import SPECIAL_SUFFIXES

# 特殊标记翻译字典
SPECIAL_MARKER_TRANSLATIONS = {
    "This.GetComposerColor": "生命之织缕",
    "This.GetCradleColor": "魂灵之月窠",
    "This.GetEaterColor": "世界之喰煞",
    "This.GetInstrumentColor": "渴求之器具"
}

# 颜色代码映射字典
COLOR_MAP = {
    '§Y': 'orange',
    '§G': 'green',
    '§R': 'red',
    '§B': 'blue',
    '§W': 'white',
    '§C': 'cyan',
    '§E': 'cyan',
    '§H': 'orange',
    '§I': 'yellow',
    '§K': 'orange',
    '§L': 'orange',
    '§M': 'purple',
    '§P': 'red',
    '§S': 'orange',
    '§T': 'gray',
    '§U': 'blue',
    '§V': 'green',
    '§A': 'blue_underline',
}

# 条件类型翻译字典
CONDITION_TYPE_TRANSLATIONS = {
    "ethics": "思潮",
    "civics": "国民理念",
    "authority": "政体",
    "origin": "起源",
    "species_class": "物种",
    "species_archetype": "物种类型",
    "graphical_culture": ""
}

# 存储未替换的标记
unresolved_markers = set()


class CivicOriginExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_vars = {}
        self.scripted_loc_dict = {}
        self.insensitive_localization_dict = {}

        # 存储数据
        self.civics = []
        self.origins = []
        self.councilors_dict = {}

        # 调试信息
        self.debug_info = {
            'missing_civic_name': set(),
            'missing_origin_name': set(),
            'missing_tooltip': set(),
            'missing_scripted_variable': set(),
            'missing_images': [],
            'unresolved_markers': set(),
            'missing_modifier_translations': set()  # 新增：记录未找到的修正翻译
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(CIVIC_IMAGES_DIR, exist_ok=True)
        os.makedirs(ORIGIN_IMAGES_DIR, exist_ok=True)

    def convert_dds_to_png(self, dds_path, png_path):
        """
        将DDS文件转换为PNG格式。
        """
        try:
            # 检查源文件是否存在
            if not os.path.exists(dds_path):
                print(f"警告: DDS文件不存在，无法转换: {dds_path}")
                return False

            # 如果目标PNG目录不存在，则创建
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            # 使用imageio读取DDS并写入PNG
            img = imageio.imread(dds_path)
            imageio.imwrite(png_path, img)
            print(f"已转换: {os.path.basename(dds_path)} -> {os.path.basename(png_path)}")
            return True
        except Exception as e:
            print(f"转换DDS到PNG时出错 ({dds_path}): {e}")
            return False

    def clean_localisation_value(self, value):
        """第一阶段清理：处理注释、多余引号、空格、简单标记，但保留\n"""
        if not value:
            return value

        # 移除注释
        if '#' in value:
            value = value.split('#')[0].strip()

        # 移除多余引号
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]

        # 移除所有空格（包括全角空格）
        value = re.sub(r'[\s　]+', '', value)

        # 移除简单标记：$t$, $tt$, $ttt$, $tttt$
        value = re.sub(r'\$t\$$', '', value)
        value = re.sub(r'\$tt\$$', '', value)
        value = re.sub(r'\$ttt\$$', '', value)
        value = re.sub(r'\$tttt\$$', '', value)

        # 将$TABBED_NEWLINE$替换为换行符\n
        value = value.replace('$TABBED_NEWLINE$', '\\n')

        return value.strip()

    def compress_multiple_newlines(self, text):
        """压缩多个连续换行符为单个换行符"""
        if not text:
            return text

        # 将多个连续的\n替换为单个\n
        compressed_text = re.sub(r'\\n+', '\\n', text)
        return compressed_text

    def is_inside_color_tag(self, text, position):
        """检查指定位置是否在颜色标记内"""
        text_before = text[:position]
        # 计算未闭合的颜色标记
        open_color_tags = re.findall(r'§[A-Z]', text_before)
        close_color_tags = re.findall(r'§!', text_before)
        # 如果开颜色标记数量大于闭颜色标记数量，说明在颜色标记内
        return len(open_color_tags) > len(close_color_tags)

    def preprocess_text_simple(self, text, avoid_orange=False):
        """简化的第二阶段预处理：处理$@和$$替换符，智能添加颜色标记"""
        if not text:
            return text

        original_text = text

        # 处理$@替换符 - 不标记颜色
        if self.scripted_vars:
            at_pattern = r'\$@([^|$]+)(?:\|([^$]+))?\$'

            def replace_at_var(match):
                var_name = f"@{match.group(1)}"
                suffix = match.group(2)

                if var_name in self.scripted_vars:
                    value = str(self.scripted_vars[var_name])  # 确保值为字符串

                    # 处理百分比格式
                    if suffix and '%' in suffix:
                        try:
                            num_value = float(value)
                            percentage = num_value * 100
                            if percentage == int(percentage):
                                return f"{int(percentage)}%"
                            else:
                                return f"{percentage:.1f}%"
                        except (ValueError, TypeError):
                            return value
                    else:
                        return value
                else:
                    self.debug_info['missing_scripted_variable'].add(var_name)
                    return match.group(0)

            text = re.sub(at_pattern, replace_at_var, text)

        # 处理$$替换符 - 智能添加颜色标记
        if self.localization_dict:
            max_iterations = 10
            iteration = 0
            changed = True

            while changed and iteration < max_iterations:
                changed = False
                iteration += 1

                # 查找所有$$替换符
                pattern = r'\$([^$]+)\$'
                matches = list(re.finditer(pattern, text))

                if not matches:
                    break

                # 从后往前替换，避免位置变化影响
                for match in reversed(matches):
                    full_match = match.group(0)
                    key = match.group(1)

                    if key in self.localization_dict:
                        replacement = self.localization_dict[key]

                        # 检查是否会产生自引用循环
                        if full_match not in replacement:

                            # 检查这个替换符是否在颜色标记内
                            is_in_color = self.is_inside_color_tag(text, match.start())

                            # 检查替换后的文本是否还包含$$替换符
                            has_more_dollar_signs = bool(re.search(pattern, replacement))

                            # 检查替换文本是否仅为换行符
                            is_newline_only = replacement == '\\n'

                            # 新规则：如果$$替换符中的内容全为大写字母，不添加颜色
                            is_all_uppercase = key.isupper()

                            # 如果avoid_orange为True，则不添加橙色标记
                            if not avoid_orange and not is_in_color and not has_more_dollar_signs and not is_newline_only and not is_all_uppercase:
                                replacement = f"§H{replacement}§!"

                            # 执行替换
                            start_pos = match.start()
                            end_pos = match.end()
                            text = text[:start_pos] + replacement + text[end_pos:]
                            changed = True

                # 检查是否还有更多$$替换符
                if not re.search(pattern, text):
                    break

            if iteration >= max_iterations:
                print(f"警告: $$替换达到最大迭代次数 {max_iterations}")

        return text

    def preprocess_localisation_dict_simple(self):
        """简化的预处理整个本地化字典"""
        preprocessed_dict = {}

        # 先进行第一阶段清理
        for key, value in self.localization_dict.items():
            cleaned_value = self.clean_localisation_value(value)
            preprocessed_dict[key] = cleaned_value

        # 直接进行第二阶段预处理
        max_iterations = 5
        iteration = 0
        changed = True

        while changed and iteration < max_iterations:
            changed = False
            iteration += 1

            temp_dict = preprocessed_dict.copy()

            for key, value in preprocessed_dict.items():
                original_value = value
                preprocessed_value = self.preprocess_text_simple(value)
                preprocessed_value = self.compress_multiple_newlines(preprocessed_value)

                if preprocessed_value != original_value:
                    changed = True
                    temp_dict[key] = preprocessed_value

            preprocessed_dict = temp_dict

            if not changed:
                break

        if iteration >= max_iterations:
            print(f"警告: 预处理达到最大迭代次数 {max_iterations}")

        return preprocessed_dict

    def parse_localisation_files(self):
        """解析本地化文件，构建翻译字典"""
        print("正在解析本地化文件...")
        for file_path in Path(LOCALISATION_DIR).rglob("*.yml"):
            try:
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    for line in f:
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            parts = line.split(':', 1)
                            if len(parts) == 2:
                                key = parts[0].strip()
                                value = parts[1].strip()
                                value = self.clean_localisation_value(value)
                                self.localization_dict[key] = value
            except Exception as e:
                print(f"解析本地化文件时出错 {file_path}: {e}")

        print(f"已加载 {len(self.localization_dict)} 个本地化条目")

    def parse_scripted_variables(self):
        """解析脚本变量文件，构建变量字典"""
        print("正在解析脚本变量文件...")
        scripted_vars = {}

        for file_path in Path(SCRIPTED_VARIABLES_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                    pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                    matches = re.findall(pattern, content)

                    for var_name, var_value in matches:
                        scripted_vars[f"@{var_name}"] = float(var_value)

            except Exception as e:
                print(f"解析脚本变量文件时出错 {file_path}: {e}")

        self.scripted_vars = scripted_vars
        print(f"已加载 {len(self.scripted_vars)} 个脚本变量")

    def parse_scripted_loc_files(self):
        """解析scripted_loc文件，构建替换符字典"""
        print("正在解析scripted_loc文件...")
        scripted_loc_dict = {}

        for file_path in Path(SCRIPTED_LOC_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()

                    i = 0
                    while i < len(lines):
                        line = lines[i].strip()

                        if line.startswith('defined_text = {'):
                            name = None
                            default = None
                            first_localization_key = None

                            brace_count = 1
                            i += 1

                            while i < len(lines) and brace_count > 0:
                                line = lines[i].strip()

                                if not line or line.startswith('#'):
                                    i += 1
                                    continue

                                brace_count += line.count('{')
                                brace_count -= line.count('}')

                                if name is None and 'name =' in line:
                                    name_match = re.search(r'name\s*=\s*(\w+)', line)
                                    if name_match:
                                        name = name_match.group(1)

                                if default is None and 'default =' in line:
                                    default_match = re.search(r'default\s*=\s*(\w+)', line)
                                    if default_match:
                                        default = default_match.group(1)

                                if first_localization_key is None and 'localization_key =' in line:
                                    loc_match = re.search(r'localization_key\s*=\s*(\w+)', line)
                                    if loc_match:
                                        first_localization_key = loc_match.group(1)

                                i += 1

                            if name:
                                if default:
                                    scripted_loc_dict[name] = default
                                elif first_localization_key:
                                    scripted_loc_dict[name] = first_localization_key

                        i += 1

            except Exception as e:
                print(f"解析scripted_loc文件时出错 {file_path}: {e}")

        self.scripted_loc_dict = scripted_loc_dict
        print(f"已加载 {len(self.scripted_loc_dict)} 个scripted_loc条目")

    def resolve_localisation(self, value, depth=0, max_depth=10):
        """递归解析本地化值中的变量引用"""
        if depth > max_depth:
            return value

        pattern = r'\$([^$]+)\$'
        matches = re.findall(pattern, value)
        if not matches:
            return value
        for match in matches:
            resolved_value = self.resolve_localisation(self.localization_dict.get(match, match), depth + 1, max_depth)
            value = value.replace(f'${match}$', resolved_value)
        return value

    def find_modifier_translation(self, modifier_key, max_depth=8):
        """查找修正字段的本地化翻译"""
        # 优先从内置字典查找
        if modifier_key in MODIFIER_TRANSLATIONS:
            translation = MODIFIER_TRANSLATIONS[modifier_key]
            return self.resolve_modifier_placeholders(translation, max_depth), None

        # 构建不区分大小写的本地化字典（如果尚未构建）
        if not hasattr(self, 'insensitive_localization_dict') or not self.insensitive_localization_dict:
            self.insensitive_localization_dict = {k.lower(): v for k, v in self.localization_dict.items()}

        search_key = modifier_key.lower()
        search_patterns = [
            f"mod_{search_key}",
            f"mod_{modifier_key.upper()}",
            f"mod_country_{search_key}",
            f"mod_planet_{search_key}",
            f"mod_federation_{search_key}",
            f"mod_leader_{search_key}",
            search_key
        ]

        # 在不区分大小写的字典中查找
        for pattern in search_patterns:
            if pattern in self.insensitive_localization_dict:
                translation = self.insensitive_localization_dict[pattern]
                return self.resolve_modifier_placeholders(translation, max_depth), None

        # 记录缺失的修正翻译
        self.debug_info['missing_modifier_translations'].add(modifier_key)
        return None, modifier_key

    def resolve_modifier_placeholders(self, text, max_depth=0):
        """递归解析修正文本中的占位符"""
        if max_depth <= 0:
            return text

        pound_pattern = r'£([^£]+)£'
        pound_matches = re.findall(pound_pattern, text)
        for match in pound_matches:
            concept_key = f"concept_{match}"
            concept_value = self.localization_dict.get(concept_key, "")
            if concept_value:
                resolved_concept = self.resolve_modifier_placeholders(concept_value, max_depth - 1)
                text = text.replace(f'£{match}£', resolved_concept)
            else:
                # 如果找不到对应本地化，直接删除£xxx£标记
                text = text.replace(f'£{match}£', '')

        dollar_pattern = r'\$([^$]+)\$'
        dollar_matches = re.findall(dollar_pattern, text)
        for match in dollar_matches:
            resolved_value = self.localization_dict.get(match, match)
            resolved_value = self.resolve_modifier_placeholders(resolved_value, max_depth - 1)
            text = text.replace(f'${match}$', resolved_value)

        return text

    def apply_special_replacements_with_force_blue(self, text, processed_markers=None, depth=0, force_blue=True):
        """特殊替换规则 - 可选择是否对[]标记和$$替换符使用蓝色下划线"""
        if depth > 10:
            return text

        if processed_markers is None:
            processed_markers = set()

        result = []
        i = 0
        text_len = len(text)

        while i < text_len:
            if text[i] == '[':
                mark_start = i
                brace_count = 1
                j = i + 1

                while j < text_len and brace_count > 0:
                    if text[j] == '[':
                        brace_count += 1
                    elif text[j] == ']':
                        brace_count -= 1
                    j += 1

                if brace_count == 0:
                    mark_end = j
                    full_mark = text[mark_start:mark_end]
                    content = full_mark[1:-1]

                    # 首先检查特殊标记翻译字典
                    if content in SPECIAL_MARKER_TRANSLATIONS:
                        replacement = SPECIAL_MARKER_TRANSLATIONS[content]
                        is_resolved = True
                    else:
                        replacement = ""
                        is_resolved = False

                        if '.GetIcon' in content:
                            replacement = ""
                            is_resolved = True

                        elif '.GetNamePlural' in content:
                            prefix = content.replace('.GetNamePlural', '')
                            job_key = f"job_{prefix}"
                            replacement = self.localization_dict.get(job_key, prefix)
                            is_resolved = True

                            new_text = self.apply_special_replacements_with_force_blue(
                                replacement, processed_markers, depth + 1, force_blue)
                            replacement = new_text

                        elif '.GetName' in content:
                            prefix = content.replace('.GetName', '')
                            job_key = f"job_{prefix}"
                            replacement = self.localization_dict.get(job_key, prefix)
                            is_resolved = True

                            new_text = self.apply_special_replacements_with_force_blue(
                                replacement, processed_markers, depth + 1, force_blue)
                            replacement = new_text

                        elif "'" in content:
                            quote_content = content
                            if content.startswith("'") and content.endswith("'"):
                                quote_content = content[1:-1]

                            if ',' in quote_content:
                                parts = quote_content.split(',', 1)
                                key = parts[1].strip()
                                replacement = key
                                is_resolved = True

                                new_text = self.apply_special_replacements_with_force_blue(
                                    replacement, processed_markers, depth + 1, force_blue)
                                replacement = new_text

                            elif ':' in quote_content:
                                key = quote_content.split(':')[-1].strip()
                                replacement = self.localization_dict.get(key, key)
                                is_resolved = True

                            else:
                                key = quote_content
                                if key in self.scripted_loc_dict:
                                    replacement_key = self.scripted_loc_dict[key]
                                    replacement = self.localization_dict.get(replacement_key, replacement_key)
                                    is_resolved = True
                                else:
                                    replacement = self.localization_dict.get(key, key)
                                    is_resolved = (replacement != key)

                                if replacement != key:
                                    new_text = self.apply_special_replacements_with_force_blue(
                                        replacement, processed_markers, depth + 1, force_blue)
                                    replacement = new_text

                        else:
                            if content in self.scripted_loc_dict:
                                replacement_key = self.scripted_loc_dict[content]
                                replacement = self.localization_dict.get(replacement_key, replacement_key)
                                is_resolved = True
                            else:
                                replacement = self.localization_dict.get(content, content)
                                is_resolved = (replacement != content)

                            if replacement != content:
                                new_text = self.apply_special_replacements_with_force_blue(
                                    replacement, processed_markers, depth + 1, force_blue)
                                replacement = new_text

                    has_more_brackets = '[' in replacement and ']' in replacement
                    is_newline_only = replacement == '\\n'

                    if is_resolved and not has_more_brackets and not is_newline_only and force_blue:
                        replacement_without_color = re.sub(r'§[A-Z]|§!', '', replacement)
                        replacement = f"§A{replacement_without_color}§!"

                    if not is_resolved:
                        self.debug_info['unresolved_markers'].add(content)

                    result.append(replacement)
                    i = mark_end
                    continue

            result.append(text[i])
            i += 1

        final_text = ''.join(result)

        if self.localization_dict:
            max_iterations = 10
            iteration = 0

            while iteration < max_iterations:
                pattern = r'\$([^$]+)\$'
                matches = list(re.finditer(pattern, final_text))

                if not matches:
                    break

                for match in reversed(matches):
                    full_match = match.group(0)
                    key = match.group(1)

                    if key in self.localization_dict:
                        replacement = self.localization_dict[key]

                        has_more_dollar_signs = bool(re.search(pattern, replacement))
                        is_newline_only = replacement == '\\n'

                        if not has_more_dollar_signs and not is_newline_only and force_blue:
                            replacement_without_color = re.sub(r'§[A-Z]|§!', '', replacement)
                            replacement = f"§A{replacement_without_color}§!"

                        start_pos = match.start()
                        end_pos = match.end()
                        final_text = final_text[:start_pos] + replacement + final_text[end_pos:]

                iteration += 1

                if iteration >= max_iterations:
                    print(f"警告: $$替换达到最大迭代次数 {max_iterations}")

                if not re.search(pattern, final_text):
                    break

        return final_text

    def format_modifier_value(self, key, value):
        """格式化修正值，进行百分比转换判定，并在描述和数值之间添加冒号"""
        try:
            value_float = float(value)
        except (ValueError, TypeError):
            return str(value)

        def should_format_as_percentage(modifier_key):
            if modifier_key in NON_PERCENTAGE_KEYS:
                return False
            if modifier_key in FORCE_PERCENTAGE_KEYS:
                return True
            if (modifier_key.endswith('_gain') or modifier_key.endswith('_perc') or
                    modifier_key.endswith('_mult') or modifier_key.endswith('_mult_add') or
                    modifier_key.endswith('_happiness') or modifier_key.endswith('_speed') or
                    modifier_key.endswith('_uninhabitable')):
                return True
            return False

        use_percentage = should_format_as_percentage(key)
        if key.endswith('_add') and key not in FORCE_PERCENTAGE_KEYS:
            use_percentage = False
        if use_percentage:
            percentage = value_float * 100
            if abs(percentage - round(percentage)) < 0.001:
                formatted = f"{int(percentage)}%"
            else:
                formatted = f"{percentage:.2f}%".rstrip('0').rstrip('.')
                if formatted.endswith('.'):
                    formatted = formatted[:-1]
                if not formatted.endswith('%'):
                    formatted += '%'
        else:
            if value_float == int(value_float):
                formatted = f"{int(value_float)}"
            else:
                formatted = f"{value_float}"
        sign = "+" if value_float >= 0 else ""

        return f":{sign}{formatted}"

    def extract_complete_block(self, content, start_pos):
        """提取完整的花括号块，处理嵌套结构"""
        brace_count = 0
        in_quotes = False
        escape_next = False

        for i in range(start_pos, len(content)):
            if content[i] == '{':
                start_pos = i
                break

        result = ""
        for i in range(start_pos, len(content)):
            char = content[i]
            result += char

            if escape_next:
                escape_next = False
                continue

            if char == '\\':
                escape_next = True
                continue

            if char == '"' and not escape_next:
                in_quotes = not in_quotes
                continue

            if not in_quotes:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        return result

        return None

    def extract_potential_block(self, civic_block):
        """提取potential块的内容"""
        potential_match = re.search(r'potential\s*=\s*\{', civic_block)
        if not potential_match:
            return None

        start_pos = potential_match.end() - 1
        potential_block = self.extract_complete_block(civic_block, start_pos)
        return potential_block

    def extract_possible_block(self, civic_block):
        """提取possible块的内容"""
        possible_match = re.search(r'possible\s*=\s*\{', civic_block)
        if not possible_match:
            return None

        start_pos = possible_match.end() - 1
        possible_block = self.extract_complete_block(civic_block, start_pos)
        return possible_block

    def parse_potential_authority(self, potential_content):
        """解析potential块中的authority条件，判断国策类型"""
        if not potential_content:
            return None

        if potential_content.startswith('potential = {'):
            potential_content = potential_content[len('potential = {'):-1].strip()

        authority_pattern = r'authority\s*=\s*(\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\})'
        authority_blocks = re.findall(authority_pattern, potential_content)

        for authority_block in authority_blocks:
            value_matches = re.findall(r'value\s*=\s*(\w+)', authority_block)
            for value in value_matches:
                if value == 'auth_hive_mind':
                    return "蜂巢"
                elif value == 'auth_machine_intelligence':
                    return "智械"

            logic_pattern = r'(OR|NOT|NOR)\s*=\s*(\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\})'
            logic_blocks = re.findall(logic_pattern, authority_block)

            for logic_type, logic_content in logic_blocks:
                value_matches = re.findall(r'value\s*=\s*(\w+)', logic_content)
                for value in value_matches:
                    if value == 'auth_hive_mind':
                        return "蜂巢"
                    elif value == 'auth_machine_intelligence':
                        return "智械"

        return None

    def parse_possible_conditions(self, possible_content):
        """解析possible块中的所有条件"""
        all_conditions = []

        if not possible_content:
            return all_conditions

        if possible_content.startswith('possible = {'):
            possible_content = possible_content[len('possible = {'):-1].strip()

        block_types = ['authority', 'ethics', 'origin', 'civics', 'species_class', 'species_archetype',
                       'graphical_culture']
        blocks = []

        for block_type in block_types:
            pattern = rf'{block_type}\s*=\s*(\{{[^{{}}]*(?:\{{[^{{}}]*\}}[^{{}}]*)*\}})'
            matches = re.findall(pattern, possible_content)
            for match in matches:
                blocks.append((block_type, match))

        for block_type, block_content in blocks:
            conditions = self.parse_block_corrected(block_content.strip(), block_type)
            all_conditions.extend(conditions)

        return all_conditions

    def parse_block_corrected(self, block_content, block_type):
        """解析单个块的内容"""
        conditions = []

        block_type_translation = CONDITION_TYPE_TRANSLATIONS.get(block_type, block_type)

        logic_pattern = r'(OR|NOT|NOR)\s*=\s*(\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\})'
        logic_blocks = re.findall(logic_pattern, block_content)

        for logic_type, logic_content in logic_blocks:
            text_match = re.search(r'text\s*=\s*(\S+)', logic_content)
            if text_match:
                text_key = text_match.group(1)
                text_translation = self.resolve_localisation(self.localization_dict.get(text_key, text_key))
                text_with_replacements = self.apply_special_replacements_with_force_blue(text_translation)
                conditions.append(text_with_replacements)
            else:
                value_matches = re.findall(r'value\s*=\s*(\w+)', logic_content)

                if logic_type == "OR" and len(value_matches) > 1:
                    conditions.append("以下条件至少一个满足：")
                    for value in value_matches:
                        condition_text = self.create_value_condition(value, logic_type, block_type_translation)
                        conditions.append(f"  {condition_text}")
                else:
                    for value in value_matches:
                        condition_text = self.create_value_condition(value, logic_type, block_type_translation)
                        conditions.append(condition_text)

        block_without_logic = block_content
        for logic_type, logic_content in logic_blocks:
            block_without_logic = block_without_logic.replace(logic_content, "")

        direct_value_match = re.search(r'value\s*=\s*(\w+)', block_without_logic)
        if direct_value_match:
            value = direct_value_match.group(1)
            condition_text = self.create_value_condition(value, "NEED", block_type_translation)
            conditions.append(condition_text)

        return conditions

    def apply_basic_replacements_no_color(self, text):
        """基本的文本替换，不添加任何颜色标记"""
        if not text:
            return text

        max_iterations = 5
        iteration = 0

        while iteration < max_iterations:
            pattern = r'\$([^$]+)\$'
            matches = list(re.finditer(pattern, text))

            if not matches:
                break

            for match in reversed(matches):
                key = match.group(1)
                if key in self.localization_dict:
                    replacement = self.localization_dict[key]
                    text = text[:match.start()] + replacement + text[match.end():]

            iteration += 1

        return text

    def add_orange_color_to_value(self, full_text, value_text):
        """在完整文本中为value部分添加橙色标记"""
        if not value_text:
            return full_text

        start_pos = full_text.find(value_text)
        if start_pos == -1:
            return full_text

        end_pos = start_pos + len(value_text)

        # 在添加橙色标记前，先清除value_text中可能存在的颜色标记
        value_without_color = re.sub(r'§[A-Z]|§!', '', value_text)

        # 更新full_text中的value_text部分
        full_text_without_color = full_text[:start_pos] + value_without_color + full_text[end_pos:]

        # 重新计算位置
        new_start_pos = full_text_without_color.find(value_without_color)
        if new_start_pos == -1:
            return full_text

        new_end_pos = new_start_pos + len(value_without_color)

        # 添加橙色标记
        result = full_text_without_color[:new_start_pos] + f"§H{value_without_color}§!" + full_text_without_color[
            new_end_pos:]
        return result

    def add_blue_underline_to_value(self, full_text, value_text):
        """在完整文本中为value部分添加蓝色下划线标记"""
        if not value_text:
            return full_text

        start_pos = full_text.find(value_text)
        if start_pos == -1:
            return full_text

        end_pos = start_pos + len(value_text)

        # 在添加蓝色下划线前，先清除value_text中可能存在的颜色标记
        value_without_color = re.sub(r'§[A-Z]|§!', '', value_text)

        # 更新full_text中的value_text部分
        full_text_without_color = full_text[:start_pos] + value_without_color + full_text[end_pos:]

        # 重新计算位置
        new_start_pos = full_text_without_color.find(value_without_color)
        if new_start_pos == -1:
            return full_text

        new_end_pos = new_start_pos + len(value_without_color)

        # 添加蓝色下划线标记
        result = full_text_without_color[:new_start_pos] + f"§A{value_without_color}§!" + full_text_without_color[
            new_end_pos:]
        return result

    def create_value_condition(self, value, logic_type, block_type_translation):
        """创建value类型的条件文本"""
        value_translation_raw = self.localization_dict.get(value, value)

        value_translation_clean = self.apply_basic_replacements_no_color(value_translation_raw)

        condition_text = ""

        if block_type_translation == "政体":
            if logic_type in ["NEED", "OR"]:
                condition_text = f"拥有{value_translation_clean}{block_type_translation}"
            else:
                condition_text = f"不具有{value_translation_clean}{block_type_translation}"
        elif block_type_translation == "思潮":
            if logic_type in ["NEED", "OR"]:
                condition_text = f"需要{value_translation_clean}{block_type_translation}"
            else:
                condition_text = f"没有{value_translation_clean}{block_type_translation}"
        elif block_type_translation == "国民理念":
            if logic_type in ["NEED", "OR"]:
                condition_text = f"需要{value_translation_clean}{block_type_translation}"
            else:
                condition_text = f"没有{value_translation_clean}{block_type_translation}"
        elif block_type_translation == "物种类型" or block_type_translation == "物种":
            if logic_type in ["NEED", "OR"]:
                condition_text = f"需要物种类型：{value_translation_clean}"
            else:
                condition_text = f"不具有物种类型：{value_translation_clean}"
        elif block_type_translation == "":
            if logic_type in ["NEED", "OR"]:
                condition_text = f"需要{value_translation_clean}"
            else:
                condition_text = f"不具有{value_translation_clean}"
        else:
            if logic_type in ["NEED", "OR"]:
                condition_text = f"需要{value_translation_clean}{block_type_translation}"
            else:
                condition_text = f"不具有{value_translation_clean}{block_type_translation}"

        # 现在根据块类型添加颜色标记
        if block_type_translation in ["政体", "思潮", "物种类型", "物种"]:
            # 使用橙色
            condition_text_with_color = self.add_orange_color_to_value(condition_text, value_translation_clean)
        else:
            # 使用蓝色下划线
            condition_text_with_color = self.add_blue_underline_to_value(condition_text, value_translation_clean)

        return condition_text_with_color

    def extract_outer_modifier_blocks(self, content):
        """提取所有最外层的modifier块"""
        modifier_blocks = []
        pos = 0

        while True:
            modifier_pos = content.find('modifier = {', pos)
            if modifier_pos == -1:
                break

            preceding_content = content[:modifier_pos]
            open_braces = preceding_content.count('{')
            close_braces = preceding_content.count('}')

            if open_braces == close_braces + 1:
                modifier_block = self.extract_complete_block(content, modifier_pos)
                if modifier_block:
                    modifier_blocks.append(modifier_block)

            pos = modifier_pos + 1

        return modifier_blocks

    def parse_modifier_block(self, modifier_block, civic_key=None):
        """解析单个modifier块的内容"""
        effects = []

        start_idx = modifier_block.find('{') + 1
        end_idx = modifier_block.rfind('}')

        if start_idx >= 0 and end_idx >= 0 and start_idx < end_idx:
            modifier_content = modifier_block[start_idx:end_idx].strip()
        else:
            modifier_content = modifier_block.replace('modifier = {', '').replace('}', '').strip()

        modifier_pattern = r'(\w+)\s*=\s*([\d.-]+|@\w+)'
        modifier_matches = re.findall(modifier_pattern, modifier_content)

        for modifier_key, modifier_value in modifier_matches:
            if modifier_key in ['mult', 'potential']:
                continue

            if modifier_value.startswith('@'):
                if modifier_value in self.scripted_vars:
                    actual_value = self.scripted_vars[modifier_value]
                    modifier_value = actual_value
                else:
                    continue

            translation, missing_key = self.find_modifier_translation(modifier_key)
            if translation:
                value_str = self.format_modifier_value(modifier_key, modifier_value)

                try:
                    value_float = float(modifier_value)
                    # 使用颜色转换字典判断颜色
                    if any(modifier_key.endswith(suffix) for suffix in SPECIAL_SUFFIXES):
                        color = '§G' if value_float < 0 else '§R'
                    else:
                        color = '§G' if value_float >= 0 else '§R'
                except (ValueError, TypeError):
                    color = ''

                # 先进行完整的替换符处理（包括[]和$$）
                effect_text_with_replacements = self.apply_special_replacements_with_force_blue(
                    translation)

                # 然后添加数值部分
                effect_text = f"{effect_text_with_replacements}{value_str}"

                # 对于modifier效果，清除所有颜色标记（包括蓝色下划线）
                # 但保留数值部分的颜色，冒号不带有颜色
                colon_pos = effect_text.find(':')
                if colon_pos != -1:
                    description_part = effect_text[:colon_pos]
                    colon_char = effect_text[colon_pos]  # 冒号本身
                    value_part = effect_text[colon_pos + 1:]  # 冒号后面的部分

                    # 清除描述部分的所有颜色标记
                    description_clean = re.sub(r'§[A-Z]|§!', '', description_part)

                    # 为数值部分添加颜色
                    if color:
                        value_part = f"{color}{value_part}§!"

                    # 重新组合：描述部分（无颜色）+ 冒号（无颜色）+ 数值部分（有颜色）
                    effect_text = description_clean + colon_char + value_part

                else:
                    # 如果没有冒号，清除整个效果文本的颜色标记
                    effect_text = re.sub(r'§[A-Z]|§!', '', effect_text)
            else:
                # 未找到本地化的modifier
                value_str = self.format_modifier_value(modifier_key, modifier_value)
                effect_text = f"{modifier_key}{value_str}"
                print(f"警告: 未找到modifier的本地化: {modifier_key}")

            effects.append(effect_text)

        return effects

    def is_dlc_only_trigger(self, trigger_content):
        """检查trigger是否只包含DLC条件"""
        clean_trigger = re.sub(r'\s+', '', trigger_content)

        has_dlc_pattern = r'has_\w+_dlc=yes'
        host_has_dlc_pattern = r'host_has_dlc="[^"]*"'
        is_scope_valid_pattern = r'is_scope_valid=yes'

        dlc_conditions = len(re.findall(has_dlc_pattern, clean_trigger))
        dlc_conditions += len(re.findall(host_has_dlc_pattern, clean_trigger))
        scope_conditions = len(re.findall(is_scope_valid_pattern, clean_trigger))

        total_conditions = dlc_conditions + scope_conditions
        other_conditions = len(re.findall(r'(\w+)=', clean_trigger)) - total_conditions

        return other_conditions == 0 and dlc_conditions > 0

    def extract_swap_type_modifiers(self, civic_block, civic_key=None):
        """提取swap_type中的modifier - 仅在trigger只有DLC条件时"""
        swap_type_modifiers = []

        swap_type_positions = []
        pos = 0
        while True:
            swap_type_pos = civic_block.find('swap_type = {', pos)
            if swap_type_pos == -1:
                break
            swap_type_positions.append(swap_type_pos)
            pos = swap_type_pos + 1

        for swap_type_pos in swap_type_positions:
            swap_type_block = self.extract_complete_block(civic_block, swap_type_pos)
            if not swap_type_block:
                continue

            trigger_match = re.search(r'trigger\s*=\s*\{([^}]+)\}', swap_type_block)
            if trigger_match:
                trigger_content = trigger_match.group(1)

                if self.is_dlc_only_trigger(trigger_content):
                    modifier_blocks = self.extract_outer_modifier_blocks(swap_type_block)
                    for modifier_block in modifier_blocks:
                        effects = self.parse_modifier_block(modifier_block, civic_key)
                        swap_type_modifiers.extend(effects)

        return swap_type_modifiers

    def determine_civic_type(self, file_path, potential_content):
        """根据文件路径和potential块确定国策类型"""
        filename = file_path.name.lower()

        if "03_corporate_civics" in filename:
            civic_type = "巨企"
        elif "01_special_civics" in filename:
            civic_type = "特殊"
        else:
            civic_type = "普政"

        authority_type = self.parse_potential_authority(potential_content)
        if authority_type:
            civic_type = authority_type

        return civic_type

    def extract_civic_image(self, civic_block, civic_key):
        """提取国策图片路径"""
        # 首先尝试从icon字段提取
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', civic_block)
        if icon_match:
            icon_path = icon_match.group(1)
            # 构建完整的DDS文件路径
            dds_path = os.path.join(BASE_GAME_DIR, icon_path.replace('/', os.path.sep))

            # 生成PNG文件路径
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(CIVIC_IMAGES_DIR, png_filename)

            # 如果PNG图片已存在，直接返回路径
            if os.path.exists(png_path):
                return png_path
            else:
                # 如果PNG不存在，但DDS存在，则进行转换
                if os.path.exists(dds_path):
                    print(f"发现DDS图片，正在转换为PNG: {dds_filename}")
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        # 记录缺失的图片
                        self.debug_info['missing_images'].append(f"{civic_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{civic_key} -> DDS文件不存在: {dds_path}")
                    return None

        # 如果没有icon字段，使用默认路径 + key.dds
        default_dds_path = os.path.join(CIVIC_GFX_FILE_PATH, f"{civic_key}.dds")
        png_filename = f"{civic_key}.png"
        png_path = os.path.join(CIVIC_IMAGES_DIR, png_filename)

        # 如果PNG图片已存在，直接返回路径
        if os.path.exists(png_path):
            return png_path
        else:
            # 如果PNG不存在，但DDS存在，则进行转换
            if os.path.exists(default_dds_path):
                print(f"发现默认DDS图片，正在转换为PNG: {civic_key}.dds")
                success = self.convert_dds_to_png(default_dds_path, png_path)
                if success:
                    return png_path
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{civic_key} -> 转换失败: {civic_key}.dds")
                    return None
            else:
                print(f"警告: 未找到国策图片: {png_path}")
                return None

    def extract_origin_image(self, origin_block, origin_key):
        """提取起源图片路径"""
        # 从origin_block中提取icon路径
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', origin_block)
        if icon_match:
            icon_path = icon_match.group(1)
            # 构建完整的DDS文件路径
            dds_path = os.path.join(BASE_GAME_DIR, icon_path.replace('/', os.path.sep))

            # 生成PNG文件路径
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(ORIGIN_IMAGES_DIR, png_filename)

            # 如果PNG图片已存在，直接返回路径
            if os.path.exists(png_path):
                return png_path
            else:
                # 如果PNG不存在，但DDS存在，则进行转换
                if os.path.exists(dds_path):
                    print(f"发现DDS图片，正在转换为PNG: {dds_filename}")
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        # 记录缺失的图片
                        self.debug_info['missing_images'].append(f"{origin_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    print(f"警告: 未找到起源图片: {dds_path}")
                    return None

        # 如果无法从icon提取，尝试使用默认路径 + key.dds
        default_dds_path = os.path.join(ORIGIN_GFX_FILE_PATH, f"{origin_key}.dds")
        png_filename = f"{origin_key}.png"
        png_path = os.path.join(ORIGIN_IMAGES_DIR, png_filename)

        if os.path.exists(png_path):
            return png_path
        else:
            # 如果PNG不存在，但DDS存在，则进行转换
            if os.path.exists(default_dds_path):
                print(f"发现默认DDS图片，正在转换为PNG: {origin_key}.dds")
                success = self.convert_dds_to_png(default_dds_path, png_path)
                if success:
                    return png_path
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{origin_key} -> 转换失败: {origin_key}.dds")
                    return None
            else:
                print(f"警告: 未找到起源图片: {png_path}")
                return None

    def parse_civic_files(self):
        """解析国策文件，提取国策定义"""
        print("正在解析国策文件...")

        for file_path in Path(CIVIC_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                    pattern = r'^(\w+)\s*=\s*\{'
                    matches = list(re.finditer(pattern, content, re.MULTILINE))

                    for match in matches:
                        civic_key = match.group(1)
                        if not civic_key.startswith('civic_'):
                            continue

                        start_pos = match.start()
                        civic_block = self.extract_complete_block(content, start_pos)

                        if not civic_block:
                            continue

                        description_match = re.search(r'description\s*=\s*"([^"]*)"', civic_block)
                        if not description_match:
                            description_match = re.search(r'description\s*=\s*(\S+)', civic_block)
                        description = description_match.group(1) if description_match else ""

                        negative_description_match = re.search(r'negative_description\s*=\s*"([^"]*)"', civic_block)
                        if not negative_description_match:
                            negative_description_match = re.search(r'negative_description\s*=\s*(\S+)', civic_block)
                        negative_description = negative_description_match.group(1) if negative_description_match else ""

                        # 处理custom_tooltip（保留颜色）
                        custom_tooltip_effects = []
                        custom_tooltip_match = re.search(r'custom_tooltip_with_modifiers\s*=\s*"([^"]*)"', civic_block)
                        if not custom_tooltip_match:
                            custom_tooltip_match = re.search(r'custom_tooltip_with_modifiers\s*=\s*(\S+)', civic_block)
                        if not custom_tooltip_match:
                            custom_tooltip_match = re.search(r'custom_tooltip\s*=\s*"([^"]*)"', civic_block)
                        if not custom_tooltip_match:
                            custom_tooltip_match = re.search(r'custom_tooltip\s*=\s*(\S+)', civic_block)

                        if custom_tooltip_match:
                            tooltip_key = custom_tooltip_match.group(1)
                            tooltip_key = tooltip_key.strip('"')
                            tooltip_translation = self.resolve_localisation(
                                self.localization_dict.get(tooltip_key, tooltip_key))
                            # 对custom_tooltip进行完整的替换符处理，但保留颜色
                            tooltip_final = self.apply_special_replacements_with_force_blue(
                                tooltip_translation)
                            custom_tooltip_effects.append(tooltip_final)

                        # 处理modifier效果（清理颜色）
                        modifier_effects = []
                        outer_modifier_blocks = self.extract_outer_modifier_blocks(civic_block)

                        for modifier_block in outer_modifier_blocks:
                            effects = self.parse_modifier_block(modifier_block, civic_key)
                            modifier_effects.extend(effects)

                        swap_type_modifiers = self.extract_swap_type_modifiers(civic_block, civic_key)
                        modifier_effects.extend(swap_type_modifiers)

                        # 合并效果：先custom_tooltip（保留颜色），后modifier（清理颜色）
                        all_effects = custom_tooltip_effects + modifier_effects

                        potential_content = self.extract_potential_block(civic_block)
                        possible_content = self.extract_possible_block(civic_block)
                        possible_conditions = self.parse_possible_conditions(possible_content)

                        civic_type = self.determine_civic_type(file_path, potential_content)

                        # 提取图片路径
                        image_path = self.extract_civic_image(civic_block, civic_key)

                        self.civics.append({
                            'key': civic_key,
                            'description_key': description,
                            'negative_description_key': negative_description,
                            'modifier_effects': all_effects,  # 合并后的效果
                            'possible_conditions': possible_conditions,
                            'civic_type': civic_type,
                            'image_path': image_path
                        })

            except Exception as e:
                print(f"解析国策文件时出错 {file_path}: {e}")

        print(f"已找到 {len(self.civics)} 个国策定义")

    def parse_councilor_files(self):
        """解析内阁官员文件，建立国策到内阁岗位的映射"""
        print("正在解析内阁官员文件...")

        for file_path in Path(COUNCILORS_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                    pattern = r'^(\w+)\s*=\s*\{'
                    matches = list(re.finditer(pattern, content, re.MULTILINE))

                    for match in matches:
                        councilor_key = match.group(1)
                        if not councilor_key.startswith('councilor_'):
                            continue

                        start_pos = match.start()
                        councilor_block = self.extract_complete_block(content, start_pos)

                        if not councilor_block:
                            continue

                        civic_match = re.search(r'civic\s*=\s*(\w+)', councilor_block)
                        if not civic_match:
                            continue

                        civic_key = civic_match.group(1)

                        leader_class_match = re.search(r'leader_class\s*=\s*\{([^}]+)\}', councilor_block)
                        leader_classes = []
                        if leader_class_match:
                            classes_content = leader_class_match.group(1)
                            class_matches = re.findall(r'(\w+)', classes_content)
                            leader_classes = class_matches

                        modifier_effects = []

                        # 处理custom_tooltip（保留颜色，但不为替换符添加新颜色）
                        custom_tooltip_match = re.search(r'custom_tooltip_with_modifiers\s*=\s*"([^"]*)"',
                                                         councilor_block)
                        if not custom_tooltip_match:
                            custom_tooltip_match = re.search(r'custom_tooltip_with_modifiers\s*=\s*(\S+)',
                                                             councilor_block)
                        if not custom_tooltip_match:
                            custom_tooltip_match = re.search(r'custom_tooltip\s*=\s*"([^"]*)"', councilor_block)
                        if not custom_tooltip_match:
                            custom_tooltip_match = re.search(r'custom_tooltip\s*=\s*(\S+)', councilor_block)

                        if custom_tooltip_match:
                            tooltip_key = custom_tooltip_match.group(1)
                            tooltip_key = tooltip_key.strip('"')
                            tooltip_translation = self.resolve_localisation(
                                self.localization_dict.get(tooltip_key, tooltip_key))
                            # 对custom_tooltip进行完整的替换符处理，但不添加新颜色标记
                            tooltip_final = self.apply_special_replacements_with_force_blue(
                                tooltip_translation, force_blue=False)
                            modifier_effects.append(tooltip_final)

                        # 处理modifier效果（清理颜色）
                        modifier_blocks = self.extract_outer_modifier_blocks(councilor_block)
                        for modifier_block in modifier_blocks:
                            effects = self.parse_modifier_block(modifier_block, civic_key)
                            modifier_effects.extend(effects)

                        self.councilors_dict[civic_key] = {
                            'councilor_key': councilor_key,
                            'leader_classes': leader_classes,
                            'modifier_effects': modifier_effects
                        }

            except Exception as e:
                print(f"解析内阁官员文件时出错 {file_path}: {e}")

        print(f"已找到 {len(self.councilors_dict)} 个内阁岗位映射")

    def parse_origin_files(self):
        """解析起源文件，提取起源定义"""
        print("正在解析起源文件...")

        for file_path in Path(CIVIC_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                    # 查找所有起源定义
                    pattern = r'^(\w+)\s*=\s*\{'
                    matches = list(re.finditer(pattern, content, re.MULTILINE))

                    for match in matches:
                        origin_key = match.group(1)
                        if not origin_key.startswith('origin_'):
                            continue

                        start_pos = match.start()
                        origin_block = self.extract_complete_block(content, start_pos)

                        if not origin_block:
                            continue

                        # 检查是否是真正的起源
                        is_origin_match = re.search(r'is_origin\s*=\s*yes', origin_block)
                        if not is_origin_match:
                            continue

                        # 提取描述
                        description_match = re.search(r'description\s*=\s*"([^"]*)"', origin_block)
                        if not description_match:
                            description_match = re.search(r'description\s*=\s*(\S+)', origin_block)
                        description = description_match.group(1) if description_match else ""

                        # 提取负面描述
                        negative_description_match = re.search(r'negative_description\s*=\s*"([^"]*)"', origin_block)
                        if not negative_description_match:
                            negative_description_match = re.search(r'negative_description\s*=\s*(\S+)', origin_block)
                        negative_description = negative_description_match.group(1) if negative_description_match else ""

                        # 检查non_colonizable_planet_class_neighbor
                        non_colonizable_match = re.search(r'non_colonizable_planet_class_neighbor\s*=\s*yes',
                                                          origin_block)
                        non_colonizable_text = ""
                        if non_colonizable_match:
                            non_colonizable_text = "确保宜居星球：不会生成"

                        # 提取modifier效果（与国策相同的方式）
                        modifier_effects = []
                        outer_modifier_blocks = self.extract_outer_modifier_blocks(origin_block)

                        for modifier_block in outer_modifier_blocks:
                            effects = self.parse_modifier_block(modifier_block, origin_key)
                            modifier_effects.extend(effects)

                        # 提取possible条件
                        possible_content = self.extract_possible_block(origin_block)
                        possible_conditions = self.parse_possible_conditions(possible_content)

                        # 提取图片路径
                        image_path = self.extract_origin_image(origin_block, origin_key)

                        self.origins.append({
                            'key': origin_key,
                            'description_key': description,
                            'negative_description_key': negative_description,
                            'non_colonizable_text': non_colonizable_text,
                            'modifier_effects': modifier_effects,
                            'possible_conditions': possible_conditions,
                            'image_path': image_path
                        })

            except Exception as e:
                print(f"解析起源文件时出错 {file_path}: {e}")

        print(f"已找到 {len(self.origins)} 个起源定义")

    def process_text_for_output_simple(self, text, is_name_or_desc=True):
        """简化的文本处理：先处理[]替换，不处理颜色，但保留\n"""
        # 先删除所有的£xxx£标记（不进行翻译，直接删除）
        text = re.sub(r'£[^£]*£', '', text)

        # 再处理特殊替换（[]标记），使用蓝色下划线
        processed_text = self.apply_special_replacements_with_force_blue(
            text, force_blue=True)

        # 处理$$替换符，根据规则决定是否添加颜色
        processed_text = self.preprocess_text_simple(processed_text, avoid_orange=is_name_or_desc)

        return processed_text

    def process_color_text_final(self, text):
        """最终颜色处理：处理所有颜色标记，返回纯文本和颜色信息"""
        if not text:
            return text, []

        # 第一步：清理嵌套的颜色标记和空颜色标记
        text = self.clean_nested_color_marks(text)

        # 第二步：进行正常的颜色处理
        color_pattern = r'(§[A-Z])(.*?)(§!)'
        color_matches = list(re.finditer(color_pattern, text))

        if not color_matches:
            return text, []

        # 移除颜色标记的纯文本
        clean_text = re.sub(r'§[A-Z]|§!', '', text)

        # 计算颜色位置
        color_info = []

        for match in color_matches:
            color_code = match.group(1)
            colored_text = match.group(2)
            end_marker = match.group(3)

            text_before_match = text[:match.start()]
            clean_pos_start = len(re.sub(r'§[A-Z]|§!', '', text_before_match))
            clean_pos_end = clean_pos_start + len(colored_text)

            if color_code in COLOR_MAP:
                color_name = COLOR_MAP[color_code]
                underline = color_name.endswith('_underline')
                if underline:
                    color_name = color_name.replace('_underline', '')

                color_info.append((clean_pos_start, clean_pos_end, color_name, underline))

        return clean_text, color_info

    def remove_empty_color_marks(self, text):
        """移除空的颜色标记对"""
        if not text:
            return text

        empty_color_pattern = r'§[A-Z]§!'
        cleaned_text = re.sub(empty_color_pattern, '', text)
        return cleaned_text

    def clean_nested_color_marks(self, text):
        """清理嵌套的颜色标记，但保留颜色标记本身"""
        if not text:
            return text

        text = self.remove_empty_color_marks(text)

        max_clean_iterations = 5
        clean_iteration = 0

        while clean_iteration < max_clean_iterations:
            color_pattern = r'(§[A-Z])(.*?)(§!)'
            color_matches = list(re.finditer(color_pattern, text))
            if not color_matches:
                break

            changed = False

            for match in reversed(color_matches):
                color_code = match.group(1)
                colored_text = match.group(2)
                end_marker = match.group(3)

                if re.search(r'§[A-Z]', colored_text):
                    clean_colored_text = colored_text
                    replacement = clean_colored_text

                    text = text[:match.start()] + replacement + text[match.end():]
                    changed = True

            if not changed:
                break

            clean_iteration += 1

        return text

    def remove_duplicate_text(self, text):
        """删除连续的重复文本，但跳过纯数字和单个字符"""
        if not text or len(text) < 4:
            return text

        # 首先提取颜色信息
        color_pattern = r'(§[A-Z])(.*?)(§!)'
        color_matches = list(re.finditer(color_pattern, text))

        # 创建纯文本版本用于比较
        clean_text = re.sub(r'§[A-Z]|§!', '', text)

        # 查找重复文本（长度2-10个字符）
        max_dup_length = 10
        min_dup_length = 2

        # 记录所有重复段落的起始和结束位置
        duplicate_ranges = []

        for length in range(max_dup_length, min_dup_length - 1, -1):
            i = 0
            while i <= len(clean_text) - 2 * length:
                segment1 = clean_text[i:i + length]
                segment2 = clean_text[i + length:i + 2 * length]

                # 跳过纯数字的重复
                if segment1.isdigit() and segment2.isdigit():
                    i += 1
                    continue

                # 跳过单个字符的重复（除非是特殊字符）
                if length == 1 and segment1.isalnum():
                    i += 1
                    continue

                if segment1 == segment2:
                    # 计算在原始文本中的对应位置
                    start_pos1 = self.find_original_position(text, clean_text, i)
                    end_pos1 = self.find_original_position(text, clean_text, i + length)
                    start_pos2 = self.find_original_position(text, clean_text, i + length)
                    end_pos2 = self.find_original_position(text, clean_text, i + 2 * length)

                    # 检查这个范围是否已经被标记为重复
                    is_already_duplicate = False
                    for dup_start, dup_end in duplicate_ranges:
                        if (start_pos1 >= dup_start and end_pos1 <= dup_end) or \
                                (start_pos2 >= dup_start and end_pos2 <= dup_end):
                            is_already_duplicate = True
                            break

                    if not is_already_duplicate:
                        # 检查两个段落是否有颜色
                        has_color1 = self.has_color_in_range(text, start_pos1, end_pos1)
                        has_color2 = self.has_color_in_range(text, start_pos2, end_pos2)

                        # 决定保留哪个段落
                        if has_color1 and not has_color2:
                            # 保留第一个（有颜色）
                            text = text[:start_pos2] + text[end_pos2:]
                            duplicate_ranges.append((start_pos2, end_pos2))
                        elif has_color2 and not has_color1:
                            # 保留第二个（有颜色），删除第一个
                            text = text[:start_pos1] + text[end_pos1:]
                            duplicate_ranges.append((start_pos1, end_pos1))
                        else:
                            # 都有颜色或都没有颜色，保留第二个，删除第一个
                            text = text[:start_pos1] + text[end_pos1:]
                            duplicate_ranges.append((start_pos1, end_pos1))

                        # 重新计算clean_text
                        clean_text = re.sub(r'§[A-Z]|§!', '', text)
                        break
                i += 1

        return text

    def find_original_position(self, original_text, clean_text, clean_pos):
        """找到纯文本位置在原始文本中的对应位置"""
        clean_prefix = clean_text[:clean_pos]

        original_pos = 0
        clean_index = 0

        while clean_index < len(clean_prefix):
            while original_pos < len(original_text) and original_text[original_pos] == '§':
                if original_pos + 1 < len(original_text) and original_text[
                    original_pos + 1] in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                    original_pos += 2
                elif original_pos + 1 < len(original_text) and original_text[original_pos + 1] == '!':
                    original_pos += 2
                else:
                    break

            if clean_index < len(clean_prefix) and original_pos < len(original_text):
                if original_text[original_pos] == clean_prefix[clean_index]:
                    clean_index += 1
                    original_pos += 1
                else:
                    original_pos += 1
            else:
                break

        return original_pos

    def has_color_in_range(self, text, start, end):
        """检查指定范围内的文本是否有颜色标记"""
        segment = text[start:end]
        return bool(re.search(r'§[A-Z]', segment))

    def generate_integrated_description(self, civic):
        """生成国策描述整合内容"""
        civic_key = civic['key']

        # 获取国策名称
        civic_name_raw = self.localization_dict.get(civic_key, civic_key)
        civic_name_final = self.process_text_for_output_simple(
            civic_name_raw, is_name_or_desc=True)
        civic_name_final = self.remove_duplicate_text(civic_name_final)

        # 获取国策描述
        desc_key = f"{civic_key}_desc"
        civic_desc_raw = self.localization_dict.get(desc_key, "")
        civic_desc_final = self.process_text_for_output_simple(
            civic_desc_raw, is_name_or_desc=True)
        civic_desc_final = self.remove_duplicate_text(civic_desc_final)

        # 获取国策说明效果
        description_key = civic['description_key']
        civic_effects_desc_raw = self.localization_dict.get(description_key, description_key)
        civic_effects_desc_final = self.process_text_for_output_simple(
            civic_effects_desc_raw, is_name_or_desc=False)
        civic_effects_desc_final = self.remove_duplicate_text(civic_effects_desc_final)
        civic_effects_desc_final = re.sub(r'£[^£]*£', '', civic_effects_desc_final)

        # 获取国策惩罚
        negative_description_key = civic['negative_description_key']
        civic_negative_effects_raw = self.localization_dict.get(negative_description_key, negative_description_key)
        civic_negative_effects_final = self.process_text_for_output_simple(
            civic_negative_effects_raw, is_name_or_desc=False)
        civic_negative_effects_final = self.remove_duplicate_text(civic_negative_effects_final)
        civic_negative_effects_final = re.sub(r'£[^£]*£', '', civic_negative_effects_final)

        # 获取国策具体效果
        civic_effects_list = civic['modifier_effects'] if civic['modifier_effects'] else []

        # 获取国策条件
        possible_conditions_list = civic['possible_conditions'] if civic['possible_conditions'] else []

        # 获取内阁席位信息
        councilor_info = self.councilors_dict.get(civic_key, {})

        # 开始构建描述整合
        integrated_description = []

        # 1. 国策名称
        integrated_description.append(f"{civic_name_final.replace('\\n', '')}：")
        integrated_description.append("")

        # 2. 国策描述
        if civic_desc_final:
            integrated_description.append(f"{civic_desc_final.replace('\\n', '')}")
            integrated_description.append("")

        # 3. 效果（来自description_key）
        if civic_effects_desc_final and civic_effects_desc_final != description_key:
            integrated_description.append("效果：")
            # 处理多行效果，每行前面添加一个空格
            effect_lines = civic_effects_desc_final.replace('\\n', '\n').split('\n')
            for line in effect_lines:
                if line.strip():
                    integrated_description.append(f" {line.strip()}")
            integrated_description.append("")

        # 4. 修正（来自modifier_effects）
        if civic_effects_list:
            integrated_description.append("修正：")
            for effect in civic_effects_list:
                effect_final = self.process_text_for_output_simple(
                    effect, is_name_or_desc=False)
                effect_final = self.remove_duplicate_text(effect_final)
                effect_final = re.sub(r'£[^£]*£', '', effect_final)
                integrated_description.append(f" -{effect_final.replace('\\n', '')}")
            integrated_description.append("")

        # 5. 惩罚（红色字体）
        if civic_negative_effects_final and civic_negative_effects_final != negative_description_key:
            integrated_description.append("§R惩罚：§!")
            integrated_description.append(f"§R -{civic_negative_effects_final.replace('\\n', '')}§!")
            integrated_description.append("")

        # 6. 内阁席位
        if councilor_info:
            councilor_key = councilor_info['councilor_key']
            councilor_name_raw = self.localization_dict.get(councilor_key, councilor_key)
            councilor_name_final = self.process_text_for_output_simple(
                councilor_name_raw, is_name_or_desc=False)
            councilor_name_final = self.remove_duplicate_text(councilor_name_final)

            councilor_effects_list = councilor_info['modifier_effects'] if councilor_info['modifier_effects'] else []

            integrated_description.append("内阁席位：")
            integrated_description.append(f"{councilor_name_final.replace('\\n', '')}：")

            # 每技能等级的帝国效果（橙色）
            if councilor_effects_list:
                # 将"技能等级"和"帝国效果"设为橙色
                integrated_description.append(" 每§H技能等级§!的§H帝国效果§!：")
                for effect in councilor_effects_list:
                    effect_final = self.process_text_for_output_simple(
                        effect, is_name_or_desc=False)
                    effect_final = self.remove_duplicate_text(effect_final)
                    effect_final = re.sub(r'£[^£]*£', '', effect_final)
                    integrated_description.append(f"  -{effect_final.replace('\\n', '')}")

            # 可选职业（橙色）
            if councilor_info['leader_classes']:
                leader_class_translations = {
                    'official': '行政官',
                    'commander': '指挥官',
                    'scientist': '科学家'
                }

                translated_classes = []
                for cls in councilor_info['leader_classes']:
                    if cls in leader_class_translations:
                        translated_classes.append(leader_class_translations[cls])
                    else:
                        translated_classes.append(cls)

                priority_order = ['行政官', '指挥官', '科学家']
                sorted_classes = []
                for priority in priority_order:
                    if priority in translated_classes:
                        sorted_classes.append(priority)

                for cls in translated_classes:
                    if cls not in sorted_classes:
                        sorted_classes.append(cls)

                integrated_description.append(" 可选：")
                for cls in sorted_classes:
                    integrated_description.append(f"  -{cls}")

            integrated_description.append("")

        # 7. 要求
        if possible_conditions_list:
            integrated_description.append("要求：")
            for condition in possible_conditions_list:
                condition_final = self.remove_duplicate_text(condition)
                condition_final = re.sub(r'£[^£]*£', '', condition_final)

                # 处理"以下条件至少一个满足"的特殊格式
                if condition_final.startswith("以下条件至少一个满足："):
                    integrated_description.append(f" -{condition_final}")
                elif "  " in condition_final:  # 这是子条件（有缩进）
                    # 移除前面的"-"符号，只保留缩进
                    clean_condition = condition_final.lstrip('-').strip()
                    integrated_description.append(f"   {clean_condition}")
                else:
                    integrated_description.append(f" -{condition_final.replace('\\n', '')}")

        return "\n".join(integrated_description)

    def generate_origin_integrated_description(self, origin):
        """生成起源描述整合内容"""
        origin_key = origin['key']

        # 获取起源名称
        origin_name_raw = self.localization_dict.get(origin_key, origin_key)
        origin_name_final = self.process_text_for_output_simple(
            origin_name_raw, is_name_or_desc=True)
        origin_name_final = self.remove_duplicate_text(origin_name_final)

        # 获取起源描述
        desc_key = f"{origin_key}_desc"
        origin_desc_raw = self.localization_dict.get(desc_key, "")
        origin_desc_final = self.process_text_for_output_simple(
            origin_desc_raw, is_name_or_desc=True)
        origin_desc_final = self.remove_duplicate_text(origin_desc_final)

        # 获取起源说明效果
        description_key = origin['description_key']
        origin_effects_desc_raw = self.localization_dict.get(description_key, description_key)
        origin_effects_desc_final = self.process_text_for_output_simple(
            origin_effects_desc_raw, is_name_or_desc=False)
        origin_effects_desc_final = self.remove_duplicate_text(origin_effects_desc_final)
        origin_effects_desc_final = re.sub(r'£[^£]*£', '', origin_effects_desc_final)

        # 获取起源惩罚
        negative_description_key = origin['negative_description_key']
        origin_negative_effects_raw = self.localization_dict.get(negative_description_key, negative_description_key)
        origin_negative_effects_final = self.process_text_for_output_simple(
            origin_negative_effects_raw, is_name_or_desc=False)
        origin_negative_effects_final = self.remove_duplicate_text(origin_negative_effects_final)
        origin_negative_effects_final = re.sub(r'£[^£]*£', '', origin_negative_effects_final)

        # 获取non_colonizable文本
        non_colonizable_text = origin['non_colonizable_text']

        # 获取起源具体效果
        origin_effects_list = origin['modifier_effects'] if origin['modifier_effects'] else []

        # 获取起源条件
        possible_conditions_list = origin['possible_conditions'] if origin['possible_conditions'] else []

        # 开始构建描述整合
        integrated_description = []

        # 1. 起源名称
        integrated_description.append(f"{origin_name_final.replace('\\n', '')}：")
        integrated_description.append("")

        # 2. 起源描述
        if origin_desc_final:
            integrated_description.append(f"{origin_desc_final.replace('\\n', '')}")
            integrated_description.append("")

        # 3. 效果（来自description_key）
        if origin_effects_desc_final and origin_effects_desc_final != description_key:
            integrated_description.append("效果：")
            # 处理多行效果，每行前面添加一个空格
            effect_lines = origin_effects_desc_final.replace('\\n', '\n').split('\n')
            for line in effect_lines:
                if line.strip():
                    integrated_description.append(f" {line.strip()}")
            integrated_description.append("")

        # 4. 修正（来自modifier_effects）
        if origin_effects_list:
            integrated_description.append("修正：")
            for effect in origin_effects_list:
                effect_final = self.process_text_for_output_simple(
                    effect, is_name_or_desc=False)
                effect_final = self.remove_duplicate_text(effect_final)
                effect_final = re.sub(r'£[^£]*£', '', effect_final)
                integrated_description.append(f" -{effect_final.replace('\\n', '')}")
            integrated_description.append("")

        # 5. 惩罚（红色字体）
        punishment_parts = []
        if origin_negative_effects_final and origin_negative_effects_final != negative_description_key:
            punishment_parts.append(origin_negative_effects_final.replace('\\n', ''))

        if non_colonizable_text:
            punishment_parts.append(non_colonizable_text)

        if punishment_parts:
            integrated_description.append("§R惩罚：§!")
            for part in punishment_parts:
                integrated_description.append(f"§R -{part}§!")
            integrated_description.append("")

        # 6. 要求
        if possible_conditions_list:
            integrated_description.append("要求：")
            for condition in possible_conditions_list:
                condition_final = self.remove_duplicate_text(condition)
                condition_final = re.sub(r'£[^£]*£', '', condition_final)

                # 处理"以下条件至少一个满足"的特殊格式
                if condition_final.startswith("以下条件至少一个满足："):
                    integrated_description.append(f" -{condition_final}")
                elif "  " in condition_final:  # 这是子条件（有缩进）
                    # 移除前面的"-"符号，只保留缩进
                    clean_condition = condition_final.lstrip('-').strip()
                    integrated_description.append(f"   {clean_condition}")
                else:
                    integrated_description.append(f" -{condition_final.replace('\\n', '')}")

        return "\n".join(integrated_description)

    def process_text_with_color_markers(self, text, workbook):
        """处理文本中的颜色标记符，返回富文本格式"""
        if not text:
            black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})
            return [(black_format, "")], black_format

        # 创建颜色格式
        red_format = workbook.add_format({'font_color': 'red', 'text_wrap': True})
        green_format = workbook.add_format({'font_color': 'green', 'text_wrap': True})
        blue_format = workbook.add_format({'font_color': 'blue', 'underline': True, 'text_wrap': True})
        orange_format = workbook.add_format({'font_color': 'orange', 'text_wrap': True})
        black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})

        # 如果没有颜色标记，直接返回黑色文本
        if '§R' not in text and '§G' not in text and '§B' not in text and '§H' not in text and '§A' not in text:
            return [(black_format, text)], black_format

        rich_text_parts = []
        current_format = black_format
        current_text = ""

        i = 0
        while i < len(text):
            # 检查是否遇到颜色标记
            if text[i:i + 2] == '§R' and text.find('§!', i + 2) != -1:
                # 保存当前文本
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""

                # 找到红色标记的结束位置
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    # 提取红色内容
                    red_content = text[i + 2:end_pos]
                    rich_text_parts.append((red_format, red_content))
                    i = end_pos + 2  # 跳过§!
                else:
                    i += 2

                current_format = black_format  # 重置为黑色

            elif text[i:i + 2] == '§G' and text.find('§!', i + 2) != -1:
                # 保存当前文本
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""

                # 找到绿色标记的结束位置
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    # 提取绿色内容
                    green_content = text[i + 2:end_pos]
                    rich_text_parts.append((green_format, green_content))
                    i = end_pos + 2  # 跳过§!
                else:
                    i += 2

                current_format = black_format  # 重置为黑色

            elif text[i:i + 2] == '§B' and text.find('§!', i + 2) != -1:
                # 保存当前文本
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""

                # 找到蓝色标记的结束位置
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    # 提取蓝色内容
                    blue_content = text[i + 2:end_pos]
                    rich_text_parts.append((blue_format, blue_content))
                    i = end_pos + 2  # 跳过§!
                else:
                    i += 2

                current_format = black_format  # 重置为黑色

            elif text[i:i + 2] == '§H' and text.find('§!', i + 2) != -1:
                # 保存当前文本
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""

                # 找到橙色标记的结束位置
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    # 提取橙色内容
                    orange_content = text[i + 2:end_pos]
                    rich_text_parts.append((orange_format, orange_content))
                    i = end_pos + 2  # 跳过§!
                else:
                    i += 2

                current_format = black_format  # 重置为黑色

            elif text[i:i + 2] == '§A' and text.find('§!', i + 2) != -1:
                # 保存当前文本
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""

                # 找到蓝色下划线标记的结束位置
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    # 提取蓝色下划线内容
                    blue_underline_content = text[i + 2:end_pos]
                    rich_text_parts.append((blue_format, blue_underline_content))
                    i = end_pos + 2  # 跳过§!
                else:
                    i += 2

                current_format = black_format  # 重置为黑色

            else:
                current_text += text[i]
                i += 1

        # 添加最后一段文本
        if current_text:
            rich_text_parts.append((current_format, current_text))

        # 确保rich_text_parts不为空
        if not rich_text_parts:
            rich_text_parts.append((black_format, ""))

        return rich_text_parts, black_format

    def write_rich_text(self, worksheet, row, col, text, color_info, color_formats, default_format):
        """写入富文本（带颜色）"""
        if not color_info:
            worksheet.write(row, col, text, default_format)
            return

        parts = []
        last_pos = 0

        color_info.sort(key=lambda x: x[0])

        for start, end, color_name, underline in color_info:
            if last_pos < start:
                parts.append(text[last_pos:start])

            colored_text = text[start:end]
            format_key = f"{color_name}_underline" if underline else color_name
            if format_key in color_formats:
                parts.append(color_formats[format_key])
            parts.append(colored_text)

            last_pos = end

        if last_pos < len(text):
            parts.append(text[last_pos:])

        if parts:
            worksheet.write_rich_string(row, col, *parts)

    def write_rich_text_list(self, worksheet, row, col, text, color_info_list, color_formats, default_format):
        """写入列表类型的富文本"""
        if not color_info_list or not any(color_info_list):
            worksheet.write(row, col, text, default_format)
            return

        all_color_info = []
        current_pos = 0

        text_parts = text.split('\n')

        for i, part in enumerate(text_parts):
            if i < len(color_info_list) and color_info_list[i]:
                for start, end, color_name, underline in color_info_list[i]:
                    adjusted_start = current_pos + start
                    adjusted_end = current_pos + end
                    all_color_info.append((adjusted_start, adjusted_end, color_name, underline))

            current_pos += len(part)
            if i < len(text_parts) - 1:
                current_pos += 1

        if all_color_info:
            self.write_rich_text(worksheet, row, col, text, all_color_info, color_formats, default_format)
        else:
            worksheet.write(row, col, text, default_format)

    def generate_excel_output_simple(self):
        """生成国策Excel输出结果"""
        print("正在生成国策Excel输出...")
        excel_file = os.path.join(OUTPUT_DIR, f"{VERSION}国策.xlsx")
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet("国策数据")

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top'
        })

        color_formats = {}
        for color_name in set(COLOR_MAP.values()):
            if color_name == 'blue_underline':
                color_formats[color_name] = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'font_color': 'blue',
                    'underline': True
                })
            else:
                color_formats[color_name] = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'font_color': color_name
                })

        headers = [
            "图片", "国策名称", "国策key", "国策描述", "描述整合",
            "国策说明效果", "国策惩罚", "国策具体效果", "国策条件", "国策类型",
            "内阁席位", "内阁允许的领袖职业", "内阁席位效果"
        ]

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
            worksheet.set_column(col, col, 15)

        # 设置图片列的宽度
        worksheet.set_column(0, 0, 10)

        for row, civic in enumerate(self.civics, 1):
            civic_key = civic['key']

            # 插入图片
            image_path = civic.get('image_path')
            if image_path and os.path.exists(image_path):
                try:
                    # 插入图片到第一列
                    worksheet.insert_image(row, 0, image_path, {
                        'x_offset': 5,
                        'y_offset': 5,
                        'x_scale': 0.8,
                        'y_scale': 0.8
                    })
                except Exception as e:
                    print(f"插入国策图片失败 {image_path}: {e}")

            # 获取国策名称
            civic_name_raw = self.localization_dict.get(civic_key, civic_key)
            civic_name_final = self.process_text_for_output_simple(
                civic_name_raw, is_name_or_desc=True)
            civic_name_final = self.remove_duplicate_text(civic_name_final)
            civic_name_clean = re.sub(r'§[A-Z]|§!', '', civic_name_final)
            civic_name_color_info = []

            civic_key_display = civic_key

            # 获取国策描述
            desc_key = f"{civic_key}_desc"
            civic_desc_raw = self.localization_dict.get(desc_key, "")
            civic_desc_final = self.process_text_for_output_simple(
                civic_desc_raw, is_name_or_desc=True)
            civic_desc_final = self.remove_duplicate_text(civic_desc_final)
            civic_desc_clean, civic_desc_color_info = self.process_color_text_final(civic_desc_final)

            # 获取国策说明效果
            description_key = civic['description_key']
            civic_effects_desc_raw = self.localization_dict.get(description_key, description_key)
            civic_effects_desc_final = self.process_text_for_output_simple(
                civic_effects_desc_raw, is_name_or_desc=False)
            civic_effects_desc_final = self.remove_duplicate_text(civic_effects_desc_final)
            civic_effects_desc_final = re.sub(r'£[^£]*£', '', civic_effects_desc_final)
            civic_effects_desc_clean, civic_effects_desc_color_info = self.process_color_text_final(
                civic_effects_desc_final)

            # 获取国策惩罚
            negative_description_key = civic['negative_description_key']
            civic_negative_effects_raw = self.localization_dict.get(negative_description_key, negative_description_key)
            civic_negative_effects_final = self.process_text_for_output_simple(
                civic_negative_effects_raw, is_name_or_desc=False)
            civic_negative_effects_final = self.remove_duplicate_text(civic_negative_effects_final)
            civic_negative_effects_final = re.sub(r'£[^£]*£', '', civic_negative_effects_final)
            civic_negative_effects_clean, civic_negative_effects_color_info = self.process_color_text_final(
                civic_negative_effects_final)

            # 获取国策具体效果
            civic_effects_list = civic['modifier_effects'] if civic['modifier_effects'] else ["无效果"]
            civic_effects_final_list = []
            civic_effects_color_info_list = []

            for effect in civic_effects_list:
                effect_final = self.process_text_for_output_simple(
                    effect, is_name_or_desc=False)
                effect_final = self.remove_duplicate_text(effect_final)
                effect_clean, effect_color_info = self.process_color_text_final(effect_final)
                civic_effects_final_list.append(effect_clean)
                civic_effects_color_info_list.append(effect_color_info)

            # 获取国策条件
            possible_conditions_list = civic['possible_conditions'] if civic['possible_conditions'] else ["无条件"]
            possible_conditions_final_list = []
            possible_conditions_color_info_list = []

            for condition in possible_conditions_list:
                condition = self.remove_duplicate_text(condition)
                condition_clean, condition_color_info = self.process_color_text_final(condition)
                possible_conditions_final_list.append(condition_clean)
                possible_conditions_color_info_list.append(condition_color_info)

            civic_type = civic['civic_type']

            # 获取内阁信息
            councilor_info = self.councilors_dict.get(civic_key, {})
            councilor_name = ""
            leader_classes = ""
            councilor_effects = ""
            councilor_name_color_info = []
            councilor_effects_color_info = []

            if councilor_info:
                councilor_key = councilor_info['councilor_key']
                councilor_name_raw = self.localization_dict.get(councilor_key, councilor_key)
                councilor_name_final = self.process_text_for_output_simple(
                    councilor_name_raw, is_name_or_desc=False)
                councilor_name_final = self.remove_duplicate_text(councilor_name_final)
                councilor_name_clean, councilor_name_color_info = self.process_color_text_final(councilor_name_final)
                councilor_name = councilor_name_clean

                leader_class_translations = {
                    'official': '行政官',
                    'commander': '指挥官',
                    'scientist': '科学家'
                }

                translated_classes = []
                for cls in councilor_info['leader_classes']:
                    if cls in leader_class_translations:
                        translated_classes.append(leader_class_translations[cls])
                    else:
                        translated_classes.append(cls)

                priority_order = ['行政官', '指挥官', '科学家']
                sorted_classes = []
                for priority in priority_order:
                    if priority in translated_classes:
                        sorted_classes.append(priority)

                for cls in translated_classes:
                    if cls not in sorted_classes:
                        sorted_classes.append(cls)

                leader_classes = ", ".join(sorted_classes)

                councilor_effects_list = councilor_info['modifier_effects'] if councilor_info['modifier_effects'] else [
                    "无效果"]
                councilor_effects_final_list = []
                councilor_effects_color_info = []

                for effect in councilor_effects_list:
                    effect_final = self.process_text_for_output_simple(
                        effect, is_name_or_desc=False)
                    effect_final = self.remove_duplicate_text(effect_final)
                    effect_clean, effect_color_info = self.process_color_text_final(effect_final)
                    councilor_effects_final_list.append(effect_clean)
                    councilor_effects_color_info.append(effect_color_info)

                councilor_effects = "\n".join(councilor_effects_final_list)

            # 生成描述整合
            integrated_description = self.generate_integrated_description(civic)
            integrated_description_clean, integrated_description_color_info = self.process_color_text_final(
                integrated_description)

            # 写入Excel（从第二列开始，因为第一列是图片）
            worksheet.write(row, 1, civic_name_clean.replace('\\n', '\n'), cell_format)
            worksheet.write(row, 2, civic_key_display.replace('\\n', '\n'), cell_format)
            self.write_rich_text(worksheet, row, 3, civic_desc_clean.replace('\\n', '\n'), civic_desc_color_info,
                                 color_formats,
                                 cell_format)
            self.write_rich_text(worksheet, row, 4, integrated_description_clean.replace('\\n', '\n'),
                                 integrated_description_color_info, color_formats,
                                 cell_format)
            self.write_rich_text(worksheet, row, 5, civic_effects_desc_clean.replace('\\n', '\n'),
                                 civic_effects_desc_color_info,
                                 color_formats, cell_format)
            self.write_rich_text(worksheet, row, 6, civic_negative_effects_clean.replace('\\n', '\n'),
                                 civic_negative_effects_color_info, color_formats, cell_format)

            effects_text = "\n".join(civic_effects_final_list).replace('\\n', '\n')
            self.write_rich_text_list(worksheet, row, 7, effects_text, civic_effects_color_info_list, color_formats,
                                      cell_format)

            conditions_text = "\n".join(possible_conditions_final_list).replace('\\n', '\n')
            self.write_rich_text_list(worksheet, row, 8, conditions_text, possible_conditions_color_info_list,
                                      color_formats,
                                      cell_format)

            worksheet.write(row, 9, civic_type.replace('\\n', '\n'), cell_format)
            self.write_rich_text(worksheet, row, 10, councilor_name, councilor_name_color_info, color_formats,
                                 cell_format)
            worksheet.write(row, 11, leader_classes.replace('\\n', '\n'), cell_format)
            self.write_rich_text_list(worksheet, row, 12, councilor_effects, councilor_effects_color_info,
                                      color_formats,
                                      cell_format)

        for row_num in range(1, len(self.civics) + 1):
            worksheet.set_row(row_num, 60)

        workbook.close()
        print(f"国策Excel结果已保存到: {excel_file}")
        return excel_file

    def generate_origin_excel_output(self):
        """生成起源Excel输出结果"""
        print("正在生成起源Excel输出...")
        excel_file = os.path.join(OUTPUT_DIR, f"{VERSION}起源.xlsx")
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet("起源数据")

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top'
        })

        color_formats = {}
        for color_name in set(COLOR_MAP.values()):
            if color_name == 'blue_underline':
                color_formats[color_name] = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'font_color': 'blue',
                    'underline': True
                })
            else:
                color_formats[color_name] = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'font_color': color_name
                })

        headers = [
            "图片", "起源名称", "起源key", "起源描述", "描述整合",
            "起源说明效果", "起源惩罚", "起源具体效果", "起源条件"
        ]

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
            worksheet.set_column(col, col, 20)

        # 设置图片列的宽度
        worksheet.set_column(0, 0, 10)

        for row, origin in enumerate(self.origins, 1):
            origin_key = origin['key']

            # 插入图片
            image_path = origin.get('image_path')
            if image_path and os.path.exists(image_path):
                try:
                    # 插入图片到第一列
                    worksheet.insert_image(row, 0, image_path, {
                        'x_offset': 5,
                        'y_offset': 5,
                        'x_scale': 0.8,
                        'y_scale': 0.8
                    })
                except Exception as e:
                    print(f"插入起源图片失败 {image_path}: {e}")

            # 获取起源名称
            origin_name_raw = self.localization_dict.get(origin_key, origin_key)
            origin_name_final = self.process_text_for_output_simple(
                origin_name_raw, is_name_or_desc=True)
            origin_name_final = self.remove_duplicate_text(origin_name_final)
            origin_name_clean = re.sub(r'§[A-Z]|§!', '', origin_name_final)
            origin_name_color_info = []

            # 获取起源描述
            desc_key = f"{origin_key}_desc"
            origin_desc_raw = self.localization_dict.get(desc_key, "")
            origin_desc_final = self.process_text_for_output_simple(
                origin_desc_raw, is_name_or_desc=True)
            origin_desc_final = self.remove_duplicate_text(origin_desc_final)
            origin_desc_clean, origin_desc_color_info = self.process_color_text_final(origin_desc_final)

            # 获取起源说明效果
            description_key = origin['description_key']
            origin_effects_desc_raw = self.localization_dict.get(description_key, description_key)
            origin_effects_desc_final = self.process_text_for_output_simple(
                origin_effects_desc_raw, is_name_or_desc=False)
            origin_effects_desc_final = self.remove_duplicate_text(origin_effects_desc_final)
            origin_effects_desc_final = re.sub(r'£[^£]*£', '', origin_effects_desc_final)
            origin_effects_desc_clean, origin_effects_desc_color_info = self.process_color_text_final(
                origin_effects_desc_final)

            # 获取起源惩罚
            negative_description_key = origin['negative_description_key']
            origin_negative_effects_raw = self.localization_dict.get(negative_description_key, negative_description_key)
            origin_negative_effects_final = self.process_text_for_output_simple(
                origin_negative_effects_raw, is_name_or_desc=False)
            origin_negative_effects_final = self.remove_duplicate_text(origin_negative_effects_final)
            origin_negative_effects_final = re.sub(r'£[^£]*£', '', origin_negative_effects_final)

            # 合并non_colonizable文本
            non_colonizable_text = origin['non_colonizable_text']
            if non_colonizable_text and origin_negative_effects_final:
                origin_negative_effects_final += f"\n{non_colonizable_text}"
            elif non_colonizable_text:
                origin_negative_effects_final = non_colonizable_text

            origin_negative_effects_clean, origin_negative_effects_color_info = self.process_color_text_final(
                origin_negative_effects_final)

            # 获取起源具体效果
            origin_effects_list = origin['modifier_effects'] if origin['modifier_effects'] else ["无效果"]
            origin_effects_final_list = []
            origin_effects_color_info_list = []

            for effect in origin_effects_list:
                effect_final = self.process_text_for_output_simple(
                    effect, is_name_or_desc=False)
                effect_final = self.remove_duplicate_text(effect_final)
                effect_clean, effect_color_info = self.process_color_text_final(effect_final)
                origin_effects_final_list.append(effect_clean)
                origin_effects_color_info_list.append(effect_color_info)

            # 获取起源条件
            possible_conditions_list = origin['possible_conditions'] if origin['possible_conditions'] else ["无条件"]
            possible_conditions_final_list = []
            possible_conditions_color_info_list = []

            for condition in possible_conditions_list:
                condition = self.remove_duplicate_text(condition)
                condition_clean, condition_color_info = self.process_color_text_final(condition)
                possible_conditions_final_list.append(condition_clean)
                possible_conditions_color_info_list.append(condition_color_info)

            # 生成描述整合
            integrated_description = self.generate_origin_integrated_description(origin)
            integrated_description_clean, integrated_description_color_info = self.process_color_text_final(
                integrated_description)

            # 写入Excel（从第二列开始，因为第一列是图片）
            worksheet.write(row, 1, origin_name_clean.replace('\\n', '\n'), cell_format)
            worksheet.write(row, 2, origin_key.replace('\\n', '\n'), cell_format)
            self.write_rich_text(worksheet, row, 3, origin_desc_clean.replace('\\n', '\n'), origin_desc_color_info,
                                 color_formats,
                                 cell_format)
            self.write_rich_text(worksheet, row, 4, integrated_description_clean.replace('\\n', '\n'),
                                 integrated_description_color_info, color_formats,
                                 cell_format)
            self.write_rich_text(worksheet, row, 5, origin_effects_desc_clean.replace('\\n', '\n'),
                                 origin_effects_desc_color_info,
                                 color_formats, cell_format)
            self.write_rich_text(worksheet, row, 6, origin_negative_effects_clean.replace('\\n', '\n'),
                                 origin_negative_effects_color_info, color_formats, cell_format)

            effects_text = "\n".join(origin_effects_final_list).replace('\\n', '\n')
            self.write_rich_text_list(worksheet, row, 7, effects_text, origin_effects_color_info_list, color_formats,
                                      cell_format)

            conditions_text = "\n".join(possible_conditions_final_list).replace('\\n', '\n')
            self.write_rich_text_list(worksheet, row, 8, conditions_text, possible_conditions_color_info_list,
                                      color_formats,
                                      cell_format)

        for row_num in range(1, len(self.origins) + 1):
            worksheet.set_row(row_num, 80)

        workbook.close()
        print(f"起源Excel结果已保存到: {excel_file}")
        return excel_file

    def save_debug_info(self):
        """保存调试信息"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}国策起源调试信息.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 国策起源调试信息 ===\n\n")

            if self.debug_info['missing_civic_name']:
                f.write("=== 缺失国策名称 ===\n")
                for info in sorted(self.debug_info['missing_civic_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_origin_name']:
                f.write("=== 缺失起源名称 ===\n")
                for info in sorted(self.debug_info['missing_origin_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_tooltip']:
                f.write("=== 缺失效果 ===\n")
                for info in sorted(self.debug_info['missing_tooltip']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_scripted_variable']:
                f.write("=== 缺失脚本变量 ===\n")
                for info in sorted(self.debug_info['missing_scripted_variable']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_images']:
                f.write("=== 缺失图片 ===\n")
                for info in self.debug_info['missing_images']:
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['unresolved_markers']:
                f.write("=== 未解析的标记 ===\n")
                for info in sorted(self.debug_info['unresolved_markers']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 新增：未找到的修正翻译
            if self.debug_info['missing_modifier_translations']:
                f.write("=== 未找到的修正翻译 ===\n")
                for info in sorted(self.debug_info['missing_modifier_translations']):
                    f.write(f"{info}\n")
                f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取国策起源信息...")

        self.parse_localisation_files()
        self.parse_scripted_variables()
        self.parse_scripted_loc_files()

        print("正在预处理本地化字典...")
        self.localization_dict = self.preprocess_localisation_dict_simple()

        # 构建不区分大小写的本地化字典
        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localization_dict.items()}

        self.parse_civic_files()
        self.parse_councilor_files()
        self.parse_origin_files()

        self.generate_excel_output_simple()
        self.generate_origin_excel_output()
        # 不再生成国策TXT文件
        self.save_debug_info()

        print(f"完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = CivicOriginExtractor()
    extractor.run()