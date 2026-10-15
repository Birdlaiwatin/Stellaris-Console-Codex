import os
import re
from pathlib import Path
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 颜色转换字典 import SPECIAL_SUFFIXES
from 修正翻译字典 import MODIFIER_TRANSLATIONS

# 固定£替换符翻译字典
SPECIAL_POUND_TRANSLATIONS = {
    'society': '社会学研究',
    'physics': '物理学研究',
    'engineering': '工程学研究',
    'trade_value': '贸易'
}

# 特质属性优先级映射
TRAIT_ATTRIBUTE_PRIORITY = {
    '宜居特质': 1,
    '有机特质': 2,
    '灵能特质': 3,
    '基因高级特质': 4,
    '可塑特质': 5,
    '变异特质': 6,
    '强夺天工特质': 7,
    '义体特质': 8,
    '机械特质': 9,
    '未开智物种特质': 10,
    '特殊特质': 11
}

# 特质属性排序顺序
TRAIT_ATTRIBUTE_ORDER = [
    '宜居特质',
    '有机特质',
    '灵能特质',
    '基因高级特质',
    '可塑特质',
    '变异特质',
    '强夺天工特质',
    '义体特质',
    '机械特质',
    '未开智物种特质',
    '特殊特质'
]


class StellarisSpeciesTraitParser:
    def __init__(self):
        self.localisation_dict = {}
        self.modifier_localisation_dict = {}
        self.scripted_variables = {}
        self.scripted_loc_dict = {}
        self.species_traits = []
        self.processed_references = set()
        self.untranslated_pound_placeholders = set()
        self.traits_without_effects = []
        self.untranslated_modifiers = set()
        self.trait_image_map = {}

        # 调试信息
        self.debug_info = {
            'missing_tooltip': set(),
            'missing_scripted_variable': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
            'missing_images': []
        }

        # 构建大小写不敏感的本地化字典
        self.insensitive_localization_dict = {}

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(TRAITS_IMAGES_DIR, exist_ok=True)

    def convert_dds_to_png(self, dds_path, png_path):
        """将DDS文件转换为PNG格式"""
        try:
            if not os.path.exists(dds_path):
                return False

            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            img = imageio.imread(dds_path)
            imageio.imwrite(png_path, img)
            return True
        except Exception as e:
            return False

    def format_modifier_value(self, key, value):
        """格式化修正值，进行百分比转换判定"""
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
                    modifier_key.endswith('_uninhabitable') or modifier_key.endswith('_reduction') or
                    modifier_key.endswith('_habitability')):
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
        return f"{sign}{formatted}"

    def get_value_color(self, key, value):
        """根据修正键和值确定颜色"""
        try:
            value_float = float(value)
        except (ValueError, TypeError):
            return 'black'

        if any(key.endswith(suffix) for suffix in SPECIAL_SUFFIXES):
            return 'red' if value_float > 0 else 'green'

        return 'green' if value_float > 0 else 'red'

    def preprocess_localization_text(self, text):
        """预处理本地化文本 - 保留§R和§G颜色标记符，删除其他颜色标记符"""
        if not text:
            return text

        text, _ = self.process_dollar_replacement(text)

        def replace_color_markers(match):
            color_code = match.group(1)
            content = match.group(2)

            if color_code == 'R':
                return f"§R{content}§!"
            elif color_code == 'G':
                return f"§G{content}§!"
            else:
                return content

        text = re.sub(r'§([A-Z])(.*?)§!', replace_color_markers, text)
        return text

    def remove_duplicate_text(self, text):
        """清除连续重复文本，只处理文字，不处理数字"""
        if not text:
            return text

        text = re.sub(r'\s+', '', text)
        pattern = r'([\u4e00-\u9fa5a-zA-Z]{2,10})\1+'

        changed = True
        while changed:
            changed = False
            matches = re.findall(pattern, text)
            if matches:
                for match in matches:
                    text = text.replace(match + match, match)
                    changed = True

        return text

    def clean_whitespace_text(self, text):
        """清理空白文本"""
        if not text:
            return text

        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def process_special_at_replacement(self, text):
        """处理特殊@替换符结构：@[ xxx * 数值]"""
        if not text:
            return text

        pattern = r'@\[\s*(\w+)\s*\*\s*([\d.-]+)\s*\]'

        def replace_special_at(match):
            var_name = match.group(1)
            multiplier = float(match.group(2))

            if var_name in self.scripted_variables:
                result = self.scripted_variables[var_name] * multiplier
                if abs(result - round(result)) < 1e-10:
                    result = round(result)
                if result == int(result):
                    return str(int(result))
                else:
                    return str(result)
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                return match.group(0)

        return re.sub(pattern, replace_special_at, text)

    def process_special_dollar_at_replacement(self, text):
        """处理特殊替换符：$@xxx|任意数值%$"""
        if not text:
            return text

        pattern = r'\$@(\w+)(?:\|([^$]+))?%?\$'

        def replace_special_dollar_at(match):
            var_name = match.group(1)
            default_value = match.group(2)
            has_percent = match.group(0).endswith('%$')
            has_default = default_value is not None

            if var_name in self.scripted_variables:
                result = self.scripted_variables[var_name]
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
                    if abs(result - round(result)) < 1e-10:
                        result = round(result)
                    if result == int(result):
                        return str(int(result))
                    else:
                        return str(result)
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                if has_default:
                    if has_percent:
                        return f"{default_value}%"
                    else:
                        return default_value
                else:
                    return match.group(0)

        return re.sub(pattern, replace_special_dollar_at, text)

    def load_localisation(self):
        """加载本地化文件"""
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
                            self.localisation_dict[current_key] = value_str

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
                    self.localisation_dict[current_key] = value_str

            except Exception as e:
                print(f"加载本地化文件失败: {file_path}, 错误: {e}")

        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localisation_dict.items()}

        for key, value in self.localisation_dict.items():
            if (key.startswith('mod_') or key.startswith('MOD_') or
                    key.startswith('MOD_COUNTRY_') or key.startswith('MOD_FEDERATION_') or
                    key.startswith('MOD_LEADER_') or key.startswith('concept_')):
                self.modifier_localisation_dict[key] = value
                self.modifier_localisation_dict[key.lower()] = value

        print(f"已加载 {len(self.localisation_dict)} 个本地化条目")

    def load_scripted_variables(self):
        """加载脚本变量文件"""
        print("正在加载脚本变量文件...")
        if not os.path.exists(SCRIPTED_VARIABLES_DIR):
            print(f"脚本变量目录不存在: {SCRIPTED_VARIABLES_DIR}")
            return

        for file_name in os.listdir(SCRIPTED_VARIABLES_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(SCRIPTED_VARIABLES_DIR, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                    matches = re.findall(pattern, content)

                    for var_name, var_value in matches:
                        self.scripted_variables[var_name] = float(var_value)

                except Exception as e:
                    print(f"加载脚本变量文件失败: {file_path}, 错误: {e}")

        if os.path.exists(TRAITS_DIR):
            for file_name in os.listdir(TRAITS_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(TRAITS_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        lines = content.split('\n')
                        brace_level = 0

                        for line in lines:
                            line = line.strip()
                            if not line or line.startswith('#'):
                                continue

                            brace_count = 0
                            for char in line:
                                if char == '{':
                                    brace_count += 1
                                elif char == '}':
                                    brace_count -= 1

                            if brace_level == 0:
                                pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                                match = re.match(pattern, line)
                                if match:
                                    var_name, var_value = match.groups()
                                    self.scripted_variables[var_name] = float(var_value)

                    except Exception as e:
                        print(f"加载特质文件中的脚本变量失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.scripted_variables)} 个脚本变量")

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
            return value_str

        if isinstance(value_str, str) and '#' in value_str:
            value_str = value_str.split('#')[0].strip()

        value_str = self.process_special_at_replacement(value_str)
        value_str = self.process_special_dollar_at_replacement(value_str)

        if isinstance(value_str, str) and value_str.startswith('@'):
            var_name = value_str[1:]
            if var_name in self.scripted_variables:
                return self.scripted_variables[var_name]
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                return value_str

        if isinstance(value_str, str) and '$' in value_str:
            value_str, _ = self.process_dollar_replacement(value_str)

        return value_str

    def process_dollar_replacement(self, text, max_depth=5, current_depth=0):
        """处理$xxx$替换符"""
        if current_depth >= max_depth:
            return text, True

        dollar_matches = re.findall(r'\$([^$]+)\$', text)
        if not dollar_matches:
            return text, False

        for match in dollar_matches:
            if match in self.localisation_dict:
                replacement = self.localisation_dict[match]
                replacement, exceeded = self.process_dollar_replacement(replacement, max_depth, current_depth + 1)
                text = text.replace(f'${match}$', replacement)

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
                    if job_key_plural in self.localisation_dict:
                        result = self.localisation_dict[job_key_plural]
                    elif job_key in self.localisation_dict:
                        result = self.localisation_dict[job_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到工作本地化: {job_key} (来自 [{inner_text}])"
                        )
                else:
                    if job_key in self.localisation_dict:
                        result = self.localisation_dict[job_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到工作本地化: {job_key} (来自 [{inner_text}])"
                        )

            elif inner_text.startswith("'") and ',' in inner_text:
                pattern = r"'(.*?)',\s*(.*)"
                comma_match = re.match(pattern, inner_text)
                if comma_match:
                    second_part = comma_match.group(2).strip()
                    if '$' in second_part:
                        second_part, _ = self.process_dollar_replacement(second_part)
                    result = second_part

            elif inner_text.startswith("'") and inner_text.endswith(
                    "'") and ':' not in inner_text and ',' not in inner_text:
                concept_key = inner_text[1:-1]
                if concept_key in self.localisation_dict:
                    result = self.localisation_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localisation_dict:
                        result = self.localisation_dict[building_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到建筑本地化: {building_key} (来自 [{inner_text}])"
                        )

            elif inner_text in self.scripted_loc_dict:
                loc_key = self.scripted_loc_dict[inner_text]
                if loc_key in self.localisation_dict:
                    result = self.localisation_dict[loc_key]
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

        return re.sub(r'\[([^\]]+)\]', replace_brackets, text)

    def process_pound_replacement(self, text):
        """处理£xxx£替换符"""
        if not text:
            return text

        def replace_pounds(match):
            inner_text = match.group(1)
            concept_key = f"concept_{inner_text}"

            if concept_key in self.localisation_dict:
                return self.localisation_dict[concept_key]
            else:
                if inner_text in self.localisation_dict:
                    return self.localisation_dict[inner_text]
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

        for pound_key, translation in SPECIAL_POUND_TRANSLATIONS.items():
            text = text.replace(f'£{pound_key}£', translation)

        return text

    def process_color_tags(self, text):
        """处理颜色标记符：§Rxxx§!为红色，§Gxxx§!为绿色，其他颜色标记符直接删除"""
        if not text:
            return text

        def replace_red(match):
            content = match.group(1)
            return f'§RED§{content}§END§'

        text = re.sub(r'§R([^§]*)§!', replace_red, text)

        def replace_green(match):
            content = match.group(1)
            return f'§GREEN§{content}§END§'

        text = re.sub(r'§G([^§]*)§!', replace_green, text)

        def replace_other(match):
            content = match.group(1)
            return content

        text = re.sub(r'§[A-Z]([^§]*)§!', replace_other, text)

        return text

    def process_localisation_text(self, text, modifier_key=""):
        """通用文本处理方法"""
        if not text:
            return text

        text = self.preprocess_localization_text(text)
        text = self.clean_whitespace_text(text)
        text = self.process_special_dollar_at_replacement(text)
        text = self.process_bracket_replacement(text)
        text = self.process_special_pound_replacement(text)
        text = self.process_pound_replacement(text)
        text = self.process_color_tags(text)
        text = self.remove_duplicate_text(text)
        text = text.replace('\\n', '\n')
        text = self.clean_whitespace_text(text)

        return text

    def get_localisation(self, key, resolve_nested=True):
        """获取本地化文本"""
        if key in self.localisation_dict:
            text = self.localisation_dict[key]
        elif key.lower() in self.localisation_dict:
            text = self.localisation_dict[key.lower()]
        else:
            if key.endswith(('_2', '_3')):
                base_key = key.rsplit('_', 1)[0]
                roman_numeral = 'II' if key.endswith('_2') else 'III'

                base_name = self.get_localisation(base_key, resolve_nested=False)
                if base_name != base_key:
                    if '[' in base_name and ']' in base_name:
                        base_name = self.process_bracket_replacement(base_name)
                    return f"{base_name}{roman_numeral}"

            return key

        if resolve_nested:
            return self.process_localisation_text(text)
        else:
            return text

    def get_modifier_localisation(self, modifier_key):
        """获取修正字段的中文描述"""
        if modifier_key in MODIFIER_TRANSLATIONS:
            return MODIFIER_TRANSLATIONS[modifier_key]

        search_patterns = [
            [f"mod_{modifier_key}", f"mod_{modifier_key}".lower(), f"MOD_{modifier_key}".lower()],
            [f"MOD_{modifier_key}", f"MOD_{modifier_key}".lower()],
            [modifier_key],
            [f"MOD_COUNTRY_{modifier_key}", f"MOD_COUNTRY_{modifier_key}".lower()],
            [f"MOD_FEDERATION_{modifier_key}", f"MOD_FEDERATION_{modifier_key}".lower()],
            [f"MOD_LEADER_{modifier_key}", f"MOD_LEADER_{modifier_key}".lower()],
            [f"edict_{modifier_key}", f"edict_{modifier_key}".lower()],
        ]

        for pattern_group in search_patterns:
            for pattern in pattern_group:
                if pattern in self.modifier_localisation_dict:
                    return self.process_localisation_text(self.modifier_localisation_dict[pattern])

                if pattern in self.localisation_dict:
                    return self.process_localisation_text(self.localisation_dict[pattern])

        if modifier_key.endswith('_mult'):
            base_key = modifier_key[:-5]
            mod_patterns = [
                f"MOD_{base_key}",
                f"MOD_{base_key}".lower(),
                f"mod_{base_key}",
                f"mod_{base_key}".lower()
            ]

            for pattern in mod_patterns:
                if pattern in self.modifier_localisation_dict:
                    return self.process_localisation_text(self.modifier_localisation_dict[pattern])

                if pattern in self.localisation_dict:
                    return self.process_localisation_text(self.localisation_dict[pattern])

        if modifier_key.endswith('_mult'):
            base_key = modifier_key[:-5]
            if base_key in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key])
            if base_key in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key])
            if base_key.lower() in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key.lower()])
            if base_key.lower() in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key.lower()])

        elif modifier_key.endswith('_add'):
            base_key = modifier_key[:-4]
            if base_key in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key])
            if base_key in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key])
            if base_key.lower() in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key.lower()])
            if base_key.lower() in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key.lower()])

        self.untranslated_modifiers.add(modifier_key)
        return modifier_key

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

    def remove_excluded_blocks(self, content):
        """移除需要排除的块"""
        excluded_blocks = ['selectable_weight', 'ai_weight', 'cost', 'random_weight',
                           'slave_cost', 'assembly_score', 'species_possible_add',
                           'species_possible_remove', 'species_potential_add']

        for block_type in excluded_blocks:
            pattern = rf'{block_type}\s*=\s*{{'
            pos = 0
            while pos < len(content):
                match = re.search(pattern, content[pos:])
                if not match:
                    break

                brace_start = pos + match.end() - 1
                block_content, new_pos = self.extract_balanced_braces(content, brace_start)

                if block_content:
                    content = content[:brace_start] + content[new_pos:]
                    pos = brace_start
                else:
                    pos = brace_start + 1

        return content

    def remove_triggered_modifier_blocks(self, content):
        """移除triggered_xxx_modifier块，避免重复提取"""
        triggered_types = [
            'triggered_planet_growth_habitability_modifier',
            'triggered_pop_group_modifier',
            'triggered_leader_modifier',
            'triggered_species_modifier'
        ]

        for triggered_type in triggered_types:
            pattern = rf'{triggered_type}\s*=\s*{{'
            pos = 0
            while pos < len(content):
                match = re.search(pattern, content[pos:])
                if not match:
                    break

                brace_start = pos + match.end() - 1
                triggered_block, new_pos = self.extract_balanced_braces(content, brace_start)
                if triggered_block:
                    content = content[:brace_start] + content[new_pos:]
                    pos = brace_start
                else:
                    pos = brace_start + 1

        return content

    def process_inline_script(self, content):
        """处理inline_script块 - 读取对应脚本文件并替换变量"""
        if not content:
            return content

        # 首先处理简单的 inline_script = "path" 形式
        simple_pattern = r'inline_script\s*=\s*"([^"]+)"'
        pos = 0

        while pos < len(content):
            match = re.search(simple_pattern, content[pos:])
            if not match:
                break

            script_start = pos + match.start()
            script_end = pos + match.end()
            script_path = match.group(1)

            script_file_path = os.path.join(INLINE_SCRIPTS_DIR, f"{script_path}.txt")

            if os.path.exists(script_file_path):
                try:
                    with open(script_file_path, 'r', encoding='utf-8') as f:
                        script_content = f.read()

                    # 直接替换整个 inline_script = "path" 为脚本内容
                    content = content[:script_start] + script_content + content[script_end:]
                    pos = script_start + len(script_content)
                except Exception as e:
                    print(f"加载内联脚本文件失败: {script_file_path}, 错误: {e}")
                    pos = script_end
            else:
                print(f"内联脚本文件不存在: {script_file_path}")
                pos = script_end

        # 然后处理复杂的 inline_script = { ... } 形式（原有逻辑）
        inline_pattern = r'inline_script\s*=\s*\{'
        pos = 0

        while pos < len(content):
            match = re.search(inline_pattern, content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            inline_block, new_pos = self.extract_balanced_braces(content, brace_start)

            if inline_block:
                script_match = re.search(r'script\s*=\s*([^\s\n]+)', inline_block)
                if script_match:
                    script_path = script_match.group(1)
                    script_file_path = os.path.join(INLINE_SCRIPTS_DIR, f"{script_path}.txt")

                    if os.path.exists(script_file_path):
                        try:
                            with open(script_file_path, 'r', encoding='utf-8') as f:
                                script_content = f.read()

                            variables = {}
                            var_pattern = r'([A-Z][A-Z0-9_]*)\s*=\s*([^\n]+)'
                            var_matches = re.findall(var_pattern, inline_block)
                            for var_name, var_value in var_matches:
                                var_value = var_value.strip()
                                if var_value.startswith('"') and var_value.endswith('"'):
                                    var_value = var_value[1:-1]
                                variables[var_name] = var_value

                            for var_name, var_value in variables.items():
                                script_content = script_content.replace(f'${var_name}$', var_value)

                            content = content[:brace_start] + script_content + content[new_pos:]
                            pos = brace_start
                        except Exception as e:
                            print(f"处理内联脚本块失败: {script_file_path}, 错误: {e}")
                            pos = new_pos
                    else:
                        print(f"内联脚本文件不存在: {script_file_path}")
                        pos = new_pos
                else:
                    pos = new_pos
            else:
                pos = brace_start + 1

        return content

    def get_trait_image_path(self, trait_key, trait_content):
        """获取特质图片路径，如果是DDS则自动转换为PNG"""
        icon_path = None

        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', trait_content)
        if icon_match:
            icon_path = icon_match.group(1)
        else:
            icon_match = re.search(r'icon\s*=\s*([^\s{]+)', trait_content)
            if icon_match:
                icon_path = icon_match.group(1)

        if icon_path:
            dds_path = os.path.join(BASE_GAME_DIR, icon_path.replace('/', os.path.sep))
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(TRAITS_IMAGES_DIR, png_filename)

            if os.path.exists(png_path):
                return png_path
            else:
                if os.path.exists(dds_path):
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        self.debug_info['missing_images'].append(f"{trait_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    self.debug_info['missing_images'].append(f"{trait_key} -> DDS文件不存在: {dds_path}")
                    return None

        default_dds_path = os.path.join(TRAITS_GFX_FILE_PATH, f"{trait_key}.dds")
        png_filename = f"{trait_key}.png"
        png_path = os.path.join(TRAITS_IMAGES_DIR, png_filename)

        if os.path.exists(png_path):
            return png_path
        else:
            if os.path.exists(default_dds_path):
                success = self.convert_dds_to_png(default_dds_path, png_path)
                if success:
                    return png_path
                else:
                    self.debug_info['missing_images'].append(f"{trait_key} -> 转换失败: {trait_key}.dds")
                    return None
            else:
                self.debug_info['missing_images'].append(f"{trait_key} -> 默认DDS文件不存在: {default_dds_path}")
                return None

    def parse_trait_file(self, file_path):
        """解析特质文件，通过是否含有leader_class判定是否为物种特质"""
        print(f"正在解析文件: {file_path}")

        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                content = f.read()

            content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)
            content = self.process_inline_script(content)

            pos = 0
            while pos < len(content):
                match = re.search(r'(\w+)\s*=\s*\{', content[pos:])
                if not match:
                    break

                trait_key = match.group(1)
                brace_start = pos + match.end() - 1

                trait_block, new_pos = self.extract_balanced_braces(content, brace_start)
                if trait_block:
                    if 'leader_class' not in trait_block and 'leader_trait_type' not in trait_block:
                        self.parse_species_trait(trait_key, trait_block, os.path.basename(file_path))

                    pos = new_pos
                else:
                    pos = brace_start + 1

        except Exception as e:
            print(f"解析文件失败 {file_path}: {e}")

    def determine_trait_attribute(self, trait_key, tags, category, is_special):
        """确定特质属性"""
        if (trait_key.endswith('_preference') or trait_key.endswith('_terraforming')):
            return '宜居特质'

        if is_special == '是' and 'background' not in tags and category not in ['malleable', 'mutation']:
            return '特殊特质'

        if 'presapient' in tags:
            return '未开智物种特质'

        if 'genetic_ascension' in tags:
            return '基因高级特质'

        category_mapping = {
            'normal': '有机特质',
            'psionic': '灵能特质',
            'malleable': '可塑特质',
            'mutation': '变异特质',
            'robotic': '机械特质',
            'overtuned': '强夺天工特质',
            'cyborg': '义体特质'
        }

        if category in category_mapping:
            return category_mapping[category]

        return '有机特质'

    def parse_triggered_modifiers(self, trait_data, content):
        """解析triggered_xxx_modifier"""
        if trait_data['tooltip']:
            return

        triggered_types = [
            'triggered_planet_growth_habitability_modifier',
            'triggered_pop_group_modifier',
            'triggered_leader_modifier',
            'triggered_species_modifier'
        ]

        for triggered_type in triggered_types:
            pattern = rf'{triggered_type}\s*=\s*{{'
            pos = 0
            while pos < len(content):
                match = re.search(pattern, content[pos:])
                if not match:
                    break

                brace_start = pos + match.end() - 1
                triggered_block, new_pos = self.extract_balanced_braces(content, brace_start)
                if triggered_block:
                    if 'potential' in triggered_block:
                        pos = new_pos
                        continue

                    inner_content = triggered_block[1:-1]
                    effects = self.parse_modifier_block(inner_content)
                    if effects:
                        trait_data['modifier_effects'].extend(effects)

                    pos = new_pos
                else:
                    pos = brace_start + 1

    def parse_species_trait(self, trait_key, content, source_file):
        """解析物种特质"""
        trait_data = {
            'key': trait_key,
            'name': self.get_localisation(trait_key),
            'category': 'normal',
            'tags': [],
            'tooltip': '',
            'modifier_effects': [],
            'is_special': '否',
            'source_file': source_file,
            'icon': None,
            'attribute': '有机特质',
            'has_negative': False
        }

        trait_data['icon'] = self.get_trait_image_path(trait_key, content)

        cleaned_content = self.remove_excluded_blocks(content)
        category_match = re.search(r'category\s*=\s*(\w+)', cleaned_content)
        if category_match:
            trait_data['category'] = category_match.group(1)

        tags_match = re.search(r'\btags\s*=\s*\{([^}]+)\}', content)
        if tags_match:
            tags_content = tags_match.group(1)
            tags = re.findall(r'(\w+)', tags_content)
            trait_data['tags'] = tags

            if 'negative' in tags:
                trait_data['has_negative'] = True

        tooltip_pattern = r'custom_tooltip_with_modifiers\s*=\s*"?([^"\s]+)"?'
        tooltip_matches = re.findall(tooltip_pattern, content)

        if tooltip_matches:
            tooltip_key = tooltip_matches[0]
            tooltip_text = self.get_localisation(tooltip_key)
            trait_data['tooltip'] = tooltip_text

        if not trait_data['tooltip']:
            tooltip_pattern = r'custom_tooltip\s*=\s*"?([^"\s]+)"?'
            tooltip_matches = re.findall(tooltip_pattern, content)

            if tooltip_matches:
                tooltip_key = tooltip_matches[0]
                tooltip_text = self.get_localisation(tooltip_key)
                trait_data['tooltip'] = tooltip_text

        self.parse_triggered_modifiers(trait_data, content)

        content_without_triggered = self.remove_triggered_modifier_blocks(content)
        cleaned_content = self.remove_excluded_blocks(content_without_triggered)
        self.parse_species_modifiers(trait_data, cleaned_content)

        possible_add_match = re.search(r'species_possible_add\s*=\s*\{([^}]*?always\s*=\s*no[^}]*)\}', content,
                                       re.DOTALL)
        if possible_add_match:
            trait_data['is_special'] = '是'

        potential_add_match = re.search(r'species_potential_add\s*=\s*\{([^}]*?always\s*=\s*no[^}]*)\}', content,
                                        re.DOTALL)
        if potential_add_match:
            trait_data['is_special'] = '是'

        trait_data['attribute'] = self.determine_trait_attribute(
            trait_data['key'],
            trait_data['tags'],
            trait_data['category'],
            trait_data['is_special']
        )

        has_effects = bool(trait_data['tooltip'] or trait_data['modifier_effects'])
        if not has_effects:
            self.traits_without_effects.append(trait_key)

        self.species_traits.append(trait_data)
        print(f"已解析物种特质: {trait_key}")

    def parse_species_modifiers(self, trait_data, content):
        """解析物种特质的所有修正效果"""
        pattern = r'modifier\s*=\s*{'
        pos = 0
        processed_positions = {}

        while pos < len(content):
            match = re.search(pattern, content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            modifier_block, new_pos = self.extract_balanced_braces(content, brace_start)

            if modifier_block and brace_start not in processed_positions:
                if modifier_block.startswith('{') and modifier_block.endswith('}'):
                    modifier_block = modifier_block[1:-1]

                effects = self.parse_modifier_block(modifier_block)
                if effects:
                    trait_data['modifier_effects'].extend(effects)
                    processed_positions[brace_start] = True

            pos = new_pos

    def parse_modifier_block(self, modifier_content):
        """解析修正块，只提取实际的修正效果"""
        effects = []

        lines = modifier_content.split('\n') if '\n' in modifier_content else [modifier_content]

        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            if line.startswith('potential'):
                continue

            if any(key in line for key in ['show_if_not_potential', 'not_potential_override_text_key']):
                continue

            if '=' in line and not line.startswith('potential'):
                parts = line.split('=', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    value = parts[1].strip().strip('"')

                    if key in ['potential', 'trigger', 'custom_tooltip', 'custom_tooltip_with_modifiers', 'factor']:
                        continue

                    if key and value:
                        processed_value = self.extract_modifier_value(value)
                        formatted_value = self.format_modifier_value(key, processed_value)
                        value_color = self.get_value_color(key, processed_value)

                        chinese_desc = self.get_modifier_localisation(key)
                        effects.append({
                            'description': chinese_desc,
                            'formatted_value': formatted_value,
                            'value_color': value_color,
                            'key': key,
                            'value': processed_value
                        })

        return effects

    def sort_traits_by_attribute(self):
        """根据特质属性对特质进行排序"""

        def get_attribute_priority(trait):
            attribute = trait.get('attribute', '有机特质')
            return TRAIT_ATTRIBUTE_PRIORITY.get(attribute, 99)

        self.species_traits.sort(key=get_attribute_priority)

    def process_all_traits(self):
        """处理所有特质文件"""
        print("开始处理物种特质文件...")

        self.load_scripted_variables()
        self.load_localisation()
        self.load_scripted_loc()

        trait_files = list(Path(TRAITS_DIR).glob("*.txt"))
        for file_path in trait_files:
            self.parse_trait_file(file_path)

        self.sort_traits_by_attribute()

        print(f"解析完成！找到 {len(self.species_traits)} 个物种特质")
        self.export_results()

    def export_results(self):
        """导出结果到Excel"""
        if not self.species_traits:
            print("没有找到物种特质数据")
            return

        print("正在生成Excel文件...")
        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}物种特质.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('物种特质')

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        light_green_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'text_wrap': True,
            'bg_color': '#E3F2D9'
        })

        light_red_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'text_wrap': True,
            'bg_color': '#FADADE'
        })

        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 30)
        worksheet.set_column('F:F', 80)
        worksheet.set_column('G:G', 15)
        worksheet.set_column('H:H', 15)
        worksheet.set_column('I:I', 30)

        worksheet.set_default_row(40)

        headers = ['图片', '特质名称', '特质KEY', '特质类型', '特质标签', '特质效果', '特殊', '特质属性', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        row = 1
        for trait in self.species_traits:
            cell_format = light_red_format if trait['has_negative'] else light_green_format

            worksheet.write(row, 0, '', cell_format)

            if trait['icon'] and os.path.exists(trait['icon']):
                worksheet.insert_image(row, 0, trait['icon'], {
                    'x_offset': 5,
                    'y_offset': 5,
                    'x_scale': 1,
                    'y_scale': 1,
                    'object_position': 1
                })

            worksheet.write(row, 1, trait['name'], cell_format)
            worksheet.write(row, 2, trait['key'], cell_format)
            worksheet.write(row, 3, trait['category'], cell_format)
            worksheet.write(row, 4, ', '.join(trait['tags']), cell_format)

            combined_effects = []

            tooltip_text = trait['tooltip']
            if tooltip_text:
                combined_effects.append(tooltip_text)

            if trait['modifier_effects']:
                modifier_texts = []
                for effect in trait['modifier_effects']:
                    if effect['value_color'] == 'green':
                        modifier_text = f"{effect['description']}§GREEN§{effect['formatted_value']}§END§"
                    elif effect['value_color'] == 'red':
                        modifier_text = f"{effect['description']}§RED§{effect['formatted_value']}§END§"
                    else:
                        modifier_text = f"{effect['description']}{effect['formatted_value']}"
                    modifier_texts.append(modifier_text)

                if modifier_texts:
                    combined_effects.append("，".join(modifier_texts))

            final_effect_text = "\n".join(combined_effects)

            if final_effect_text and ('§RED§' in final_effect_text or '§GREEN§' in final_effect_text):
                rich_parts = []
                segments = []

                temp_text = final_effect_text
                while temp_text:
                    red_start = temp_text.find('§RED§')
                    green_start = temp_text.find('§GREEN§')

                    next_start = -1
                    if red_start != -1 and green_start != -1:
                        next_start = min(red_start, green_start)
                    elif red_start != -1:
                        next_start = red_start
                    elif green_start != -1:
                        next_start = green_start

                    if next_start == -1:
                        if temp_text:
                            segments.append(('black', temp_text))
                        break

                    if next_start > 0:
                        segments.append(('black', temp_text[:next_start]))

                    if next_start == red_start:
                        end_pos = temp_text.find('§END§', next_start)
                        if end_pos != -1:
                            content = temp_text[next_start + 5:end_pos]
                            segments.append(('red', content))
                            temp_text = temp_text[end_pos + 5:]
                        else:
                            segments.append(('red', temp_text[next_start + 5:]))
                            break
                    elif next_start == green_start:
                        end_pos = temp_text.find('§END§', next_start)
                        if end_pos != -1:
                            content = temp_text[next_start + 7:end_pos]
                            segments.append(('green', content))
                            temp_text = temp_text[end_pos + 5:]
                        else:
                            segments.append(('green', temp_text[next_start + 7:]))
                            break

                for color, text in segments:
                    if text:
                        if color == 'red':
                            red_bg_format = workbook.add_format({
                                'border': 1,
                                'font_color': 'red',
                                'valign': 'vcenter',
                                'text_wrap': True,
                                'bg_color': '#FADADE' if trait['has_negative'] else '#E3F2D9'
                            })
                            rich_parts.append(red_bg_format)
                        elif color == 'green':
                            green_bg_format = workbook.add_format({
                                'border': 1,
                                'font_color': 'green',
                                'valign': 'vcenter',
                                'text_wrap': True,
                                'bg_color': '#FADADE' if trait['has_negative'] else '#E3F2D9'
                            })
                            rich_parts.append(green_bg_format)
                        else:
                            black_bg_format = workbook.add_format({
                                'border': 1,
                                'font_color': 'black',
                                'valign': 'vcenter',
                                'text_wrap': True,
                                'bg_color': '#FADADE' if trait['has_negative'] else '#E3F2D9'
                            })
                            rich_parts.append(black_bg_format)
                        rich_parts.append(text)

                if len(rich_parts) >= 3:
                    try:
                        worksheet.write_rich_string(row, 5, *rich_parts)
                    except Exception:
                        clean_text = re.sub(r'§RED§(.*?)§END§', r'\1', final_effect_text)
                        clean_text = re.sub(r'§GREEN§(.*?)§END§', r'\1', clean_text)
                        worksheet.write(row, 5, clean_text, cell_format)
                else:
                    clean_text = re.sub(r'§RED§(.*?)§END§', r'\1', final_effect_text)
                    clean_text = re.sub(r'§GREEN§(.*?)§END§', r'\1', clean_text)
                    worksheet.write(row, 5, clean_text, cell_format)
            else:
                worksheet.write(row, 5, final_effect_text, cell_format)

            worksheet.write(row, 6, trait['is_special'], cell_format)
            worksheet.write(row, 7, trait['attribute'], cell_format)
            worksheet.write(row, 8, trait['source_file'], cell_format)

            row += 1

        workbook.close()
        print(f"物种特质已导出到: {output_path}")

        self.export_debug_info()

    def export_debug_info(self):
        """导出调试信息到TXT文件"""
        debug_path = os.path.join(OUTPUT_DIR, f"{VERSION}物种特质调试.txt")

        with open(debug_path, 'w', encoding='utf-8') as f:
            f.write("=== 物种特质调试信息 ===\n\n")

            f.write("1. 未找到本地化的修正名称:\n")
            if self.untranslated_modifiers:
                for modifier in sorted(self.untranslated_modifiers):
                    f.write(f"   - {modifier}\n")
                f.write(f"总计: {len(self.untranslated_modifiers)} 个\n")
            else:
                f.write("   无\n")
            f.write("\n")

            f.write("2. 缺失的脚本变量:\n")
            if self.debug_info['missing_scripted_variable']:
                for var in sorted(self.debug_info['missing_scripted_variable']):
                    f.write(f"   - {var}\n")
                f.write(f"总计: {len(self.debug_info['missing_scripted_variable'])} 个\n")
            else:
                f.write("   无\n")
            f.write("\n")

            f.write("3. 缺失的括号替换:\n")
            if self.debug_info['missing_bracket_replacement']:
                for item in sorted(self.debug_info['missing_bracket_replacement']):
                    f.write(f"   - {item}\n")
                f.write(f"总计: {len(self.debug_info['missing_bracket_replacement'])} 个\n")
            else:
                f.write("   无\n")
            f.write("\n")

            f.write("4. 缺失的英镑替换:\n")
            if self.debug_info['missing_pound_replacement']:
                for item in sorted(self.debug_info['missing_pound_replacement']):
                    f.write(f"   - {item}\n")
                f.write(f"总计: {len(self.debug_info['missing_pound_replacement'])} 个\n")
            else:
                f.write("   无\n")
            f.write("\n")

            f.write("5. 没有效果的特质:\n")
            if self.traits_without_effects:
                for trait in sorted(self.traits_without_effects):
                    f.write(f"   - {trait}\n")
                f.write(f"总计: {len(self.traits_without_effects)} 个\n")
            else:
                f.write("   无\n")
            f.write("\n")

            f.write("6. 特质属性统计:\n")
            attribute_count = {}
            for trait in self.species_traits:
                attribute = trait['attribute']
                attribute_count[attribute] = attribute_count.get(attribute, 0) + 1

            for attribute in sorted(attribute_count.keys()):
                f.write(f"   - {attribute}: {attribute_count[attribute]} 个\n")
            f.write(f"总计: {len(self.species_traits)} 个特质\n")

            f.write("7. 缺失的图片:\n")
            if self.debug_info['missing_images']:
                for image_info in self.debug_info['missing_images']:
                    f.write(f"   - {image_info}\n")
                f.write(f"总计: {len(self.debug_info['missing_images'])} 个\n")
            else:
                f.write("   无\n")

        print(f"调试信息已导出到: {debug_path}")


def main():
    parser = StellarisSpeciesTraitParser()
    parser.process_all_traits()
    print("物种特质处理完成！")


if __name__ == "__main__":
    main()