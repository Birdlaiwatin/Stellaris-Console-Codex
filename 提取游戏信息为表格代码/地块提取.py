import os
import re
import xlsxwriter
import math
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 颜色转换字典 import SPECIAL_SUFFIXES
from 修正翻译字典 import MODIFIER_TRANSLATIONS
from 条件翻译字典 import CONDITION_TRANSLATIONS

# 固定£替换符翻译字典
SPECIAL_POUND_TRANSLATIONS = {
    'society': '社会学研究',
    'physics': '物理学研究',
    'engineering': '工程学研究',
    'trade_value': '贸易'
}


class DepositExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_variables_dict = {}
        self.scripted_loc_dict = {}
        self.deposits = []
        self.insensitive_localization_dict = {}

        # 调试信息 - 只保留三种
        self.debug_info = {
            'missing_tooltip': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set()
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(DEPOSITS_IMAGES_DIR, exist_ok=True)

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

    def get_deposit_image_path(self, deposit_key, deposit_content):
        """获取地块图片路径，如果是DDS则自动转换为PNG"""
        icon_path = None

        # 首先从deposit_content中查找icon定义
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', deposit_content)
        if icon_match:
            icon_path = icon_match.group(1)
        else:
            icon_match = re.search(r'icon\s*=\s*([^\s{]+)', deposit_content)
            if icon_match:
                icon_path = icon_match.group(1)

        if icon_path:
            # 构建完整的DDS路径
            dds_path = os.path.join(BASE_GAME_DIR, icon_path.replace('/', os.path.sep))
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(DEPOSITS_IMAGES_DIR, png_filename)

            if os.path.exists(png_path):
                return png_path
            else:
                if os.path.exists(dds_path):
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        return None
                else:
                    return None

        # 如果没有找到icon定义，尝试使用默认路径
        default_dds_path = os.path.join(DEPOSITS_GFX_FILE_PATH, f"{deposit_key}.dds")
        png_filename = f"{deposit_key}.png"
        png_path = os.path.join(DEPOSITS_IMAGES_DIR, png_filename)

        if os.path.exists(png_path):
            return png_path
        else:
            if os.path.exists(default_dds_path):
                success = self.convert_dds_to_png(default_dds_path, png_path)
                if success:
                    return png_path
                else:
                    return None
            else:
                return None

    def process_inline_script(self, content, file_name, deposit_key=None):
        """处理inline_script块 - 读取对应脚本文件并替换变量"""
        if not content:
            return content, 0

        processed_count = 0
        original_content = content

        # 首先处理简单的 inline_script = "path" 形式
        simple_pattern = r'inline_script\s*=\s*(?:"([^"]+)"|([^\s{]+))'
        pos = 0

        while pos < len(content):
            match = re.search(simple_pattern, content[pos:])
            if not match:
                break

            script_start = pos + match.start()
            script_end = pos + match.end()

            # 获取脚本路径，可能是带引号或不带引号的
            script_path = match.group(1) if match.group(1) else match.group(2)

            # 修正路径：将正斜杠替换为系统分隔符，并添加.txt扩展名
            normalized_script_path = script_path.replace('/', os.path.sep)
            script_file_path = os.path.join(INLINE_SCRIPTS_DIR, f"{normalized_script_path}.txt")

            if os.path.exists(script_file_path):
                try:
                    with open(script_file_path, 'r', encoding='utf-8') as f:
                        script_content = f.read()

                    # 直接替换整个 inline_script = "path" 为脚本内容
                    content = content[:script_start] + script_content + content[script_end:]
                    processed_count += 1

                    pos = script_start + len(script_content)
                except Exception as e:
                    pos = script_end
            else:
                pos = script_end

        # 然后处理复杂的 inline_script = { ... } 形式
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
                    # 修正路径：将正斜杠替换为系统分隔符，并添加.txt扩展名
                    normalized_script_path = script_path.replace('/', os.path.sep)
                    script_file_path = os.path.join(INLINE_SCRIPTS_DIR, f"{normalized_script_path}.txt")

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
                            processed_count += 1

                            pos = brace_start
                        except Exception as e:
                            pos = new_pos
                    else:
                        pos = new_pos
                else:
                    pos = new_pos
            else:
                pos = brace_start + 1

        return content, processed_count

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
                    modifier_key.endswith('_uninhabitable') or modifier_key.endswith('_reduction')):
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

        # 特殊后缀：正数为红色，负数为绿色
        if any(key.endswith(suffix) for suffix in SPECIAL_SUFFIXES):
            return 'red' if value_float > 0 else 'green'

        # 一般规则：正数为绿色，负数为红色
        return 'green' if value_float > 0 else 'red'

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

                # 处理浮点数精度问题
                # 如果结果非常接近整数，则使用整数
                if abs(result - round(result)) < 1e-10:
                    result = round(result)

                # 如果结果是整数，返回整数形式
                if result == int(result):
                    return str(int(result))
                else:
                    # 对于非整数，返回原始浮点数
                    return str(result)
            else:
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
                # 如果找不到变量，使用默认值
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

    def load_scripted_variables(self):
        """加载脚本变量"""

        # 从scripted_variables目录加载
        if os.path.exists(SCRIPTED_VARIABLES_DIR):
            for file_name in os.listdir(SCRIPTED_VARIABLES_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(SCRIPTED_VARIABLES_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        # 匹配 @variable = value 模式
                        pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                        matches = re.findall(pattern, content)

                        for var_name, var_value in matches:
                            self.scripted_variables_dict[var_name] = float(var_value)

                    except Exception as e:
                        pass

        # 从deposits目录加载
        if os.path.exists(DEPOSITS_DIR):
            for file_name in os.listdir(DEPOSITS_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(DEPOSITS_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        # 查找整个文件中不在任何块内的@变量定义
                        lines = content.split('\n')
                        brace_level = 0

                        for line in lines:
                            line = line.strip()
                            # 跳过注释和空行
                            if not line or line.startswith('#'):
                                continue

                            # 统计花括号
                            for char in line:
                                if char == '{':
                                    brace_level += 1
                                elif char == '}':
                                    brace_level -= 1

                            # 只在最外层（brace_level == 0）查找@变量定义
                            if brace_level == 0:
                                # 匹配 @variable = value 模式
                                pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                                match = re.match(pattern, line)
                                if match:
                                    var_name, var_value = match.groups()
                                    self.scripted_variables_dict[var_name] = float(var_value)

                    except Exception as e:
                        pass

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

                    # 使用更强大的方法提取defined_text块，处理嵌套花括号
                    self.extract_defined_text_blocks(content)

                except Exception as e:
                    pass

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

    def extract_modifier_value(self, value_str):
        """提取修正值，处理@变量 - 增强版本"""
        # 如果已经是数字，直接返回
        if isinstance(value_str, (int, float)):
            return value_str

        # 忽略#注释
        if isinstance(value_str, str) and '#' in value_str:
            value_str = value_str.split('#')[0].strip()

        # 处理特殊@[ xxx * 数值] 格式
        value_str = self.process_special_at_replacement(value_str)

        # 处理特殊$@xxx|任意数值%$ 格式
        value_str = self.process_special_dollar_at_replacement(value_str)

        if isinstance(value_str, str) and value_str.startswith('@'):
            var_name = value_str[1:]
            if var_name in self.scripted_variables_dict:
                return self.scripted_variables_dict[var_name]
            else:
                return value_str

        # 处理$$替换符
        if isinstance(value_str, str) and '$' in value_str:
            value_str, _ = self.process_dollar_replacement(value_str)

        return value_str

    def get_modifier_description(self, modifier_key):
        """获取修正效果的本地化描述"""
        if modifier_key in MODIFIER_TRANSLATIONS:
            return MODIFIER_TRANSLATIONS[modifier_key]

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

        for pattern in search_patterns:
            if pattern in self.insensitive_localization_dict:
                return self.insensitive_localization_dict[pattern]

        self.debug_info['missing_tooltip'].add(modifier_key)
        return modifier_key

    def get_resource_localization(self, resource_key):
        """获取资源的本地化名称"""
        # 先尝试concept_前缀
        concept_key = f"concept_{resource_key}"
        if concept_key in self.localization_dict:
            return self.localization_dict[concept_key]

        # 再尝试直接查找
        if resource_key in self.localization_dict:
            return self.localization_dict[resource_key]

        # 如果都找不到，返回原始键
        return resource_key

    def process_dollar_replacement(self, text, max_depth=5, current_depth=0):
        """处理$xxx$替换符"""
        if current_depth >= max_depth:
            return text, True  # 返回文本和是否超过最大深度

        # 查找所有$xxx$替换符
        dollar_matches = re.findall(r'\$([^$]+)\$', text)
        if not dollar_matches:
            return text, False

        for match in dollar_matches:
            # 查找替换符对应的本地化
            if match in self.localization_dict:
                replacement = self.localization_dict[match]
                # 递归处理替换后的文本
                replacement, exceeded = self.process_dollar_replacement(replacement, max_depth, current_depth + 1)
                text = text.replace(f'${match}$', replacement)
            else:
                # 如果找不到对应的本地化，保留原样
                pass

        # 检查是否还有$xxx$，如果有且达到最大深度，则标记为超过最大深度
        if current_depth >= max_depth - 1 and re.search(r'\$[^$]+\$', text):
            return text, True

        return text, False

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符"""
        if not text or current_depth >= max_depth:
            return text

        # 处理 [GetIcon] 类型的替换
        text = re.sub(r'\[[^\]]*\.GetIcon\]', '', text)

        # 处理特殊格式的[]替换符
        def replace_brackets(match):
            inner_text = match.group(1)
            result = inner_text

            # 1. 优先级最高：处理 [xxx.GetNamePlural] 或 [xxx.GetName]
            name_match = re.match(r'(\w+)\.GetName(Plural)?', inner_text)
            if name_match:
                prefix = name_match.group(1)
                is_plural = name_match.group(2) is not None

                # 尝试查找 job_前缀的本地化
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

            # 2. 优先级次高：处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' not in inner_text:
                concept_key = inner_text[1:-1]  # 去掉单引号
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            # 3. 优先级再次：处理 ['xxx:yyy'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)  # 去掉单引号并按冒号分割
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localization_dict:
                        result = self.localization_dict[building_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到建筑本地化: {building_key} (来自 [{inner_text}])"
                        )

            # 4. 常规[]替换符处理 - 查找scripted_loc_dict
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

            # 递归处理替换后的文本中的[]替换符
            if current_depth < max_depth - 1:
                result = self.process_bracket_replacement(result, max_depth, current_depth + 1)

            return result

        text = re.sub(r'\[([^\]]+)\]', replace_brackets, text)

        return text

    def process_pound_replacement(self, text):
        """处理£xxx£替换符 - 修复版本，替换后继续处理其他替换符"""
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

        # 替换后继续处理其他替换符
        text, _ = self.process_dollar_replacement(text)
        text = self.process_bracket_replacement(text)

        return text

    def process_special_pound_replacement(self, text):
        """处理特殊的£xxx£替换符 - 固定翻译"""
        if not text:
            return text

        # 固定清除£blocker£字段
        text = text.replace('£blocker£', '')

        # 固定替换特定£xxx£为指定翻译
        for pound_key, translation in SPECIAL_POUND_TRANSLATIONS.items():
            text = text.replace(f'£{pound_key}£', translation)

        return text

    def process_deposit_name(self, text, modifier_key=""):
        """通用文本处理方法 - 替换原来的process_localization_text"""
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

        # 5. 处理特殊£xxx£替换（固定翻译）
        text = self.process_special_pound_replacement(text)

        # 6. 处理常规£xxx£替换
        text = self.process_pound_replacement(text)

        # 7. 清除重复文本
        text = self.remove_duplicate_text(text)

        # 8. 处理换行符
        text = text.replace('\\n', '\n')

        # 9. 最终清理空白
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

    def translate_condition(self, key, value, in_not_block=False):
        """翻译条件 - 修复版本，处理NOT块中的具体值条件"""
        # 检查是否在条件翻译字典中
        if key in CONDITION_TRANSLATIONS:
            base_translation = CONDITION_TRANSLATIONS[key]
        else:
            # 如果不在字典中，尝试获取本地化翻译
            base_translation = self.get_localization_value(key)
            if base_translation == key:
                # 如果都找不到，使用原始键
                base_translation = key

        # 处理布尔条件（yes/no）- 简化逻辑
        if value in ['yes', 'no']:
            if value == 'yes':
                return base_translation
            else:  # value == 'no'
                return f"非{base_translation}"

        # 处理具体值条件
        else:
            # 尝试本地化值
            value_localized = self.get_localization_value(value)
            if value_localized == value:
                value_localized = value  # 保持原样

            # 根据是否在NOT块中决定翻译
            if in_not_block:
                return f"没有{base_translation}{value_localized}"
            else:
                return f"拥有{base_translation}{value_localized}"

    def get_localization_value(self, value):
        """获取值的本地化翻译"""
        # 尝试直接查找本地化
        if value in self.localization_dict:
            return self.localization_dict[value]

        # 尝试查找concept_前缀
        concept_key = f"concept_{value}"
        if concept_key in self.localization_dict:
            return self.localization_dict[concept_key]

        # 如果都找不到，返回原始值
        return value

    def extract_conditions_from_potential(self, potential_block):
        """从potential块中提取条件 - 统一版本"""
        conditions = []

        # 移除最外层花括号
        if potential_block.startswith('{') and potential_block.endswith('}'):
            potential_content = potential_block[1:-1].strip()
        else:
            potential_content = potential_block

        # 移除exists = owner条件
        potential_content = re.sub(r'exists\s*=\s*owner', '', potential_content)

        # 使用递归解析嵌套条件
        self.parse_nested_conditions(potential_content, conditions)

        return conditions

    def extract_trigger_conditions(self, trigger_block):
        """从trigger块中提取条件 - 统一版本"""
        conditions = []

        # 移除最外层花括号
        if trigger_block.startswith('{') and trigger_block.endswith('}'):
            trigger_content = trigger_block[1:-1].strip()
        else:
            trigger_content = trigger_block

        # 使用递归解析嵌套条件
        self.parse_nested_conditions(trigger_content, conditions)

        return conditions

    def process_single_resource_block(self, resource_block, deposit_data, block_id, resource_type='produces'):
        """处理单个资源块（produces或upkeep）- 修复版本，更彻底地移除trigger块"""
        # 提取trigger条件（如果有）
        conditions = []
        trigger_match = re.search(r'trigger\s*=\s*\{', resource_block)

        if trigger_match:
            brace_start = trigger_match.end() - 1
            trigger_block, _ = self.extract_balanced_braces(resource_block, brace_start)

            if trigger_block:
                conditions = self.extract_trigger_conditions(trigger_block)

        # 更彻底地移除trigger块 - 使用平衡花括号确保完全移除
        cleaned_content = resource_block
        if trigger_match:
            # 使用平衡花括号提取trigger块，然后从原内容中移除
            brace_start = trigger_match.end() - 1
            trigger_block, trigger_end_pos = self.extract_balanced_braces(resource_block, brace_start)
            if trigger_block:
                # 从资源块中移除整个trigger块
                cleaned_content = resource_block[:trigger_match.start()] + resource_block[trigger_end_pos:]

        # 提取资源键值对 - 修复正则表达式以匹配脚本变量
        resources = []
        lines = cleaned_content.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 跳过明显不是资源键的行
            if any(keyword in line for keyword in ['trigger', 'planet', 'owner', 'pop', 'OR', 'AND', 'NOT']):
                continue

            # 修复：修改正则表达式以匹配脚本变量 (@xxx)
            match = re.search(r'(\w+)\s*=\s*([^\s#]+)', line)  # 匹配非空白非注释字符
            if match:
                resource_key = match.group(1)
                resource_value = match.group(2)

                # 额外检查：跳过明显不是资源的键
                if resource_key in ['trigger', 'planet', 'owner', 'pop']:
                    continue

                # 处理资源值（包括@替换符和$$替换符）
                processed_value = self.extract_modifier_value(resource_value)

                # 如果是upkeep，值取负数
                if resource_type == 'upkeep':
                    try:
                        processed_value = -float(processed_value)
                    except (ValueError, TypeError):
                        # 如果无法转换为数字，保持原样
                        pass

                # 获取资源名称的本地化
                resource_name = self.get_resource_localization(resource_key)
                resource_name = self.process_deposit_name(resource_name, resource_key)

                # 格式化资源值
                formatted_value = self.format_modifier_value(resource_key, processed_value)

                # 资源值颜色：produces为绿色，upkeep为红色
                value_color = 'green' if resource_type == 'produces' else 'red'

                resources.append({
                    'key': resource_key,
                    'value': processed_value,
                    'formatted_value': formatted_value,
                    'description': resource_name,
                    'value_color': value_color
                })

        # 合并同一个资源块内的资源，共享条件
        if resources and conditions:
            # 构建条件前缀
            condition_text = f"（{'，'.join(conditions)}）"

            # 创建合并后的资源项
            merged_description = f"{condition_text}"

            # 添加合并后的资源项
            deposit_data['modifiers'].append({
                'key': f'{resource_type}_merged_{block_id}',
                'value': '',
                'formatted_value': '',
                'description': merged_description,
                'value_color': value_color,
                'merged_modifiers': resources  # 存储合并的资源信息，但保持原始结构
            })
        elif resources:
            # 如果没有条件，直接添加资源
            for resource in resources:
                deposit_data['modifiers'].append(resource)

    def parse_nested_conditions(self, content, conditions, logic_context=None, parent_key=None):
        """递归解析嵌套条件 - 修复版本，正确处理NOT块中的具体值条件"""
        if not content.strip():
            return

        # 移除最外层花括号（如果有）
        if content.startswith('{') and content.endswith('}'):
            content = content[1:-1].strip()

        pos = 0
        content_length = len(content)

        # 检查是否在NOT块中
        in_not_block = (logic_context == 'NOT')

        while pos < content_length:
            # 跳过空白字符
            while pos < content_length and content[pos].isspace():
                pos += 1

            if pos >= content_length:
                break

            # 查找逻辑块
            logic_type, logic_content, new_pos = self.extract_logic_block(content, pos)
            if logic_type:
                self.process_logic_block(logic_type, logic_content, conditions, parent_key)
                pos = new_pos
                continue

            # 查找内联块（如 owner = { condition }）
            inline_match = re.match(r'(\w+)\s*=\s*\{', content[pos:])
            if inline_match:
                key = inline_match.group(1)

                # 找到开花括号的位置
                brace_start = pos + content[pos:].find('{')

                # 提取完整的内联块
                full_block, end_pos = self.extract_balanced_braces(content, brace_start)

                if full_block:
                    # 提取内部内容（去掉外层花括号）
                    inner_content = full_block[full_block.find('{') + 1:full_block.rfind('}')].strip()

                    # 跳过一些不需要的键
                    if key in ['exists', 'FROM']:
                        pos = end_pos
                        continue

                    # 处理不同的内联块类型
                    if key == 'owner':
                        owner_conditions = []
                        self.parse_nested_conditions(inner_content, owner_conditions, logic_context, key)
                        for condition in owner_conditions:
                            conditions.append(f"{condition}")
                        pos = end_pos
                        continue

                    elif key == 'planet':
                        planet_conditions = []
                        self.parse_nested_conditions(inner_content, planet_conditions, logic_context, key)
                        for condition in planet_conditions:
                            if not condition.startswith("所在星球："):
                                conditions.append(f"所在星球：{condition}")
                            else:
                                conditions.append(condition)
                        pos = end_pos
                        continue

                    else:
                        # 对于其他块，递归处理
                        other_conditions = []
                        self.parse_nested_conditions(inner_content, other_conditions, logic_context, key)
                        conditions.extend(other_conditions)
                        pos = end_pos
                        continue

            # 查找简单的键值对条件
            if pos < content_length:
                # 找到下一个换行或分号
                line_end = content.find('\n', pos)
                if line_end == -1:
                    line_end = content_length

                line = content[pos:line_end].strip()

                if '=' in line and '{' not in line and '}' not in line:
                    parts = line.split('=', 1)
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()

                        # 跳过一些不需要的条件
                        if key not in ['exists', 'FROM', 'always']:
                            # 传递是否在NOT块中的信息
                            condition_text = self.translate_condition(key, value, in_not_block)

                            # 根据上下文添加前缀
                            if parent_key == 'planet' and not condition_text.startswith("所在星球："):
                                condition_text = f"所在星球：{condition_text}"

                            conditions.append(condition_text)

                pos = line_end + 1

    def extract_logic_block(self, content, current_pos=0):
        """提取逻辑块（NOR, NAND, NOT, OR, AND）- 修复版本，处理嵌套块"""
        logic_types = ['NOR', 'NAND', 'NOT', 'OR', 'AND']

        # 从当前位置开始查找逻辑块
        remaining_content = content[current_pos:]

        for logic_type in logic_types:
            pattern = rf'{logic_type}\s*=\s*{{'
            match = re.search(pattern, remaining_content)
            if match:
                # 计算在整个内容中的实际位置
                brace_start = current_pos + match.end() - 1

                # 使用extract_balanced_braces提取完整的逻辑块
                logic_block, end_pos = self.extract_balanced_braces(content, brace_start)

                if logic_block:
                    # 提取逻辑块内容（去掉外层花括号）
                    content_start = logic_block.find('{') + 1
                    content_end = logic_block.rfind('}')
                    logic_content = logic_block[content_start:content_end].strip()
                    return logic_type, logic_content, end_pos
                else:
                    # 继续查找其他逻辑类型
                    continue

        return None, None, current_pos

    def process_logic_block(self, logic_type, logic_content, conditions, parent_key=None):
        """处理逻辑块 - 修复版本"""
        inner_conditions = []

        # 传递逻辑类型作为上下文
        self.parse_nested_conditions(logic_content, inner_conditions, logic_type, parent_key)

        if not inner_conditions:
            return

        # 根据逻辑类型添加前缀
        if logic_type == 'NOT':
            # NOT块只能有一个条件
            if len(inner_conditions) == 1:
                # 对于NOT块，我们不需要再添加"非"前缀，因为具体值条件已经在translate_condition中处理了
                conditions.append(inner_conditions[0])
            else:
                # 如果有多个条件，取第一个并记录警告
                conditions.append(inner_conditions[0])
        elif logic_type == 'AND':
            conditions.append(f"同时满足（{'，'.join(inner_conditions)}）")
        elif logic_type == 'OR':
            conditions.append(f"满足以下任意（{'，'.join(inner_conditions)}）")
        elif logic_type == 'NOR':
            conditions.append(f"不满足以下任意（{'，'.join(inner_conditions)}）")
        elif logic_type == 'NAND':
            conditions.append(f"同时不满足（{'，'.join(inner_conditions)}）")

    def extract_triggered_planet_modifier_conditions(self, modifier_block):
        """提取triggered_planet_modifier中的条件 - 修复版本"""
        conditions = []

        # 查找potential块
        potential_match = re.search(r'potential\s*=\s*\{', modifier_block)
        if potential_match:
            brace_start = potential_match.end() - 1
            potential_block, _ = self.extract_balanced_braces(modifier_block, brace_start)

            if potential_block:
                # 提取条件内容
                conditions = self.extract_conditions_from_potential(potential_block)

        return conditions

    def process_per_pop_modifier(self, key, value, description):
        """处理_per_pop和_per_crime后缀的修正，计算每多少人口/犯罪度提供一个岗位"""
        try:
            value_float = float(value)
            if value_float == 0:
                return description, value, 'black'

            # 计算每多少人口/犯罪度提供一个岗位
            per_value = 1 / value_float

            # 向下取整到整数位
            per_value = math.floor(per_value)

            # 确定修正类型和对应的描述
            if key.endswith('_per_pop'):
                per_type = "人口"
            elif key.endswith('_per_crime'):
                per_type = "点犯罪度"
            else:
                return description, value, 'black'

            # 从原始描述中提取岗位名称
            # 处理描述中的替换符
            processed_description = self.process_deposit_name(description)

            # 尝试从描述中提取岗位名称
            job_match = re.search(r'提供\+1个([^岗]*)岗位', processed_description)
            if job_match:
                job_name = job_match.group(1)
            else:
                # 如果无法提取，尝试从key中提取
                if key.endswith('_per_pop'):
                    job_match = re.match(r'job_(\w+)_per_pop', key)
                elif key.endswith('_per_crime'):
                    job_match = re.match(r'job_(\w+)_per_crime', key)

                if job_match:
                    job_key = f"job_{job_match.group(1)}"
                    if job_key in self.localization_dict:
                        job_name = self.localization_dict[job_key]
                        job_name = self.process_deposit_name(job_name)
                    else:
                        job_name = job_match.group(1)
                else:
                    job_name = "岗位"

            # 构建新的描述 - 只显示岗位，不显示+1
            new_description = f"每{per_value}{per_type}提供+1个{job_name}岗位"

            return new_description, "", 'green'  # 空值和绿色，因为数字已经在描述中

        except (ValueError, TypeError, ZeroDivisionError):
            return description, value, 'black'

    def parse_deposit_file(self, file_path):
        """解析单个地块文件"""
        file_name = os.path.basename(file_path)

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 处理inline_script块
            original_content = content
            content, inline_script_count = self.process_inline_script(content, file_name)

            # 使用更强大的方法提取完整的地块定义
            deposits = self.extract_complete_blocks(content)

            for deposit_key, deposit_content in deposits:
                # 为特定地块处理inline_script
                if deposit_key in ['d_mountain_range', 'd_nanotech_devastation', 'd_space_junk_hazards',
                                   'd_arcane_generator']:
                    # 重新处理该地块的inline_script以获得详细调试信息
                    deposit_content, _ = self.process_inline_script(deposit_content, file_name, deposit_key)
                self.parse_single_deposit(deposit_key, deposit_content, file_name)

        except Exception as e:
            pass

    def extract_complete_blocks(self, content):
        """提取完整的地块定义块，正确处理嵌套花括号"""
        blocks = []
        lines = content.split('\n')

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 寻找地块定义开始
            if line and '=' in line and '{' in line:
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

    def parse_single_deposit(self, deposit_key, deposit_content, source_file):
        """解析单个地块定义 - 修复版本"""
        deposit_data = {
            'key': deposit_key,
            'name': '',
            'main_category': '无主类型',
            'sub_category': '无副类型',
            'tooltip': '',
            'modifiers': [],
            'triggered_modifier_conditions': [],
            'source': source_file,
            'icon': None
        }

        # 获取地块图片
        deposit_data['icon'] = self.get_deposit_image_path(deposit_key, deposit_content)

        # 获取地块名称 - 使用专门的地块名称处理
        raw_name = self.localization_dict.get(deposit_key, deposit_key)
        deposit_data['name'] = self.process_deposit_name(raw_name)

        # 提取主类型 - 最外层的category（不在任何嵌套块内）
        lines = deposit_content.split('\n')
        brace_level = 0

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 统计花括号
            for char in line:
                if char == '{':
                    brace_level += 1
                elif char == '}':
                    brace_level -= 1

            # 只在最外层（brace_level == 0）查找category
            if brace_level == 0 and 'category' in line and '=' in line:
                # 确保这是一个完整的键值对，不是嵌套块的一部分
                if '{' not in line and '}' not in line:
                    match = re.search(r'category\s*=\s*(\w+)', line)
                    if match:
                        deposit_data['main_category'] = match.group(1)
                        break

        # 提取最外层的custom_tooltip（忽略嵌套的）
        tooltip_match = re.search(r'^\s*custom_tooltip\s*=\s*(\w+)', deposit_content, re.MULTILINE)
        if tooltip_match:
            tooltip_key = tooltip_match.group(1)
            raw_tooltip = self.localization_dict.get(tooltip_key, tooltip_key)
            deposit_data['tooltip'] = self.process_deposit_name(raw_tooltip, deposit_key)

        # 使用全新的方法提取resources和produces - 取消副类型限制
        self.extract_resources_and_produces_new(deposit_content, deposit_data)

        # 提取planet_modifier - 现在只提取普通的planet_modifier，不包含triggered_planet_modifier
        self.extract_planet_modifiers(deposit_content, deposit_data)

        # 提取country_modifier - 新增：提取国家修正
        self.extract_country_modifiers(deposit_content, deposit_data)

        # 提取triggered_planet_modifier - 修复版本，现在会合并条件修正
        self.extract_triggered_planet_modifiers(deposit_content, deposit_data)

        # 提取habitat_modifier
        self.extract_habitat_modifiers(deposit_content, deposit_data)

        self.deposits.append(deposit_data)

    def extract_planet_modifiers(self, deposit_content, deposit_data):
        """提取planet_modifier - 修复版本，避免匹配triggered_planet_modifier"""
        # 使用负向前瞻来避免匹配triggered_planet_modifier
        pattern = r'(?<!triggered_)planet_modifier\s*=\s*\{'
        pos = 0

        while pos < len(deposit_content):
            match = re.search(pattern, deposit_content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            modifier_block, new_pos = self.extract_balanced_braces(deposit_content, brace_start)

            if modifier_block:
                # 提取修正效果
                modifiers = self.extract_modifiers_from_block_content(modifier_block, deposit_data['key'])
                deposit_data['modifiers'].extend(modifiers)

            pos = new_pos

    def extract_country_modifiers(self, deposit_content, deposit_data):
        """提取country_modifier - 新增方法，提取规则与planet_modifier一致"""
        pattern = r'country_modifier\s*=\s*\{'
        pos = 0

        while pos < len(deposit_content):
            match = re.search(pattern, deposit_content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            modifier_block, new_pos = self.extract_balanced_braces(deposit_content, brace_start)

            if modifier_block:
                # 提取修正效果
                modifiers = self.extract_modifiers_from_block_content(modifier_block, deposit_data['key'])
                deposit_data['modifiers'].extend(modifiers)

            pos = new_pos

    def extract_habitat_modifiers(self, deposit_content, deposit_data):
        """提取habitat_modifier - 修复版本"""
        pattern = r'habitat_modifier\s*=\s*\{'
        pos = 0

        while pos < len(deposit_content):
            match = re.search(pattern, deposit_content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            modifier_block, new_pos = self.extract_balanced_braces(deposit_content, brace_start)

            if modifier_block:
                # 提取修正效果
                modifiers = self.extract_modifiers_from_block_content(modifier_block, deposit_data['key'])
                deposit_data['modifiers'].extend(modifiers)

            pos = new_pos

    def extract_triggered_planet_modifiers(self, deposit_content, deposit_data):
        """提取triggered_planet_modifier及其条件 - 修复版本，合并同一个块内的修正"""
        pattern = r'triggered_planet_modifier\s*=\s*\{'
        pos = 0

        while pos < len(deposit_content):
            match = re.search(pattern, deposit_content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            modifier_block, new_pos = self.extract_balanced_braces(deposit_content, brace_start)

            if modifier_block:
                # 提取条件
                conditions = self.extract_triggered_planet_modifier_conditions(modifier_block)

                # 移除potential块，提取剩余的修正
                cleaned_content = re.sub(r'potential\s*=\s*\{[^}]*\}', '', modifier_block)

                # 检查是否有嵌套的modifier块
                nested_modifier_match = re.search(r'modifier\s*=\s*\{', cleaned_content)
                if nested_modifier_match:
                    # 有嵌套的modifier块，提取其中的修正
                    nested_brace_start = nested_modifier_match.end() - 1
                    nested_modifier_block, _ = self.extract_balanced_braces(cleaned_content, nested_brace_start)

                    if nested_modifier_block:
                        modifiers = self.extract_modifiers_from_block_content(nested_modifier_block,
                                                                              deposit_data['key'])
                    else:
                        modifiers = []
                else:
                    # 没有嵌套的modifier块，直接提取修正
                    modifiers = self.extract_modifiers_from_block_content(cleaned_content, deposit_data['key'])

                # 合并同一个triggered_planet_modifier内的修正，共享条件
                if modifiers and conditions:
                    # 构建条件前缀
                    condition_text = f"（{'，'.join(conditions)}）"

                    # 创建合并后的修正项
                    merged_description = f"{condition_text}"

                    # 添加合并后的修正项
                    deposit_data['modifiers'].append({
                        'key': 'triggered_merged',
                        'value': '',
                        'formatted_value': '',
                        'description': merged_description,
                        'value_color': 'black',
                        'merged_modifiers': modifiers  # 存储合并的修正信息，但保持原始结构
                    })
                elif modifiers:
                    # 如果没有条件，直接添加修正
                    for modifier in modifiers:
                        deposit_data['modifiers'].append(modifier)

                # 存储条件用于调试
                if conditions:
                    deposit_data['triggered_modifier_conditions'].extend(conditions)

            pos = new_pos

    def extract_modifiers_from_block_content(self, block_content, deposit_key):
        """从块内容中提取修正效果 - 通用方法，修复版本"""
        modifiers = []

        # 分割成行处理
        lines = block_content.split('\n')

        for line in lines:
            line = line.strip()
            if not line or '=' not in line:
                continue

            # 忽略注释
            if '#' in line:
                line = line.split('#')[0].strip()

            # 跳过包含花括号的行（这些是嵌套块）
            if '{' in line or '}' in line:
                continue

            # 跳过明显的条件语句
            if any(keyword in line for keyword in
                   ['potential', 'trigger', 'owner', 'exists', 'OR', 'AND', 'NOT', 'NOR', 'has_', 'is_', 'can_']):
                continue

            # 分割键值对
            parts = line.split('=', 1)
            if len(parts) != 2:
                continue

            key = parts[0].strip()
            value = parts[1].strip()

            # 跳过不需要的键
            if key in ['potential', 'icon', 'icon_frame', 'trigger', 'owner', 'exists']:
                continue

            # 跳过以条件关键词开头的键
            if any(key.startswith(prefix) for prefix in ['has_', 'is_', 'can_']):
                continue

            # 处理修正值（包括@替换符和$$替换符）
            processed_value = self.extract_modifier_value(value)

            effect_description = self.get_modifier_description(key)

            # 如果描述与键相同，说明没有找到翻译，可能是条件语句
            if effect_description == key:
                continue

            effect_description = self.process_deposit_name(effect_description, key)

            # 特殊处理_per_pop和_per_crime后缀的修正
            if key.endswith('_per_pop') or key.endswith('_per_crime'):
                new_description, new_value, value_color = self.process_per_pop_modifier(key, processed_value,
                                                                                        effect_description)
                effect_description = new_description
                formatted_value = new_value
            else:
                # 格式化修正值
                formatted_value = self.format_modifier_value(key, processed_value)
                value_color = self.get_value_color(key, processed_value)

            modifiers.append({
                'key': key,
                'value': processed_value,
                'formatted_value': formatted_value,
                'description': effect_description,
                'value_color': value_color
            })

        return modifiers

    def extract_resources_and_produces_new(self, deposit_content, deposit_data):
        """使用全新的方法提取resources和produces - 取消副类型限制，添加upkeep提取"""

        # 首先查找resources块
        resources_pattern = r'resources\s*=\s*\{'
        resources_match = re.search(resources_pattern, deposit_content)

        if not resources_match:
            return

        brace_start = resources_match.end() - 1
        resources_block, _ = self.extract_balanced_braces(deposit_content, brace_start)

        if not resources_block:
            return

        # 在resources块内查找category
        category_match = re.search(r'category\s*=\s*(\w+)', resources_block)
        if category_match:
            deposit_data['sub_category'] = category_match.group(1)

        # 取消副类型限制，直接提取所有的produces块
        produces_blocks = self.extract_all_resource_blocks(resources_block, deposit_data['key'], 'produces')

        # 提取所有的upkeep块
        upkeep_blocks = self.extract_all_resource_blocks(resources_block, deposit_data['key'], 'upkeep')

        # 处理每个produces块
        for i, produces_block in enumerate(produces_blocks):
            self.process_single_resource_block(produces_block, deposit_data, f"produces_{i + 1}", 'produces')

        # 处理每个upkeep块
        for i, upkeep_block in enumerate(upkeep_blocks):
            self.process_single_resource_block(upkeep_block, deposit_data, f"upkeep_{i + 1}", 'upkeep')

    def extract_all_resource_blocks(self, resources_block, deposit_key, block_type):
        """提取resources块中的所有指定类型块（produces或upkeep）"""
        blocks = []
        pos = 0

        while pos < len(resources_block):
            pattern = rf'{block_type}\s*=\s*{{'
            match = re.search(pattern, resources_block[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            resource_block, new_pos = self.extract_balanced_braces(resources_block, brace_start)

            if resource_block:
                blocks.append(resource_block)

            pos = new_pos

        return blocks

    def parse_all_deposits(self):
        """解析所有地块文件"""

        if not os.path.exists(DEPOSITS_DIR):
            return

        for file_name in os.listdir(DEPOSITS_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(DEPOSITS_DIR, file_name)
                self.parse_deposit_file(file_path)

    def generate_excel(self):
        """生成Excel文件 - 修复颜色显示问题，特别是_per_pop修正"""

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}地块信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('地块信息')

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

        green_format = workbook.add_format({
            'border': 1,
            'font_color': 'green',
            'valign': 'vcenter'
        })

        red_format = workbook.add_format({
            'border': 1,
            'font_color': 'red',
            'valign': 'vcenter'
        })

        black_format = workbook.add_format({
            'border': 1,
            'font_color': 'black',
            'valign': 'vcenter'
        })

        # 设置列宽
        worksheet.set_column('A:A', 15)  # 图片
        worksheet.set_column('B:B', 30)  # 地块名称
        worksheet.set_column('C:C', 30)  # 地块KEY
        worksheet.set_column('D:D', 20)  # 地块主类型
        worksheet.set_column('E:E', 20)  # 地块副类型
        worksheet.set_column('F:F', 50)  # 地块效果
        worksheet.set_column('G:G', 80)  # 地块修正
        worksheet.set_column('H:H', 20)  # 来源文件

        # 设置行高
        worksheet.set_default_row(40)

        # 写入表头
        headers = ['图片', '地块名称', '地块KEY', '地块主类型', '地块副类型', '地块效果', '地块修正', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for deposit in self.deposits:
            # 写入图片
            worksheet.write(row, 0, '', cell_format)
            if deposit['icon'] and os.path.exists(deposit['icon']):
                worksheet.insert_image(row, 0, deposit['icon'], {
                    'x_offset': 5,
                    'y_offset': 5,
                    'x_scale': 0.8,
                    'y_scale': 0.8,
                    'object_position': 1
                })

            worksheet.write(row, 1, deposit['name'], cell_format)
            worksheet.write(row, 2, deposit['key'], cell_format)
            worksheet.write(row, 3, deposit['main_category'], cell_format)
            worksheet.write(row, 4, deposit['sub_category'], cell_format)
            worksheet.write(row, 5, deposit['tooltip'], description_format)
            worksheet.write(row, 7, deposit['source'], cell_format)

            # 构建修正效果文本
            if deposit['modifiers']:
                # 检查是否需要使用rich string
                needs_rich_string = False
                for modifier in deposit['modifiers']:
                    if 'merged_modifiers' in modifier:
                        # 合并的修正需要rich string
                        needs_rich_string = True
                        break
                    elif modifier['value_color'] != 'black' and modifier['formatted_value']:
                        needs_rich_string = True
                        break

                if needs_rich_string:
                    rich_text = []
                    for i, modifier in enumerate(deposit['modifiers']):
                        if i > 0:
                            rich_text.append(black_format)
                            rich_text.append('；')  # 使用中文分号分隔不同种类的修正

                        # 处理合并的修正
                        if 'merged_modifiers' in modifier:
                            description_text = modifier['description'] if modifier['description'] else " "
                            rich_text.append(black_format)
                            rich_text.append(description_text)

                            # 添加合并的修正项 - 修复颜色显示问题
                            for j, merged_mod in enumerate(modifier['merged_modifiers']):
                                if j > 0:
                                    rich_text.append(black_format)
                                    rich_text.append('，')  # 修正之间用逗号分隔

                                # 描述部分用黑色
                                desc_text = merged_mod['description'] if merged_mod['description'] else " "
                                rich_text.append(black_format)
                                rich_text.append(desc_text)

                                # 值部分用对应颜色（如果有值）
                                value_text = merged_mod['formatted_value'] if merged_mod['formatted_value'] else ""
                                if value_text:
                                    if merged_mod['value_color'] == 'green':
                                        rich_text.append(green_format)
                                    elif merged_mod['value_color'] == 'red':
                                        rich_text.append(red_format)
                                    else:
                                        rich_text.append(black_format)
                                    rich_text.append(value_text)
                        else:
                            # 处理普通修正 - 修复颜色显示问题
                            description_text = modifier['description'] if modifier['description'] else " "
                            formatted_value = modifier['formatted_value'] if modifier['formatted_value'] else ""

                            # 描述部分用黑色
                            rich_text.append(black_format)
                            rich_text.append(description_text)

                            # 值部分用对应颜色（如果有值）
                            if formatted_value:
                                if modifier['value_color'] == 'green':
                                    rich_text.append(green_format)
                                elif modifier['value_color'] == 'red':
                                    rich_text.append(red_format)
                                else:
                                    rich_text.append(black_format)
                                rich_text.append(formatted_value)

                    if rich_text and len(rich_text) >= 2:  # 确保至少有两个片段
                        worksheet.write_rich_string(row, 6, *rich_text)
                    else:
                        # 如果只有一个片段，使用普通写入
                        text_parts = []
                        for i, modifier in enumerate(deposit['modifiers']):
                            if i > 0:
                                text_parts.append('；')
                            if 'merged_modifiers' in modifier:
                                merged_texts = []
                                for merged_mod in modifier['merged_modifiers']:
                                    merged_texts.append(f"{merged_mod['description']}{merged_mod['formatted_value']}")
                                text_parts.append(f"{modifier['description']}{'，'.join(merged_texts)}")
                            else:
                                text_parts.append(f"{modifier['description']}{modifier['formatted_value']}")
                        worksheet.write(row, 6, ''.join(text_parts), cell_format)
                else:
                    # 所有修正都是黑色，直接拼接文本
                    text_parts = []
                    for i, modifier in enumerate(deposit['modifiers']):
                        if i > 0:
                            text_parts.append('；')
                        if 'merged_modifiers' in modifier:
                            merged_texts = []
                            for merged_mod in modifier['merged_modifiers']:
                                merged_texts.append(f"{merged_mod['description']}{merged_mod['formatted_value']}")
                            text_parts.append(f"{modifier['description']}{'，'.join(merged_texts)}")
                        else:
                            text_parts.append(f"{modifier['description']}{modifier['formatted_value']}")
                    worksheet.write(row, 6, ''.join(text_parts), cell_format)
            else:
                worksheet.write(row, 6, "无修正", cell_format)

            row += 1

        workbook.close()

    def save_debug_info(self):
        """保存调试信息"""

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}地块信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 地块信息调试信息 ===\n\n")

            if self.debug_info['missing_tooltip']:
                f.write("=== 缺失地块效果 ===\n")
                for info in sorted(self.debug_info['missing_tooltip']):
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

    def run(self):
        """运行完整的提取流程"""

        self.load_localization_files()
        self.load_scripted_variables()
        self.load_scripted_loc()
        self.parse_all_deposits()
        self.generate_excel()
        self.save_debug_info()


# 运行提取器
if __name__ == "__main__":
    extractor = DepositExtractor()
    extractor.run()