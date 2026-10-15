import os
import re
from pathlib import Path
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 条件翻译字典 import CONDITION_TRANSLATIONS
from 修正翻译字典 import MODIFIER_TRANSLATIONS
from 颜色转换字典 import SPECIAL_SUFFIXES


class RelicExtractor:
    def __init__(self):
        self.localisation_dict = {}
        self.relics = []
        self.relic_image_map = {}
        self.static_modifiers_dict = {}  # 存储静态修正用于激活效果解析
        self.scripted_variables = {}  # 存储脚本变量
        self.insensitive_localization_dict = {}  # 新增：大小写不敏感的本地化字典
        self.scripted_loc_dict = {}  # 新增：脚本本地化字典

        # 调试信息
        self.debug_info = {
            'missing_images': [],
            'missing_modifiers': set(),
            'missing_static_modifiers': set(),
            'parsing_errors': [],
            'image_mapping_details': [],
            'activation_effect_details': [],
            'activation_cost_details': [],
            'outer_modifier_details': [],
            'active_effect_extraction_details': [],  # 新增：active_effect提取详情
            'missing_tooltip': set(),  # 新增：缺失本地化
            'missing_bracket_replacement': set(),  # 新增：缺失[]替换符
            'missing_pound_replacement': set(),  # 新增：缺失£xxx£替换符
            'scripted_variable_debug': []  # 新增：脚本变量调试信息
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(RELICS_IMAGES_DIR, exist_ok=True)

    def convert_dds_to_png(self, dds_path, png_path):
        """
        将DDS文件转换为PNG格式。

        Args:
            dds_path (str): 源DDS文件路径。
            png_path (str): 目标PNG文件路径。
        """
        try:
            # 检查源文件是否存在
            if not os.path.exists(dds_path):
                return False

            # 如果目标PNG目录不存在，则创建
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            # 使用imageio读取DDS并写入PNG
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

        # 检查是否需要百分比格式化
        def should_format_as_percentage(modifier_key):
            # 在不转换百分比列表中
            if modifier_key in NON_PERCENTAGE_KEYS:
                return False

            # 在强制转换百分比列表中
            if modifier_key in FORCE_PERCENTAGE_KEYS:
                return True

            # 以特定后缀结尾
            if (modifier_key.endswith('_gain') or modifier_key.endswith('_perc') or
                    modifier_key.endswith('_mult') or modifier_key.endswith('_mult_add') or
                    modifier_key.endswith('_happiness') or modifier_key.endswith('_speed') or
                    modifier_key.endswith('_uninhabitable') or modifier_key.endswith('_reduction')):
                return True

            return False

        # 决定是否使用百分比格式
        use_percentage = should_format_as_percentage(key)

        # 特殊处理_add后缀（如果不是强制百分比）
        if key.endswith('_add') and key not in FORCE_PERCENTAGE_KEYS:
            use_percentage = False

        # 格式化数值
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
        """根据修正键和值确定颜色 - 按照标本提取中的方式"""
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
        """预处理本地化文本 - 修改：保留§R和§G颜色标记符，删除其他颜色标记符，保留换行符"""
        if not text:
            return text

        # 1. 递归处理$$替换符
        text, _ = self.process_dollar_replacement(text)

        # 2. 处理颜色标识符：保留§R和§G，删除其他颜色标识符
        # 匹配所有颜色标记符：§X内容§!
        def replace_color_markers(match):
            color_code = match.group(1)  # 颜色代码（R, G, Y等）
            content = match.group(2)  # 标记的内容

            # 只保留红色(§R)和绿色(§G)，其他颜色标记符删除标记，保留内容
            if color_code == 'R':
                return f"§R{content}§!"  # 保留红色标记
            elif color_code == 'G':
                return f"§G{content}§!"  # 保留绿色标记
            else:
                return content  # 其他颜色标记符：删除标记，只保留内容

        text = re.sub(r'§([A-Z])(.*?)§!', replace_color_markers, text)

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

            if var_name in self.scripted_variables:
                result = self.scripted_variables[var_name] * multiplier

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
        """处理特殊替换符：$@xxx|任意数值%$ - 修改：修复正则表达式"""
        if not text:
            return text

        # 查找 $@xxx|任意数值%$ 格式 - 修改：更精确的正则表达式
        pattern = r'\$@(\w+)\|([^$]+)%?\$'

        def replace_special_dollar_at(match):
            full_match = match.group(0)
            var_name = match.group(1)
            default_value = match.group(2)
            has_percent = '%' in full_match

            if var_name in self.scripted_variables:
                result = self.scripted_variables[var_name]

                if has_percent:
                    result = result * 100
                    if abs(result - round(result)) < 0.001:
                        formatted = f"{int(result)}%"
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

        return re.sub(pattern, replace_special_dollar_at, text)

    def extract_modifier_value(self, value_str):
        """提取修正值，处理@变量 - 参照传统提取重写"""
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
            if var_name in self.scripted_variables:
                return self.scripted_variables[var_name]
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
            if match in self.localisation_dict:
                replacement = self.localisation_dict[match]
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
        """处理[]替换符 - 参照传统提取重写"""
        if not text or current_depth >= max_depth:
            return text

        # 处理 [GetIcon] 类型的替换
        text = re.sub(r'\[[^\]]*\.GetIcon\]', '', text)

        # 处理特殊格式的[]替换符
        def replace_brackets(match):
            inner_text = match.group(1)
            result = inner_text

            # 1. 优先级最高：处理 ['xxx', yyy] 格式，直接输出yyy
            comma_match = re.match(r"'([^']*)'\s*,\s*([^']+)", inner_text)
            if comma_match:
                result = comma_match.group(2)
                return result

            # 2. 优先级次高：处理 ['xxx':yyy] 格式，直接输出yyy
            colon_match = re.match(r"'([^']*)':\s*([^']+)", inner_text)
            if colon_match:
                result = colon_match.group(2)
                return result

            # 3. 优先级再次：处理 [xxx.GetNamePlural] 或 [xxx.GetName]
            name_match = re.match(r'(\w+)\.GetName(Plural)?', inner_text)
            if name_match:
                prefix = name_match.group(1)
                is_plural = name_match.group(2) is not None

                # 尝试查找 job_前缀的本地化
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

            # 4. 优先级再次：处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith(
                    "'") and ':' not in inner_text and ',' not in inner_text:
                concept_key = inner_text[1:-1]  # 去掉单引号
                if concept_key in self.localisation_dict:
                    result = self.localisation_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            # 5. 优先级再次：处理 ['xxx:yyy'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)  # 去掉单引号并按冒号分割
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localisation_dict:
                        result = self.localisation_dict[building_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到建筑本地化: {building_key} (来自 [{inner_text}])"
                        )

            # 6. 常规[]替换符处理 - 查找scripted_loc_dict
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

            # 递归处理替换后的文本中的[]替换符
            if current_depth < max_depth - 1:
                result = self.process_bracket_replacement(result, max_depth, current_depth + 1)

            # 新增：在递归处理后处理$@替换符
            result = self.process_special_dollar_at_replacement(result)

            return result

        return re.sub(r'\[([^\]]+)\]', replace_brackets, text)

    def process_pound_replacement(self, text):
        """处理£xxx£替换符 - 修复版本，替换后继续处理其他替换符"""
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
                    # 直接删除未找到对应翻译的£xxx£替换符
                    return ""

        text = re.sub(r'£([^£]+)£', replace_pounds, text)

        # 替换后继续处理其他替换符
        text, _ = self.process_dollar_replacement(text)
        text = self.process_bracket_replacement(text)

        return text

    def process_special_pound_replacement(self, text):
        """处理特殊的£xxx£替换符 - 修改：删除特殊替换，只保留清除£blocker£字段"""
        if not text:
            return text

        # 固定清除£blocker£字段
        text = text.replace('£blocker£', '')

        return text

    def process_color_markers(self, text):
        """处理颜色标记符 - 新增方法：保留§R和§G，删除其他颜色标记符"""
        if not text:
            return text

        # 匹配所有颜色标记符：§X内容§!
        def replace_color_markers(match):
            color_code = match.group(1)  # 颜色代码（R, G, Y等）
            content = match.group(2)  # 标记的内容

            # 只保留红色(§R)和绿色(§G)，其他颜色标记符删除标记，保留内容
            if color_code == 'R':
                return f"§R{content}§!"  # 保留红色标记
            elif color_code == 'G':
                return f"§G{content}§!"  # 保留绿色标记
            else:
                return content  # 其他颜色标记符：删除标记，只保留内容

        text = re.sub(r'§([A-Z])(.*?)§!', replace_color_markers, text)

        return text

    def process_tradition_text(self, text, modifier_key=""):
        """通用文本处理方法 - 修改：循环处理直到没有变化"""
        if not text:
            return text

        # 设置最大迭代次数，防止无限循环
        max_iterations = 10
        iteration = 0
        last_text = ""

        while iteration < max_iterations and text != last_text:
            last_text = text
            iteration += 1

            # 1. 预处理：处理$$替换符和清理颜色标识符（保留§R和§G）
            text = self.preprocess_localization_text(text)

            # 2. 清理空白文本（保留换行符）
            text = self.clean_whitespace_text_with_newlines(text)

            # 3. 处理[]替换符
            text = self.process_bracket_replacement(text)

            # 4. 处理特殊$@xxx|任意数值%$替换符
            text = self.process_special_dollar_at_replacement(text)

            # 5. 处理特殊£xxx£替换（固定翻译）
            text = self.process_special_pound_replacement(text)

            # 6. 处理常规£xxx£替换
            text = self.process_pound_replacement(text)

            # 检查是否还有未处理的替换符
            if text == last_text:
                break

        # 7. 清除重复文本
        text = self.remove_duplicate_text(text)

        # 8. 处理换行符 - 修改：正确处理换行符
        text = self.process_newlines(text)

        # 9. 最终清理空白（保留换行符）
        text = self.clean_whitespace_text_with_newlines(text)

        return text

    def process_newlines(self, text):
        """处理换行符 - 新增方法：确保换行符正确处理"""
        if not text:
            return text

        # 将\\n替换为实际的换行符\n
        text = text.replace('\\n', '\n')

        return text

    def clean_whitespace_text_with_newlines(self, text):
        """清理空白文本但保留换行符 - 修改：正确处理换行符"""
        if not text:
            return text

        # 按行分割，清理每行的空白，然后重新组合
        lines = text.split('\n')
        cleaned_lines = []
        for line in lines:
            # 清理每行内的空白（保留换行符）
            cleaned_line = re.sub(r'[ \t]+', ' ', line)  # 替换多个空格和制表符为单个空格
            cleaned_line = cleaned_line.strip()  # 去除首尾空格
            cleaned_lines.append(cleaned_line)

        # 重新组合，保留换行符
        text = '\n'.join(cleaned_lines)

        return text

    def format_resource_value(self, value):
        """格式化资源值，用于resources块 - 修复：不添加+号"""
        try:
            value_float = float(value)
        except (ValueError, TypeError):
            return str(value)

        if value_float == int(value_float):
            return f"{int(value_float)}"
        else:
            return f"{value_float}"

    def parse_localisation_files(self):
        """解析本地化文件，构建翻译字典 - 修改：使用传统提取的加载方式"""
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

        # 构建大小写不敏感的本地化字典
        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localisation_dict.items()}
        print(f"已加载 {len(self.localisation_dict)} 个本地化条目")

    def load_scripted_variables(self):
        """加载脚本变量 - 修改：使用传统提取的加载方式"""
        print("正在加载脚本变量...")

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
                            self.scripted_variables[var_name] = float(var_value)
                            self.debug_info['scripted_variable_debug'].append(
                                f"加载脚本变量: @{var_name} = {var_value}")

                    except Exception as e:
                        print(f"加载脚本变量文件失败: {file_path}, 错误: {e}")

        # 从relics目录加载
        if os.path.exists(RELICS_DIR):
            for file_name in os.listdir(RELICS_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(RELICS_DIR, file_name)
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
                                    self.scripted_variables[var_name] = float(var_value)
                                    self.debug_info['scripted_variable_debug'].append(
                                        f"加载遗珍脚本变量: @{var_name} = {var_value}")

                    except Exception as e:
                        print(f"加载遗珍文件中的脚本变量失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.scripted_variables)} 个脚本变量")
        for var_name, var_value in self.scripted_variables.items():
            print(f"  脚本变量: @{var_name} = {var_value}")

    def load_scripted_loc(self):
        """加载脚本本地化 - 参照传统提取"""
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

    def extract_complete_block(self, content, start_pos):
        """提取完整的花括号块，处理嵌套结构"""
        brace_count = 0
        in_quotes = False
        escape_next = False

        # 找到第一个开括号的位置
        for i in range(start_pos, len(content)):
            if content[i] == '{':
                start_pos = i
                break

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
                        return content[start_pos:i + 1]

        return None

    def extract_image_name_from_portrait(self, portrait_value):
        """从portrait值中提取图片名称"""
        if not portrait_value:
            return None

        # 移除可能的引号
        portrait_value = portrait_value.strip('"\'')

        return portrait_value

    def load_relic_images(self):
        """加载遗珍图片映射 - 修复版本，处理嵌套结构"""
        self.relic_image_map = {}

        if not os.path.exists(RELICS_GFX_FILE_PATH):
            print(f"遗珍图片映射文件不存在: {RELICS_GFX_FILE_PATH}")
            return

        try:
            with open(RELICS_GFX_FILE_PATH, 'r', encoding='utf-8') as f:
                content = f.read()

            # 使用更强大的方法提取spriteType块 - 处理嵌套结构
            sprite_blocks = self.extract_all_sprite_blocks(content)

            print(f"找到 {len(sprite_blocks)} 个spriteType块")

            for i, block in enumerate(sprite_blocks):
                # 提取name值
                name_match = re.search(r'name\s*=\s*"([^"]+)"', block)
                if not name_match:
                    continue

                name = name_match.group(1)

                # 提取texturefile值
                texture_match = re.search(r'texturefile\s*=\s*"([^"]+)"', block, re.IGNORECASE)
                if not texture_match:
                    # 记录没有texturefile的块
                    self.debug_info['image_mapping_details'].append(f"{name}: 没有texturefile")
                    continue

                texture_path = texture_match.group(1)

                # 构建完整的DDS文件路径
                texture_path = texture_path.replace('/', os.path.sep)
                absolute_dds_path = os.path.join(BASE_GAME_DIR, texture_path)

                # 添加到映射字典
                self.relic_image_map[name] = absolute_dds_path

                # 记录调试信息
                self.debug_info['image_mapping_details'].append(f"{name} -> {absolute_dds_path}")

            print(f"已加载 {len(self.relic_image_map)} 个遗珍图片映射")

        except Exception as e:
            print(f"加载遗珍图片映射失败: {RELICS_GFX_FILE_PATH}, 错误: {e}")

    def extract_all_sprite_blocks(self, content):
        """提取所有spriteType块，处理嵌套结构"""
        blocks = []
        pos = 0

        while pos < len(content):
            # 查找spriteType关键字
            sprite_match = re.search(r'spriteType\s*=\s*\{', content[pos:], re.IGNORECASE)
            if not sprite_match:
                break

            start_pos = pos + sprite_match.start()
            # 提取完整的块
            block = self.extract_complete_block(content, start_pos)
            if block:
                blocks.append(block)
                pos = start_pos + len(block)
            else:
                pos = start_pos + 1

        return blocks

    def get_relic_image_path(self, relic_key, portrait_value):
        """获取遗珍图片路径 - 修复：直接使用portrait_value作为映射键"""
        if not portrait_value:
            return None

        # 移除可能的引号
        portrait_value = portrait_value.strip('"\'')

        # 直接使用portrait_value作为映射键，不构建GFX_relic_xxx
        mapping_key = portrait_value

        if mapping_key in self.relic_image_map:
            dds_path = self.relic_image_map[mapping_key]

            # 生成PNG文件路径
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(RELICS_IMAGES_DIR, png_filename)

            # 如果PNG图片已存在，直接返回路径
            if os.path.exists(png_path):
                return png_path
            else:
                # 如果PNG不存在，但DDS存在，则进行转换
                if os.path.exists(dds_path):
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        # 记录缺失的图片
                        self.debug_info['missing_images'].append(f"{relic_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{relic_key} -> DDS文件不存在: {dds_path}")
                    return None
        else:
            # 如果没有找到映射，尝试直接查找图片文件
            image_name = self.extract_image_name_from_portrait(portrait_value)
            possible_extensions = ['.png', '.jpg', '.jpeg']
            for ext in possible_extensions:
                test_path = os.path.join(RELICS_IMAGES_DIR, f"{image_name}{ext}")
                if os.path.exists(test_path):
                    return test_path

            # 记录缺失的图片
            self.debug_info['missing_images'].append(f"{relic_key} -> 无映射且文件不存在: {portrait_value}")
            return None

    def get_resource_localization(self, resource_key):
        """获取资源的本地化名称 - 按照标本提取中的方式"""
        # 先尝试concept_前缀
        concept_key = f"concept_{resource_key}"
        if concept_key in self.localisation_dict:
            return self.localisation_dict[concept_key]

        # 再尝试直接查找
        if resource_key in self.localisation_dict:
            return self.localisation_dict[resource_key]

        # 如果都找不到，返回原始键
        return resource_key

    def load_static_modifiers(self):
        """加载静态修正文件，用于解析激活效果"""
        print("正在加载静态修正文件...")

        if not os.path.exists(STATIC_MODIFIER_DIR):
            print(f"静态修正目录不存在: {STATIC_MODIFIER_DIR}")
            return

        for file_path in Path(STATIC_MODIFIER_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                # 移除单行注释
                content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                # 使用正则表达式找到所有静态修正定义的开始
                pattern = r'^(\w+)\s*=\s*\{'
                matches = list(re.finditer(pattern, content, re.MULTILINE))

                for match in matches:
                    modifier_key = match.group(1)
                    start_pos = match.start()

                    # 提取完整的修正块
                    modifier_block = self.extract_complete_block(content, start_pos)

                    if modifier_block:
                        self.static_modifiers_dict[modifier_key] = modifier_block

            except Exception as e:
                print(f"解析静态修正文件时出错 {file_path}: {e}")

        print(f"已加载 {len(self.static_modifiers_dict)} 个静态修正定义")

    def parse_static_modifier_effects(self, modifier_key):
        """解析静态修正的效果"""
        if modifier_key not in self.static_modifiers_dict:
            self.debug_info['missing_static_modifiers'].add(modifier_key)
            return []

        modifier_block = self.static_modifiers_dict[modifier_key]
        effects = []

        # 提取修正字段 (排除mult和potential)
        modifier_pattern = r'(\w+)\s*=\s*([\d.-]+)'
        modifier_matches = re.findall(modifier_pattern, modifier_block)

        for modifier_key, modifier_value in modifier_matches:
            if modifier_key in ['mult', 'potential']:
                continue

            # 查找修正的本地化翻译
            translation = self.get_modifier_description(modifier_key)

            if translation and translation != modifier_key:  # 确保找到了翻译
                # 使用智能格式化修正值
                value_str = self.format_modifier_value(modifier_key, modifier_value)
                value_color = self.get_value_color(modifier_key, modifier_value)

                # 修正：使用正确的颜色标记符格式
                color_code = 'G' if value_color == 'green' else 'R' if value_color == 'red' else ''
                if color_code:
                    effect_text = f"{translation}§{color_code}{value_str}§!"
                else:
                    effect_text = f"{translation}{value_str}"

                effects.append({
                    'text': effect_text,
                    'color': 'black'  # 文本部分保持黑色
                })
            else:
                # 只有当修正键确实找不到本地化时才记录
                if modifier_key not in MODIFIER_TRANSLATIONS:
                    self.debug_info['missing_modifiers'].add(modifier_key)

        return effects

    def extract_outer_add_modifier(self, active_effect_block):
        """提取最外层的add_modifier，排除hidden_effect中的"""
        # 首先移除hidden_effect块
        cleaned_block = re.sub(r'hidden_effect\s*=\s*\{[^{}]*?(?:\{[^{}]*\}[^{}]*?)*\}', '', active_effect_block)

        # 查找最外层的add_modifier - 修改：支持带引号和不带引号的modifier
        outer_modifier_match = re.search(r'add_modifier\s*=\s*\{([^{}]*)\}', cleaned_block)

        return outer_modifier_match

    def extract_active_effect_block(self, relic_block, relic_key):
        """提取active_effect块 - 只使用方法1"""
        # 方法1：使用正则表达式查找active_effect
        active_effect_match = re.search(r'active_effect\s*=\s*\{', relic_block)

        if active_effect_match:
            start_pos = active_effect_match.start()
            active_effect_block = self.extract_complete_block(relic_block, start_pos)

            if active_effect_block:
                self.debug_info['active_effect_extraction_details'].append(
                    f"{relic_key}: 成功 - 找到active_effect块，长度: {len(active_effect_block)}"
                )
                return active_effect_block
            else:
                self.debug_info['active_effect_extraction_details'].append(
                    f"{relic_key}: 失败 - 正则找到但无法提取完整块"
                )
        else:
            self.debug_info['active_effect_extraction_details'].append(
                f"{relic_key}: 失败 - 未找到active_effect关键字"
            )

        return None

    def extract_all_custom_tooltips(self, content):
        """提取所有custom_tooltip，包括hidden_effect外面的"""
        tooltips = []

        # 查找所有custom_tooltip，包括带引号和不带引号的
        tooltip_pattern1 = r'custom_tooltip\s*=\s*"([^"]+)"'
        tooltip_matches1 = re.findall(tooltip_pattern1, content)

        tooltip_pattern2 = r'custom_tooltip\s*=\s*([^\s{]+)'
        tooltip_matches2 = re.findall(tooltip_pattern2, content)

        all_tooltips = tooltip_matches1 + tooltip_matches2

        # 过滤掉不需要的tooltip
        filtered_tooltips = [tooltip for tooltip in all_tooltips if
                             tooltip not in ["relic_triumph_cooldown", "relic_triumph_cooldown_short"]]

        return filtered_tooltips

    def parse_relic_files(self):
        """解析遗珍文件，提取遗珍定义"""
        print("正在解析遗珍文件...")
        relics = []

        for file_path in Path(RELICS_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                    # 移除单行注释
                    content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                    # 使用正则表达式找到所有遗珍定义的开始
                    pattern = r'^(\w+)\s*=\s*\{'
                    matches = list(re.finditer(pattern, content, re.MULTILINE))

                    for match in matches:
                        relic_key = match.group(1)
                        start_pos = match.start()

                        # 提取完整的遗珍块（处理嵌套结构）
                        relic_block = self.extract_complete_block(content, start_pos)

                        if not relic_block:
                            print(f"警告: 无法找到遗珍 {relic_key} 的完整定义块")
                            continue

                        # 提取肖像 (portrait) - 用于图片
                        portrait_value = None
                        portrait_patterns = [
                            r'portrait\s*=\s*"([^"]*)"',  # 带双引号
                            r"portrait\s*=\s*'([^']*)'",  # 带单引号
                            r'portrait\s*=\s*(\S+)'  # 不带引号
                        ]

                        for pattern in portrait_patterns:
                            portrait_match = re.search(pattern, relic_block)
                            if portrait_match:
                                portrait_value = portrait_match.group(1)
                                break

                        # 方法2: 如果正则表达式失败，使用逐行搜索
                        if not portrait_value:
                            lines = relic_block.split('\n')
                            for line in lines:
                                line = line.strip()
                                if line.startswith('portrait'):
                                    # 尝试提取portrait值
                                    parts = re.split(r'\s*=\s*', line, 1)
                                    if len(parts) > 1:
                                        portrait_value = parts[1].strip()
                                        break

                        # 获取图片路径
                        image_path = self.get_relic_image_path(relic_key, portrait_value)

                        # 获取遗珍名称翻译
                        raw_name = self.localisation_dict.get(relic_key, relic_key)
                        relic_name = self.process_tradition_text(raw_name)

                        # 提取被动效果
                        passive_effects = []

                        # 1. 提取resources块中produces的内容
                        resources_match = re.search(r'resources\s*=\s*\{([^{}]*?(?:\{[^{}]*\}[^{}]*?)*)\}', relic_block)
                        if resources_match:
                            resources_block = resources_match.group(1)
                            produces_match = re.search(r'produces\s*=\s*\{([^{}]*)\}', resources_block)
                            if produces_match:
                                produces_content = produces_match.group(1)
                                resource_pattern = r'(\w+)\s*=\s*([\d.-]+)'
                                resource_matches = re.findall(resource_pattern, produces_content)
                                for resource_key, resource_value in resource_matches:
                                    resource_translation = self.get_resource_localization(resource_key)
                                    # 确保被动效果文本也经过process_tradition_text处理
                                    resource_translation = self.process_tradition_text(resource_translation)
                                    # 修复：使用format_resource_value而不是format_modifier_value，避免多输出+号
                                    processed_value = self.extract_modifier_value(resource_value)
                                    value_str = self.format_resource_value(processed_value)
                                    passive_effects.append({
                                        'text': f"每月{resource_translation}§G+{value_str}§!",
                                        'color': 'black'
                                    })

                        # 2. 提取triggered_country_modifier效果 - 包括custom_tooltip
                        modifier_pattern = r'triggered_country_modifier\s*=\s*\{([^{}]*?(?:\{[^{}]*\}[^{}]*?)*)\}'
                        modifier_blocks = re.findall(modifier_pattern, relic_block)
                        for modifier_block in modifier_blocks:
                            # 提取custom_tooltip
                            tooltip_matches = self.extract_all_custom_tooltips(modifier_block)
                            for tooltip_key in tooltip_matches:
                                tooltip_key = tooltip_key.strip('"\'')
                                tooltip_translation = self.process_tradition_text(
                                    self.localisation_dict.get(tooltip_key, tooltip_key)
                                )
                                passive_effects.append({
                                    'text': tooltip_translation,
                                    'color': 'black'
                                })

                            # 提取限制条件（如果有potential块）- 忽略always = yes
                            condition = ""
                            potential_match = re.search(r'potential\s*=\s*\{([^}]*)\}', modifier_block)
                            if potential_match:
                                potential_content = potential_match.group(1)
                                condition = self.parse_condition(potential_content)

                            # 提取修正字段 (排除mult和potential)
                            modifier_pattern = r'(\w+)\s*=\s*([\d.-]+)'
                            modifier_matches = re.findall(modifier_pattern, modifier_block)

                            for modifier_key, modifier_value in modifier_matches:
                                if modifier_key in ['mult', 'potential']:
                                    continue

                                # 查找修正的本地化翻译
                                translation = self.get_modifier_description(modifier_key)

                                if translation and translation != modifier_key:
                                    # 处理修正值（包括@变量）
                                    processed_value = self.extract_modifier_value(modifier_value)
                                    # 使用智能格式化修正值
                                    value_str = self.format_modifier_value(modifier_key, processed_value)
                                    value_color = self.get_value_color(modifier_key, processed_value)

                                    # 修正：使用正确的颜色标记符格式
                                    color_code = 'G' if value_color == 'green' else 'R' if value_color == 'red' else ''
                                    if color_code:
                                        effect_text = f"{translation}§{color_code}{value_str}§!"
                                    else:
                                        effect_text = f"{translation}{value_str}"

                                    # 确保被动效果文本也经过process_tradition_text处理
                                    effect_text = self.process_tradition_text(effect_text)

                                    if condition:
                                        effect_text = f"（{condition}）{effect_text}"

                                    passive_effects.append({
                                        'text': effect_text,
                                        'color': 'black'
                                    })
                                else:
                                    if modifier_key not in MODIFIER_TRANSLATIONS:
                                        self.debug_info['missing_modifiers'].add(modifier_key)

                        # 3. 提取score值 - 修改为胜利分数
                        score_match = re.search(r'score\s*=\s*(\d+)', relic_block)
                        if score_match:
                            score_value = score_match.group(1)
                            passive_effects.append({
                                'text': f"胜利分数§G+{score_value}§!",
                                'color': 'black'
                            })

                        # 提取激活花费
                        activation_cost = []
                        activation_cost_details = []

                        # 1. 提取resources块中cost的内容
                        if resources_match:
                            resources_block = resources_match.group(1)
                            cost_match = re.search(r'cost\s*=\s*\{([^{}]*)\}', resources_block)
                            if cost_match:
                                cost_content = cost_match.group(1)
                                cost_pattern = r'(\w+)\s*=\s*([\d.-]+)'
                                cost_matches = re.findall(cost_pattern, cost_content)
                                for cost_key, cost_value in cost_matches:
                                    cost_translation = self.get_resource_localization(cost_key)
                                    cost_translation = self.process_tradition_text(cost_translation)
                                    # 修复：正确处理@变量
                                    processed_value = self.extract_modifier_value(cost_value)
                                    value_str = self.format_resource_value(processed_value)
                                    activation_cost.append({
                                        'text': f"{value_str}{cost_translation}",
                                        'color': 'black'
                                    })
                                    activation_cost_details.append(
                                        f"找到cost: {cost_key} = {cost_value} -> {value_str}")
                            else:
                                activation_cost_details.append("未找到cost块")
                        else:
                            activation_cost_details.append("未找到resources块")

                        # 2. 提取activation_duration的值 - 修复：正确替换脚本变量
                        duration_match = re.search(r'activation_duration\s*=\s*@?(\w+)', relic_block)
                        if duration_match:
                            duration_value = duration_match.group(1)
                            activation_cost_details.append(f"找到activation_duration: {duration_value}")

                            # 处理脚本变量替换
                            if duration_value.startswith('@'):
                                var_name = duration_value[1:]
                                if var_name in self.scripted_variables:
                                    actual_duration = self.scripted_variables[var_name]
                                    duration_display = f"持续{actual_duration}天"
                                    activation_cost_details.append(f"脚本变量替换: @{var_name} -> {actual_duration}")
                                else:
                                    duration_display = f"持续{var_name}天"
                                    activation_cost_details.append(f"脚本变量未找到: @{var_name}")
                            else:
                                duration_display = f"持续{duration_value}天"

                            activation_cost.append({
                                'text': duration_display,
                                'color': 'black'
                            })
                        else:
                            activation_cost_details.append("未找到activation_duration")

                        # 3. 提取active_effect中最外层的add_modifier中的modifier的值的本地化翻译
                        # 使用改进的方法提取active_effect块
                        active_effect_block = self.extract_active_effect_block(relic_block, relic_key)
                        outer_modifier_match = None  # 初始化变量
                        outer_modifier_details = []

                        if active_effect_block:
                            outer_modifier_details.append(f"active_effect块长度: {len(active_effect_block)}")

                            # 提取active_effect块的内容（去掉active_effect = { ... }的外层）
                            if active_effect_block.startswith('active_effect'):
                                # 找到第一个{和最后一个}
                                start_brace = active_effect_block.find('{')
                                end_brace = active_effect_block.rfind('}')
                                if start_brace != -1 and end_brace != -1:
                                    active_effect_content = active_effect_block[start_brace + 1:end_brace]
                                else:
                                    active_effect_content = active_effect_block
                            else:
                                active_effect_content = active_effect_block

                            # 使用新方法提取最外层的add_modifier
                            outer_modifier_match = self.extract_outer_add_modifier(active_effect_content)

                            if outer_modifier_match:
                                modifier_block = outer_modifier_match.group(1)
                                outer_modifier_details.append(f"找到最外层add_modifier: {modifier_block}")

                                # 修改：支持带引号和不带引号的modifier
                                modifier_name_match = re.search(r'modifier\s*=\s*"([^"]+)"', modifier_block)
                                if not modifier_name_match:
                                    # 尝试不带引号的匹配
                                    modifier_name_match = re.search(r'modifier\s*=\s*(\w+)', modifier_block)

                                if modifier_name_match:
                                    modifier_key = modifier_name_match.group(1)
                                    modifier_translation = self.process_tradition_text(
                                        self.localisation_dict.get(modifier_key, modifier_key)
                                    )
                                    activation_cost.append({
                                        'text': f"添加{modifier_translation}修正",
                                        'color': 'black'
                                    })
                                    activation_cost_details.append(f"找到modifier: {modifier_key}")
                                    outer_modifier_details.append(f"找到modifier: {modifier_key}")
                                else:
                                    activation_cost_details.append("在add_modifier中未找到modifier字段")
                                    outer_modifier_details.append("在add_modifier中未找到modifier字段")
                            else:
                                activation_cost_details.append("未找到最外层的add_modifier")
                                outer_modifier_details.append("未找到最外层的add_modifier")

                                # 详细分析active_effect块内容
                                add_modifier_count = len(re.findall(r'add_modifier\s*=', active_effect_content))
                                hidden_effect_count = len(re.findall(r'hidden_effect\s*=', active_effect_content))
                                outer_modifier_details.append(f"active_effect中add_modifier总数: {add_modifier_count}")
                                outer_modifier_details.append(
                                    f"active_effect中hidden_effect总数: {hidden_effect_count}")
                        else:
                            activation_cost_details.append("未找到active_effect块")
                            outer_modifier_details.append("未找到active_effect块")

                        # 保存激活花费和outer_modifier的调试信息
                        self.debug_info['activation_cost_details'].append({
                            'relic_key': relic_key,
                            'details': activation_cost_details
                        })
                        self.debug_info['outer_modifier_details'].append({
                            'relic_key': relic_key,
                            'details': outer_modifier_details
                        })

                        # 提取激活效果
                        activation_effects = []

                        # 1. 提取active_effect中所有custom_tooltip的值的本地化（排除特定值）
                        if active_effect_block:
                            # 提取active_effect块的内容（去掉active_effect = { ... }的外层）
                            if active_effect_block.startswith('active_effect'):
                                # 找到第一个{和最后一个}
                                start_brace = active_effect_block.find('{')
                                end_brace = active_effect_block.rfind('}')
                                if start_brace != -1 and end_brace != -1:
                                    active_effect_content = active_effect_block[start_brace + 1:end_brace]
                                else:
                                    active_effect_content = active_effect_block
                            else:
                                active_effect_content = active_effect_block

                            # 记录active_effect块的内容用于调试
                            self.debug_info['activation_effect_details'].append(
                                f"{relic_key}: active_effect内容长度 = {len(active_effect_content)}"
                            )

                            # 使用新方法提取所有custom_tooltip
                            filtered_tooltips = self.extract_all_custom_tooltips(active_effect_content)

                            for tooltip_key in filtered_tooltips:
                                tooltip_key = tooltip_key.strip('"\'')
                                tooltip_translation = self.process_tradition_text(
                                    self.localisation_dict.get(tooltip_key, tooltip_key)
                                )
                                activation_effects.append({
                                    'text': tooltip_translation,
                                    'color': 'black'
                                })

                            # 记录激活效果提取详情
                            self.debug_info['activation_effect_details'].append(
                                f"{relic_key}: 找到 {len(filtered_tooltips)} 个custom_tooltip: {filtered_tooltips}"
                            )

                        # 2. 提取active_effect中最外层的add_modifier中的modifier的对应的全局修正的内容
                        if outer_modifier_match:
                            modifier_block = outer_modifier_match.group(1)
                            # 修改：支持带引号和不带引号的modifier
                            modifier_name_match = re.search(r'modifier\s*=\s*"([^"]+)"', modifier_block)
                            if not modifier_name_match:
                                # 尝试不带引号的匹配
                                modifier_name_match = re.search(r'modifier\s*=\s*(\w+)', modifier_block)

                            if modifier_name_match:
                                modifier_key = modifier_name_match.group(1)
                                static_effects = self.parse_static_modifier_effects(modifier_key)
                                activation_effects.extend(static_effects)

                                # 记录静态修正提取详情
                                self.debug_info['activation_effect_details'].append(
                                    f"{relic_key}: 找到静态修正 {modifier_key}，包含 {len(static_effects)} 个效果"
                                )
                            else:
                                # 记录没有找到modifier的情况
                                self.debug_info['activation_effect_details'].append(
                                    f"{relic_key}: 在add_modifier块中未找到modifier字段"
                                )
                        else:
                            # 记录没有找到outer_modifier_match的情况
                            self.debug_info['activation_effect_details'].append(
                                f"{relic_key}: 未找到最外层的add_modifier"
                            )

                        # 3. 提取其他可能的激活效果（如add_monthly_resource_mult等）
                        if active_effect_block:
                            # 提取active_effect块的内容（去掉active_effect = { ... }的外层）
                            if active_effect_block.startswith('active_effect'):
                                # 找到第一个{和最后一个}
                                start_brace = active_effect_block.find('{')
                                end_brace = active_effect_block.rfind('}')
                                if start_brace != -1 and end_brace != -1:
                                    active_effect_content = active_effect_block[start_brace + 1:end_brace]
                                else:
                                    active_effect_content = active_effect_block
                            else:
                                active_effect_content = active_effect_block

                            # 查找add_monthly_resource_mult效果
                            monthly_resource_pattern = r'add_monthly_resource_mult\s*=\s*\{([^{}]*)\}'
                            monthly_resource_matches = re.findall(monthly_resource_pattern, active_effect_content)
                            for resource_block in monthly_resource_matches:
                                # 提取resource和value
                                resource_match = re.search(r'resource\s*=\s*(\w+)', resource_block)
                                value_match = re.search(r'value\s*=\s*@?(\w+)', resource_block)
                                if resource_match and value_match:
                                    resource_key = resource_match.group(1)
                                    value = value_match.group(1)
                                    resource_translation = self.get_resource_localization(resource_key)
                                    resource_translation = self.process_tradition_text(resource_translation)
                                    activation_effects.append({
                                        'text': f"每月获得{resource_translation}",
                                        'color': 'black'
                                    })

                            # 记录其他激活效果
                            if monthly_resource_matches:
                                self.debug_info['activation_effect_details'].append(
                                    f"{relic_key}: 找到 {len(monthly_resource_matches)} 个add_monthly_resource_mult效果"
                                )

                        relics.append({
                            'key': relic_key,
                            'name': relic_name,
                            'portrait': portrait_value,
                            'image_path': image_path,
                            'image_name': self.extract_image_name_from_portrait(portrait_value) or relic_key,
                            'passive_effects': passive_effects,
                            'activation_cost': activation_cost,
                            'activation_effects': activation_effects,
                            'original_index': len(relics)  # 记录原始解析顺序
                        })

            except Exception as e:
                error_msg = f"解析遗珍文件时出错 {file_path}: {e}"
                print(error_msg)
                self.debug_info['parsing_errors'].append(error_msg)

        self.relics = relics
        print(f"已找到 {len(relics)} 个遗珍定义")

    def parse_condition(self, condition_content):
        """解析条件并返回翻译后的文本 - 忽略always = yes条件"""
        # 查找所有条件键值对
        condition_pattern = r'(\w+)\s*=\s*(\w+)'
        condition_matches = re.findall(condition_pattern, condition_content)

        conditions = []
        for condition_key, condition_value in condition_matches:
            # 忽略 always = yes
            if condition_key == "always" and condition_value == "yes":
                continue

            if condition_key in CONDITION_TRANSLATIONS:
                condition_text = CONDITION_TRANSLATIONS[condition_key]
                if condition_value == "yes":
                    conditions.append(f"{condition_text}效果")
                elif condition_value == "no":
                    conditions.append(f"非{condition_text}效果")
                else:
                    conditions.append(f"{condition_text}({condition_value})效果")
            else:
                # 未知条件，直接显示
                conditions.append(f"{condition_key}={condition_value}")

        return "；".join(conditions) if conditions else ""

    def generate_output(self):
        """生成输出结果"""
        excel_data = []

        # 生成Excel数据
        for relic in self.relics:
            excel_data.append({
                "图片": relic['image_path'] or "",  # 图片路径
                "遗珍名称": relic['name'],
                "遗珍KEY": relic['key'],
                "被动效果": relic['passive_effects'],
                "激活花费": relic['activation_cost'],
                "激活效果": relic['activation_effects'],
                "图片名称": relic['image_name']
            })

        return excel_data

    def generate_debug_output(self):
        """生成调试输出 - 包含详细调试信息"""
        output_lines = []

        # 添加脚本变量调试信息
        if self.debug_info['scripted_variable_debug']:
            output_lines.append("=== 脚本变量调试信息 ===")
            for debug_info in self.debug_info['scripted_variable_debug']:
                output_lines.append(f"  - {debug_info}")
            output_lines.append("")

        # 添加图片映射详情
        if self.debug_info['image_mapping_details']:
            output_lines.append("=== 图片映射详情 ===")
            output_lines.append(f"总共加载了 {len(self.relic_image_map)} 个有效的图片映射")
            output_lines.append(f"总共记录了 {len(self.debug_info['image_mapping_details'])} 个映射详情")

            # 按类型分类显示
            valid_mappings = [d for d in self.debug_info['image_mapping_details'] if '->' in d]
            no_texturefile = [d for d in self.debug_info['image_mapping_details'] if '没有texturefile' in d]

            if valid_mappings:
                output_lines.append("")
                output_lines.append("有效的图片映射:")
                for detail in valid_mappings[:20]:  # 只显示前20个
                    output_lines.append(f"  - {detail}")
                if len(valid_mappings) > 20:
                    output_lines.append(f"  - ... 还有 {len(valid_mappings) - 20} 个映射")

            if no_texturefile:
                output_lines.append("")
                output_lines.append("没有texturefile的spriteType:")
                for detail in no_texturefile[:10]:  # 只显示前10个
                    output_lines.append(f"  - {detail}")
                if len(no_texturefile) > 10:
                    output_lines.append(f"  - ... 还有 {len(no_texturefile) - 10} 个")

            output_lines.append("")

        # 添加active_effect提取详情
        if self.debug_info['active_effect_extraction_details']:
            output_lines.append("=== active_effect提取详情 ===")
            for detail in self.debug_info['active_effect_extraction_details']:
                output_lines.append(f"  - {detail}")
            output_lines.append("")

        # 添加激活花费提取详情
        if self.debug_info['activation_cost_details']:
            output_lines.append("=== 激活花费提取详情 ===")
            for cost_info in self.debug_info['activation_cost_details']:
                output_lines.append(f"{cost_info['relic_key']}:")
                for detail in cost_info['details']:
                    output_lines.append(f"  - {detail}")
                output_lines.append("")
            output_lines.append("")

        # 添加最外层add_modifier提取详情
        if self.debug_info['outer_modifier_details']:
            output_lines.append("=== 最外层add_modifier提取详情 ===")
            for modifier_info in self.debug_info['outer_modifier_details']:
                output_lines.append(f"{modifier_info['relic_key']}:")
                for detail in modifier_info['details']:
                    output_lines.append(f"  - {detail}")
                output_lines.append("")
            output_lines.append("")

        # 添加激活效果提取详情
        if self.debug_info['activation_effect_details']:
            output_lines.append("=== 激活效果提取详情 ===")
            # 按遗珍分组显示
            relic_effects = {}
            for detail in self.debug_info['activation_effect_details']:
                if ':' in detail:
                    relic_key = detail.split(':')[0]
                    if relic_key not in relic_effects:
                        relic_effects[relic_key] = []
                    relic_effects[relic_key].append(detail)

            for relic_key, details in relic_effects.items():
                output_lines.append(f"{relic_key}:")
                for detail in details:
                    output_lines.append(f"  - {detail.split(':', 1)[1]}")
                output_lines.append("")

            output_lines.append("")

        # 添加缺失修正字段的说明
        if self.debug_info['missing_modifiers']:
            output_lines.append("以下修正找不到本地化文件：")
            for modifier in sorted(self.debug_info['missing_modifiers']):
                output_lines.append(f"  - {modifier}")
            output_lines.append("")

        # 添加缺失静态修正的说明
        if self.debug_info['missing_static_modifiers']:
            output_lines.append("以下静态修正找不到定义：")
            for modifier in sorted(self.debug_info['missing_static_modifiers']):
                output_lines.append(f"  - {modifier}")
            output_lines.append("")

        # 添加缺失图片的说明
        if self.debug_info['missing_images']:
            output_lines.append("以下遗珍找不到对应图片：")
            missing_relic_keys = set()
            for info in self.debug_info['missing_images']:
                match = re.search(r'^([^->]+) ->', info)
                if match:
                    missing_relic_keys.add(match.group(1))

            for key in sorted(missing_relic_keys):
                output_lines.append(f"  - {key}")
            output_lines.append("")

        # 添加解析错误
        if self.debug_info['parsing_errors']:
            output_lines.append("=== 解析错误 ===")
            for error in self.debug_info['parsing_errors']:
                output_lines.append(f"  - {error}")
            output_lines.append("")

        # 添加脚本变量信息
        if self.scripted_variables:
            output_lines.append("=== 脚本变量 ===")
            for var_name, var_value in sorted(self.scripted_variables.items()):
                output_lines.append(f"  - @{var_name} = {var_value}")
            output_lines.append("")

        # 添加传统提取中的调试信息
        if self.debug_info['missing_tooltip']:
            output_lines.append("=== 缺失传统效果 ===")
            for info in sorted(self.debug_info['missing_tooltip']):
                output_lines.append(f"{info}")
            output_lines.append("")

        if self.debug_info['missing_bracket_replacement']:
            output_lines.append("=== 缺失[]替换符 ===")
            for info in sorted(self.debug_info['missing_bracket_replacement']):
                output_lines.append(f"{info}")
            output_lines.append("")

        if self.debug_info['missing_pound_replacement']:
            output_lines.append("=== 缺失£xxx£替换符 ===")
            for info in sorted(self.debug_info['missing_pound_replacement']):
                output_lines.append(f"{info}")
            output_lines.append("")

        # 如果没有调试信息
        if not any(self.debug_info.values()) and not self.scripted_variables:
            output_lines.append("没有发现调试信息")

        return output_lines

    def process_text_with_color_markers(self, text, workbook, default_color='black'):
        """处理文本中的颜色标记符，返回富文本格式"""
        if not text:
            return [], workbook.add_format({'font_color': default_color, 'text_wrap': True})

        # 创建颜色格式
        red_format = workbook.add_format({'font_color': 'red', 'text_wrap': True})
        green_format = workbook.add_format({'font_color': 'green', 'text_wrap': True})
        black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})

        # 如果没有颜色标记，直接返回默认颜色文本
        if '§R' not in text and '§G' not in text:
            format_obj = workbook.add_format({'font_color': default_color, 'text_wrap': True})
            return [(format_obj, text)], format_obj

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

            else:
                current_text += text[i]
                i += 1

        # 添加最后一段文本
        if current_text:
            rich_text_parts.append((current_format, current_text))

        return rich_text_parts, black_format

    def save_to_excel(self, excel_data):
        """将数据保存到Excel文件，并嵌入图片 - 使用xlsxwriter"""
        # 如果没有数据，直接返回
        if not excel_data:
            print("没有数据可导出到Excel")
            return

        # 保存到Excel文件，添加版本号
        excel_file = os.path.join(OUTPUT_DIR, f"{VERSION}遗珍.xlsx")

        # 使用xlsxwriter创建工作簿和工作表
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet("遗珍数据")

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

        wrap_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'text_wrap': True
        })

        # 设置列宽
        worksheet.set_column('A:A', 15)  # 图片
        worksheet.set_column('B:B', 20)  # 遗珍名称
        worksheet.set_column('C:C', 15)  # 遗珍KEY
        worksheet.set_column('D:D', 40)  # 被动效果
        worksheet.set_column('E:E', 40)  # 激活花费
        worksheet.set_column('F:F', 40)  # 激活效果
        worksheet.set_column('G:G', 25)  # 图片名称

        # 设置行高以容纳图片和文本
        worksheet.set_default_row(80)

        # 写入表头
        headers = ["图片", "遗珍名称", "遗珍KEY", "被动效果", "激活花费", "激活效果", "图片名称"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 添加数据
        for row, data in enumerate(excel_data, 1):
            # 添加图片
            image_path = data["图片"]
            if image_path and os.path.exists(image_path):
                try:
                    # 插入图片，调整大小适应单元格
                    worksheet.insert_image(row, 0, image_path, {
                        'x_scale': 0.8,
                        'y_scale': 0.8,
                        'x_offset': 5,
                        'y_offset': 5,
                        'object_position': 1  # 移动和调整大小单元格
                    })
                except Exception as e:
                    print(f"插入图片失败: {image_path}, 错误: {e}")
                    worksheet.write(row, 0, "图片加载失败", cell_format)
            else:
                worksheet.write(row, 0, "无图片", cell_format)

            # 写入基本数据
            worksheet.write(row, 1, data["遗珍名称"], cell_format)
            worksheet.write(row, 2, data["遗珍KEY"], cell_format)
            worksheet.write(row, 6, data["图片名称"], cell_format)

            # 处理被动效果（带颜色）
            passive_effects = data["被动效果"]
            if passive_effects:
                effect_text_parts = []
                for i, effect in enumerate(passive_effects):
                    if i > 0:
                        effect_text_parts.append("\n")

                    # 处理文本中的颜色标记
                    rich_parts, _ = self.process_text_with_color_markers(effect['text'], workbook)

                    for format_obj, text in rich_parts:
                        effect_text_parts.append((format_obj, text))

                # 写入富文本
                if len(effect_text_parts) == 1 and not isinstance(effect_text_parts[0], tuple):
                    # 只有一段文本
                    worksheet.write(row, 3, effect_text_parts[0], wrap_format)
                else:
                    # 多段富文本
                    rich_text_args = []
                    for part in effect_text_parts:
                        if isinstance(part, tuple):
                            rich_text_args.extend(part)
                        else:
                            rich_text_args.append(wrap_format)
                            rich_text_args.append(part)

                    worksheet.write_rich_string(row, 3, *rich_text_args)
            else:
                worksheet.write(row, 3, "无被动效果", wrap_format)

            # 处理激活花费（带颜色）
            activation_cost = data["激活花费"]
            if activation_cost:
                cost_text_parts = []
                for i, cost in enumerate(activation_cost):
                    if i > 0:
                        cost_text_parts.append("\n")

                    # 处理文本中的颜色标记
                    rich_parts, _ = self.process_text_with_color_markers(cost['text'], workbook)

                    for format_obj, text in rich_parts:
                        cost_text_parts.append((format_obj, text))

                # 写入富文本
                if len(cost_text_parts) == 1 and not isinstance(cost_text_parts[0], tuple):
                    # 只有一段文本
                    worksheet.write(row, 4, cost_text_parts[0], wrap_format)
                else:
                    # 多段富文本
                    rich_text_args = []
                    for part in cost_text_parts:
                        if isinstance(part, tuple):
                            rich_text_args.extend(part)
                        else:
                            rich_text_args.append(wrap_format)
                            rich_text_args.append(part)

                    worksheet.write_rich_string(row, 4, *rich_text_args)
            else:
                worksheet.write(row, 4, "无激活花费", wrap_format)

            # 处理激活效果（带颜色）
            activation_effects = data["激活效果"]
            if activation_effects:
                effect_text_parts = []
                for i, effect in enumerate(activation_effects):
                    if i > 0:
                        effect_text_parts.append("\n")

                    # 处理文本中的颜色标记
                    rich_parts, _ = self.process_text_with_color_markers(effect['text'], workbook)

                    for format_obj, text in rich_parts:
                        effect_text_parts.append((format_obj, text))

                # 写入富文本
                if len(effect_text_parts) == 1 and not isinstance(effect_text_parts[0], tuple):
                    # 只有一段文本
                    worksheet.write(row, 5, effect_text_parts[0], wrap_format)
                else:
                    # 多段富文本
                    rich_text_args = []
                    for part in effect_text_parts:
                        if isinstance(part, tuple):
                            rich_text_args.extend(part)
                        else:
                            rich_text_args.append(wrap_format)
                            rich_text_args.append(part)

                    worksheet.write_rich_string(row, 5, *rich_text_args)
            else:
                worksheet.write(row, 5, "无激活效果", wrap_format)

        workbook.close()
        print(f"Excel文件已保存到: {excel_file}")

    def save_debug_info(self):
        """保存调试信息到TXT文件"""
        print("正在保存调试信息...")

        # 生成调试输出
        debug_lines = self.generate_debug_output()

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}遗珍信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 遗珍信息调试信息 ===\n\n")
            f.write('\n'.join(debug_lines))

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取遗珍信息...")

        self.parse_localisation_files()
        self.load_scripted_variables()  # 加载脚本变量用于持续时间替换
        self.load_scripted_loc()  # 加载脚本本地化
        self.load_relic_images()
        self.load_static_modifiers()  # 加载静态修正用于激活效果解析
        self.parse_relic_files()

        print("正在生成输出...")
        excel_data = self.generate_output()

        # 保存调试信息到TXT文件
        self.save_debug_info()

        # 保存到Excel
        self.save_to_excel(excel_data)

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = RelicExtractor()
    extractor.run()