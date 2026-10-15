import os
import re
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 颜色转换字典 import SPECIAL_SUFFIXES
from 修正翻译字典 import MODIFIER_TRANSLATIONS
from 条件翻译字典 import CONDITION_TRANSLATIONS


class TraditionExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_variables_dict = {}
        self.scripted_loc_dict = {}
        self.traditions = []
        self.insensitive_localization_dict = {}
        self.tradition_image_map = {}

        # 调试信息 - 简化版本
        self.debug_info = {
            'missing_tooltip': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
            'missing_images': [],
            'skipped_traditions': []  # 恢复：跳过的传统记录
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(TRADITIONS_IMAGES_DIR, exist_ok=True)

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

            if var_name in self.scripted_variables_dict:
                result = self.scripted_variables_dict[var_name]

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

    def load_scripted_variables(self):
        """加载脚本变量"""
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
                            self.scripted_variables_dict[var_name] = float(var_value)

                    except Exception as e:
                        print(f"加载脚本变量文件失败: {file_path}, 错误: {e}")

        # 从traditions目录加载
        if os.path.exists(TRADITIONS_DIR):
            for file_name in os.listdir(TRADITIONS_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(TRADITIONS_DIR, file_name)
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
                        print(f"加载传统文件中的脚本变量失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.scripted_variables_dict)} 个脚本变量")

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
        """处理[]替换符 - 修改：在递归处理后添加$@替换符处理"""
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

            # 4. 优先级再次：处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith(
                    "'") and ':' not in inner_text and ',' not in inner_text:
                concept_key = inner_text[1:-1]  # 去掉单引号
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            # 5. 优先级再次：处理 ['xxx:yyy'] 格式
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

            # 6. 常规[]替换符处理 - 查找scripted_loc_dict
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

    def parse_nested_conditions(self, content, conditions, logic_context=None, parent_key=None, depth=0):
        """改进版的条件解析，正确处理嵌套逻辑块 - 修改：简化NOT块处理"""
        if not content.strip():
            return

        # 移除最外层花括号（如果有）
        if content.startswith('{') and content.endswith('}'):
            content = content[1:-1].strip()

        # 检查是否在NOT块中
        in_not_block = (logic_context == 'NOT')

        # 按行处理内容
        lines = content.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue

            # 跳过注释
            if '#' in line:
                line = line.split('#')[0].strip()

            # 处理逻辑块
            logic_match = re.match(r'(OR|AND|NOT|NOR|NAND)\s*=\s*\{', line)
            if logic_match:
                logic_type = logic_match.group(1)

                # 找到完整的逻辑块
                brace_start = line.find('{')
                full_line = line
                j = i

                # 如果当前行没有完整的块，继续读取后续行
                if full_line.count('{') != full_line.count('}'):
                    j = i
                    brace_count = 0
                    in_block = False
                    block_lines = []

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

                    full_block = '\n'.join(block_lines)
                    # 提取块内容（去掉逻辑类型和花括号）
                    block_start = full_block.find('{') + 1
                    block_end = full_block.rfind('}')
                    block_content = full_block[block_start:block_end].strip()
                else:
                    # 当前行包含完整的块
                    block_start = brace_start + 1
                    block_end = full_line.rfind('}')
                    block_content = full_line[block_start:block_end].strip()
                    j = i + 1

                # 特殊处理NOT块 - 修改：不嵌套括号，直接翻译条件
                if logic_type == 'NOT':
                    # NOT块中只有一个条件，直接提取并翻译
                    nested_conditions = []
                    # 递归解析NOT块内的条件，传递in_not_block=True
                    self.parse_nested_conditions(block_content, nested_conditions, 'NOT', parent_key, depth + 1)

                    # NOT块内只有一个条件，直接使用
                    if nested_conditions:
                        conditions.append(nested_conditions[0])
                else:
                    # 其他逻辑块正常处理
                    nested_conditions = []
                    self.parse_nested_conditions(block_content, nested_conditions, logic_type, parent_key, depth + 1)

                    # 根据逻辑类型添加条件
                    if logic_type == 'OR':
                        conditions.append(f"满足以下任意条件：({', '.join(nested_conditions)})")
                    elif logic_type == 'AND':
                        conditions.append(f"同时满足以下条件：({', '.join(nested_conditions)})")
                    elif logic_type == 'NOR':
                        conditions.append(f"不满足以下任意条件：({', '.join(nested_conditions)})")
                    elif logic_type == 'NAND':
                        conditions.append(f"不同时满足以下条件：({', '.join(nested_conditions)})")

                i = j
                continue

            # 处理简单的键值对条件
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

            i += 1

    def load_tradition_images(self):
        """加载传统图片映射 - 修改：使用imageio进行DDS到PNG转换"""
        self.tradition_image_map = {}

        if not os.path.exists(TRADITIONS_GFX_FILE_PATH):
            print(f"传统图片映射文件不存在: {TRADITIONS_GFX_FILE_PATH}")
            return

        try:
            with open(TRADITIONS_GFX_FILE_PATH, 'r', encoding='utf-8') as f:
                content = f.read()

            # 提取所有spriteType块
            sprite_blocks = re.findall(r'spriteType\s*=\s*\{([^}]+)\}', content)

            for block in sprite_blocks:
                # 提取name值
                name_match = re.search(r'name\s*=\s*"([^"]+)"', block)
                if not name_match:
                    continue

                name = name_match.group(1)

                # 检查是否是传统图片（以GFX_tr_开头）
                if not name.startswith('GFX_tr_'):
                    continue

                # 提取传统key（去掉GFX_前缀）
                tradition_key = name[4:]  # 去掉"GFX_"前缀

                # 提取textureFile值
                texture_match = re.search(r'textureFile\s*=\s*"([^"]+)"', block)
                if not texture_match:
                    continue

                texture_path = texture_match.group(1)

                # 构建完整的DDS文件路径（基础路径 + texturefile的值）
                # 注意：texturefile中的路径使用正斜杠，需要转换为系统路径分隔符
                texture_path = texture_path.replace('/', os.path.sep)
                absolute_dds_path = os.path.join(BASE_GAME_DIR, texture_path)

                # 添加到映射字典
                self.tradition_image_map[tradition_key] = absolute_dds_path

            print(f"已加载 {len(self.tradition_image_map)} 个传统图片映射")

        except Exception as e:
            print(f"加载传统图片映射失败: {TRADITIONS_GFX_FILE_PATH}, 错误: {e}")

    def get_tradition_image_path(self, tradition_key):
        """获取传统图片路径 - 修改：使用imageio进行DDS到PNG转换"""
        if hasattr(self, 'tradition_image_map') and tradition_key in self.tradition_image_map:
            dds_path = self.tradition_image_map[tradition_key]

            # 生成PNG文件路径（在TRADITIONS_IMAGES_DIR中，使用相同的文件名但扩展名为.png）
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(TRADITIONS_IMAGES_DIR, png_filename)

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
                        self.debug_info['missing_images'].append(f"{tradition_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{tradition_key} -> DDS文件不存在: {dds_path}")
                    return None
        else:
            # 如果没有找到映射，记录调试信息
            if tradition_key not in self.debug_info['missing_images']:
                self.debug_info['missing_images'].append(f"{tradition_key} -> 无映射")
            return None

    def process_inline_scripts(self, content, tradition_key, depth=0):
        """处理inline_script块 - 新增方法：递归替换inline_script块"""
        if depth > 5:  # 最大递归深度为5
            return content

        # 查找inline_script块
        inline_script_match = re.search(r'inline_script\s*=\s*\{', content)
        if not inline_script_match:
            return content  # 没有inline_script块，直接返回

        # 提取完整的inline_script块
        brace_start = inline_script_match.end() - 1
        inline_script_block, end_pos = self.extract_balanced_braces(content, brace_start)

        if not inline_script_block:
            return content  # 无法提取完整的块，直接返回

        # 解析inline_script块内容
        script_file_path = None
        extra_variables = {}

        # 提取script路径
        script_match = re.search(r'script\s*=\s*(\S+)', inline_script_block)
        if script_match:
            script_path = script_match.group(1)
            # 构建完整文件路径
            script_path = script_path.replace('/', '\\') + '.txt'
            script_file_path = os.path.join(INLINE_SCRIPTS_DIR, script_path)

        # 提取其他变量（如tradition）
        variable_matches = re.findall(r'(\w+)\s*=\s*(\S+)', inline_script_block)
        for var_name, var_value in variable_matches:
            if var_name != 'script':  # 跳过script，已经处理过了
                extra_variables[var_name] = var_value

        # 读取并处理inline_script文件
        if script_file_path and os.path.exists(script_file_path):
            try:
                with open(script_file_path, 'r', encoding='utf-8') as f:
                    script_content = f.read()

                # 移除注释
                script_content = re.sub(r'#.*', '', script_content)

                # 替换变量
                for var_name, var_value in extra_variables.items():
                    script_content = script_content.replace(f'${var_name}$', var_value)

                # 递归处理script_content中的inline_script
                script_content = self.process_inline_scripts(script_content, tradition_key, depth + 1)

                # 替换原始的inline_script块
                content = content.replace(inline_script_block, script_content)

                # 递归处理替换后的内容
                content = self.process_inline_scripts(content, tradition_key, depth + 1)

            except Exception as e:
                print(f"处理inline_script文件失败: {script_file_path}, 错误: {e}")
        else:
            print(f"inline_script文件不存在: {script_file_path}")

        return content

    def parse_tradition_file(self, file_path):
        """解析单个传统文件 - 修改：添加inline_script处理"""
        file_name = os.path.basename(file_path)
        print(f"正在解析文件: {file_name}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 使用更强大的方法提取完整的传统定义
            traditions = self.extract_complete_blocks(content)

            # 按文件分组存储传统
            file_traditions = []
            for tradition_key, tradition_content in traditions:
                print(f"  解析传统: {tradition_key}")

                # 处理inline_script块 - 新增
                processed_content = self.process_inline_scripts(tradition_content, tradition_key)

                tradition_data = self.parse_single_tradition(tradition_key, processed_content, file_name)
                if tradition_data:  # 只有当传统数据有效时才添加
                    file_traditions.append(tradition_data)

            # 对当前文件的传统进行排序：_finish结尾的排在最后
            if file_traditions:
                # 分离_finish和非_finish传统
                non_finish_traditions = [t for t in file_traditions if not t['key'].endswith('_finish')]
                finish_traditions = [t for t in file_traditions if t['key'].endswith('_finish')]

                # 合并：非_finish在前，_finish在后
                sorted_traditions = non_finish_traditions + finish_traditions

                # 添加到总列表
                self.traditions.extend(sorted_traditions)

        except Exception as e:
            print(f"解析文件失败: {file_path}, 错误: {e}")

    def extract_complete_blocks(self, content):
        """提取完整的传统定义块，正确处理嵌套花括号"""
        blocks = []
        lines = content.split('\n')

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 寻找传统定义开始
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

    def should_skip_tradition_swap(self, trigger_block):
        """检查是否应该跳过tradition_swap块（包含has_xxx_dlc = no条件）"""
        if not trigger_block:
            return False

        # 查找所有has_xxx_dlc条件
        dlc_pattern = r'has_\w+_dlc\s*=\s*(yes|no)'
        dlc_matches = re.findall(dlc_pattern, trigger_block)

        # 如果有任意一个has_xxx_dlc条件为no，则跳过
        for match in dlc_matches:
            if match == 'no':
                return True

        return False

    def filter_dlc_conditions(self, conditions):
        """过滤掉has_xxx_dlc条件"""
        filtered_conditions = []
        for condition in conditions:
            # 检查是否包含has_xxx_dlc条件
            if not re.search(r'has_\w+_dlc', condition):
                filtered_conditions.append(condition)
        return filtered_conditions

    def extract_trigger_conditions_with_dlc_filter(self, trigger_block):
        """从trigger块中提取条件，过滤掉has_xxx_dlc条件"""
        conditions = []

        # 移除最外层花括号
        if trigger_block.startswith('{') and trigger_block.endswith('}'):
            trigger_content = trigger_block[1:-1].strip()
        else:
            trigger_content = trigger_block

        # 使用递归解析嵌套条件
        self.parse_nested_conditions_with_dlc_filter(trigger_content, conditions)

        # 过滤掉has_xxx_dlc条件
        conditions = self.filter_dlc_conditions(conditions)

        return conditions

    def parse_nested_conditions_with_dlc_filter(self, content, conditions, logic_context=None, parent_key=None,
                                                depth=0):
        """改进版的条件解析，过滤掉has_xxx_dlc条件"""
        if not content.strip():
            return

        # 移除最外层花括号（如果有）
        if content.startswith('{') and content.endswith('}'):
            content = content[1:-1].strip()

        # 检查是否在NOT块中
        in_not_block = (logic_context == 'NOT')

        # 按行处理内容
        lines = content.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue

            # 跳过注释
            if '#' in line:
                line = line.split('#')[0].strip()

            # 处理逻辑块
            logic_match = re.match(r'(OR|AND|NOT|NOR|NAND)\s*=\s*\{', line)
            if logic_match:
                logic_type = logic_match.group(1)

                # 找到完整的逻辑块
                brace_start = line.find('{')
                full_line = line
                j = i

                # 如果当前行没有完整的块，继续读取后续行
                if full_line.count('{') != full_line.count('}'):
                    j = i
                    brace_count = 0
                    in_block = False
                    block_lines = []

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

                    full_block = '\n'.join(block_lines)
                    # 提取块内容（去掉逻辑类型和花括号）
                    block_start = full_block.find('{') + 1
                    block_end = full_block.rfind('}')
                    block_content = full_block[block_start:block_end].strip()
                else:
                    # 当前行包含完整的块
                    block_start = brace_start + 1
                    block_end = full_line.rfind('}')
                    block_content = full_line[block_start:block_end].strip()
                    j = i + 1

                # 特殊处理NOT块
                if logic_type == 'NOT':
                    # NOT块中只有一个条件，直接提取并翻译
                    nested_conditions = []
                    # 递归解析NOT块内的条件，传递in_not_block=True
                    self.parse_nested_conditions_with_dlc_filter(block_content, nested_conditions, 'NOT', parent_key,
                                                                 depth + 1)

                    # NOT块内只有一个条件，直接使用
                    if nested_conditions:
                        conditions.append(nested_conditions[0])
                else:
                    # 其他逻辑块正常处理
                    nested_conditions = []
                    self.parse_nested_conditions_with_dlc_filter(block_content, nested_conditions, logic_type,
                                                                 parent_key, depth + 1)

                    # 根据逻辑类型添加条件
                    if logic_type == 'OR':
                        conditions.append(f"满足以下任意条件：({', '.join(nested_conditions)})")
                    elif logic_type == 'AND':
                        conditions.append(f"同时满足以下条件：({', '.join(nested_conditions)})")
                    elif logic_type == 'NOR':
                        conditions.append(f"不满足以下任意条件：({', '.join(nested_conditions)})")
                    elif logic_type == 'NAND':
                        conditions.append(f"不同时满足以下条件：({', '.join(nested_conditions)})")

                i = j
                continue

            # 处理简单的键值对条件
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

            i += 1

    def parse_single_tradition(self, tradition_key, tradition_content, source_file):
        """解析单个传统定义 - 修改：调整tradition_swap处理"""
        # 恢复：检查possible块中是否有always = no
        skipped_due_to_always_no = False
        if self.should_skip_tradition(tradition_content, tradition_key):
            self.debug_info['skipped_traditions'].append(tradition_key)
            skipped_due_to_always_no = True

        tradition_data = {
            'key': tradition_key,
            'name': '',
            'description': '',
            'unlocks': '',
            'effects': '',
            'modifiers': [],
            'other_effects': [],
            'other_names': [],
            'source': source_file,
            'image_path': self.get_tradition_image_path(tradition_key),
            'skipped_due_to_always_no': skipped_due_to_always_no  # 恢复：标记是否因always=no被跳过
        }

        # 获取传统名称
        raw_name = self.localization_dict.get(tradition_key, tradition_key)
        tradition_data['name'] = self.process_tradition_text(raw_name)

        # 获取传统描述（特殊规则）
        tradition_data['description'] = self.get_tradition_description(tradition_key)

        # 提取解锁内容
        tradition_data['unlocks'] = self.extract_unlocks(tradition_content, tradition_key)

        # 提取传统效果（最外层的custom_tooltip或custom_tooltip_with_modifiers）
        tradition_data['effects'] = self.extract_effects(tradition_content, tradition_key)

        # 提取传统修正（分离处理modifier和triggered_modifier）
        self.extract_tradition_modifiers(tradition_content, tradition_data)

        # 提取tradition_swap块（修改：现在返回有效列表）
        valid_swaps = self.extract_tradition_swaps(tradition_content, tradition_data)

        # 将有效的tradition_swap添加到other_effects中
        for swap_data in valid_swaps:
            tradition_data['other_effects'].append(swap_data)

            # 构建其他名称条目
            name_text = swap_data['name']
            if swap_data['trigger_conditions']:
                conditions_text = '，'.join(swap_data['trigger_conditions'])
                name_text = f"{name_text}（{conditions_text}）"

            tradition_data['other_names'].append(name_text)

        return tradition_data

    def get_tradition_description(self, tradition_key):
        """获取传统描述（特殊规则）"""
        if tradition_key.endswith('_finish'):
            return ""
        elif tradition_key.endswith('_adopt'):
            # 特殊规则：tr_xxx_adopt -> tradition_xxx_desc
            middle_part = tradition_key[3:-6]  # 去掉'tr_'和'_adopt'
            desc_key = f"tradition_{middle_part}_desc"
        else:
            desc_key = f"{tradition_key}_delayed"

        raw_description = self.localization_dict.get(desc_key, desc_key)
        processed_description = self.process_tradition_text(raw_description)

        return processed_description

    def should_skip_tradition(self, tradition_content, tradition_key):
        """恢复：检查是否应该跳过该传统（possible块中有always = no）"""
        # 查找possible块
        possible_match = re.search(r'possible\s*=\s*\{', tradition_content)
        if possible_match:
            brace_start = possible_match.end() - 1
            possible_block, _ = self.extract_balanced_braces(tradition_content, brace_start)

            if possible_block:
                # 检查是否有always = no
                if 'always = no' in possible_block:
                    return True

        return False

    def extract_unlocks(self, tradition_content, tradition_key):
        """提取解锁内容"""
        unlocks = []

        # 查找unlocks_agenda
        agenda_match = re.search(r'unlocks_agenda\s*=\s*(\w+)', tradition_content)
        if agenda_match:
            agenda_key = agenda_match.group(1)

            # 构建本地化键：council_agenda_ + 议程key + _name
            agenda_localization_key = f"council_agenda_{agenda_key}_name"

            agenda_name = self.localization_dict.get(agenda_localization_key, agenda_key)
            agenda_name = self.process_tradition_text(agenda_name)

            unlocks.append(f"解锁议程：{agenda_name}")

        return '；'.join(unlocks) if unlocks else ""

    def extract_effects(self, tradition_content, tradition_key):
        """提取传统效果 - 修改：确保效果文本经过完整的替换符处理"""
        effects = []

        # 先移除tradition_swap块，确保只提取最外层的custom_tooltip
        tradition_content_without_swaps = self.remove_tradition_swaps(tradition_content)

        # 查找custom_tooltip
        tooltip_match = re.search(r'custom_tooltip\s*=\s*(\w+)', tradition_content_without_swaps)
        if tooltip_match:
            tooltip_key = tooltip_match.group(1)

            tooltip_text = self.localization_dict.get(tooltip_key, tooltip_key)
            # 确保效果文本经过完整的替换符处理
            tooltip_text = self.process_tradition_text(tooltip_text)

            effects.append(tooltip_text)

        # 查找custom_tooltip_with_modifiers
        tooltip_with_modifiers_match = re.search(r'custom_tooltip_with_modifiers\s*=\s*(\w+)',
                                                 tradition_content_without_swaps)
        if tooltip_with_modifiers_match:
            tooltip_key = tooltip_with_modifiers_match.group(1)

            tooltip_text = self.localization_dict.get(tooltip_key, tooltip_key)
            # 确保效果文本经过完整的替换符处理
            tooltip_text = self.process_tradition_text(tooltip_text)

            effects.append(tooltip_text)

        return '；'.join(effects) if effects else ""

    def remove_tradition_swaps(self, content):
        """移除tradition_swap块，确保只提取最外层的内容"""
        result = content
        pos = 0

        while pos < len(result):
            swap_match = re.search(r'tradition_swap\s*=\s*\{', result[pos:])
            if not swap_match:
                break

            brace_start = pos + swap_match.end() - 1
            swap_block, new_pos = self.extract_balanced_braces(result, brace_start)

            if swap_block:
                # 移除这个tradition_swap块
                result = result[:pos + swap_match.start()] + result[new_pos:]
            else:
                pos += swap_match.end()

        return result

    def extract_tradition_modifiers(self, tradition_content, tradition_data):
        """提取传统修正（分离处理modifier和triggered_modifier）"""
        # 修复问题1：使用负向前瞻避免匹配triggered_modifier
        modifier_match = re.search(r'(?<!triggered_)modifier\s*=\s*\{', tradition_content)
        if modifier_match:
            brace_start = modifier_match.end() - 1
            modifier_block, _ = self.extract_balanced_braces(tradition_content, brace_start)

            if modifier_block:
                modifiers = self.extract_modifiers_from_block_content(modifier_block, tradition_data['key'])
                tradition_data['modifiers'].extend(modifiers)

        # 提取最外层的triggered_modifier
        pos = 0
        while pos < len(tradition_content):
            triggered_modifier_match = re.search(r'triggered_modifier\s*=\s*\{', tradition_content[pos:])
            if not triggered_modifier_match:
                break

            brace_start = pos + triggered_modifier_match.end() - 1
            triggered_modifier_block, new_pos = self.extract_balanced_braces(tradition_content, brace_start)

            if triggered_modifier_block:
                # 提取potential条件
                conditions = []
                potential_match = re.search(r'potential\s*=\s*\{', triggered_modifier_block)
                if potential_match:
                    potential_brace_start = potential_match.end() - 1
                    potential_block, _ = self.extract_balanced_braces(triggered_modifier_block, potential_brace_start)

                    if potential_block:
                        conditions = self.extract_conditions_from_potential(potential_block)

                # 提取修正
                modifiers = self.extract_modifiers_from_block_content(triggered_modifier_block, tradition_data['key'])

                # 为每个修正添加条件前缀
                for modifier in modifiers:
                    if conditions:
                        condition_text = f"（{'，'.join(conditions)}）"
                        modifier['description'] = f"{condition_text}{modifier['description']}"

                tradition_data['modifiers'].extend(modifiers)

            pos = new_pos

    def extract_tradition_swaps(self, tradition_content, tradition_data):
        """提取tradition_swap块 - 修改：添加DLC条件过滤和_finish特殊处理"""
        pos = 0
        swap_count = 0

        # 用于存储有效的tradition_swap块
        valid_swaps = []

        while pos < len(tradition_content):
            swap_match = re.search(r'tradition_swap\s*=\s*\{', tradition_content[pos:])
            if not swap_match:
                break

            swap_count += 1

            brace_start = pos + swap_match.end() - 1
            swap_block, new_pos = self.extract_balanced_braces(tradition_content, brace_start)

            if swap_block:
                swap_data = self.process_single_tradition_swap(swap_block, tradition_data)

                # 只有返回有效数据时才处理
                if swap_data is not None:
                    valid_swaps.append(swap_data)

            pos = new_pos

        # 特殊处理：对于_finish传统，用第一个有效的tradition_swap替换主传统效果
        if tradition_data['key'].endswith('_finish') and valid_swaps:
            main_swap = valid_swaps[0]

            # 替换主传统效果
            if main_swap['description']:
                tradition_data['effects'] = main_swap['description']
            elif main_swap['delayed_description']:
                tradition_data['effects'] = main_swap['delayed_description']

            # 替换主传统修正
            if main_swap['modifiers']:
                tradition_data['modifiers'] = main_swap['modifiers']

            # 替换名称（如果tradition_swap有独立名称）
            if not main_swap['inherit_name'] and main_swap['name']:
                tradition_data['name'] = main_swap['name']

            # 对于_finish传统，不在副效果中再次输出tradition_swap
            return []

        return valid_swaps

    def process_single_tradition_swap(self, swap_block, tradition_data):
        """处理单个tradition_swap块 - 修改：添加DLC条件过滤"""
        # 首先检查trigger条件，判断是否应该跳过
        trigger_match = re.search(r'trigger\s*=\s*\{', swap_block)
        if trigger_match:
            brace_start = trigger_match.end() - 1
            trigger_block, _ = self.extract_balanced_braces(swap_block, brace_start)

            if trigger_block:
                # 检查是否应该跳过（包含has_xxx_dlc = no）
                if self.should_skip_tradition_swap(trigger_block):
                    return None

        swap_data = {
            'name': '',
            'trigger_conditions': [],
            'description': '',
            'modifiers': [],
            'effects': '',
            'delayed_description': '',
            'inherit_effects': False,
            'inherit_name': False
        }

        # 检查inherit_name = yes
        inherit_name_match = re.search(r'inherit_name\s*=\s*yes', swap_block)
        if inherit_name_match:
            swap_data['name'] = tradition_data['name']
            swap_data['inherit_name'] = True
        else:
            # 如果没有inherit_name = yes，再查找name
            name_match = re.search(r'name\s*=\s*(\w+)', swap_block)
            if name_match:
                name_key = name_match.group(1)

                # 获取name的本地化
                swap_data['name'] = self.localization_dict.get(name_key, name_key)
                swap_data['name'] = self.process_tradition_text(swap_data['name'])

        # 检查inherit_effects
        inherit_effects_match = re.search(r'inherit_effects\s*=\s*yes', swap_block)
        if inherit_effects_match:
            swap_data['inherit_effects'] = True

        # 提取trigger条件（使用过滤版本）
        if trigger_match and trigger_block:
            conditions = self.extract_trigger_conditions_with_dlc_filter(trigger_block)
            swap_data['trigger_conditions'] = conditions

        # 提取描述
        if inherit_name_match:
            # 如果继承名称，使用主传统的描述
            swap_data['delayed_description'] = tradition_data['description']
        else:
            # 否则使用name_key + "_delayed"查找描述
            if 'name_key' in locals():
                desc_key = name_key + "_delayed"
                raw_description = self.localization_dict.get(desc_key, "")
                swap_data['delayed_description'] = self.process_tradition_text(raw_description)

        # 根据inherit_effects决定如何处理效果和修正
        if swap_data['inherit_effects']:
            # 复制主传统的修正和效果
            swap_data['modifiers'] = tradition_data['modifiers'].copy()
            swap_data['effects'] = tradition_data['effects']
        else:
            # 提取custom_tooltip
            tooltip_match = re.search(r'custom_tooltip\s*=\s*(\w+)', swap_block)
            if tooltip_match:
                tooltip_key = tooltip_match.group(1)

                tooltip_text = self.localization_dict.get(tooltip_key, tooltip_key)
                tooltip_text = self.process_tradition_text(tooltip_text)
                swap_data['description'] = tooltip_text

            # 提取custom_tooltip_with_modifiers
            tooltip_with_modifiers_match = re.search(r'custom_tooltip_with_modifiers\s*=\s*(\w+)', swap_block)
            if tooltip_with_modifiers_match:
                tooltip_key = tooltip_with_modifiers_match.group(1)

                tooltip_text = self.localization_dict.get(tooltip_key, tooltip_key)
                tooltip_text = self.process_tradition_text(tooltip_text)
                swap_data['description'] = tooltip_text

            # 提取modifier
            modifier_match = re.search(r'modifier\s*=\s*\{', swap_block)
            if modifier_match:
                brace_start = modifier_match.end() - 1
                modifier_block, _ = self.extract_balanced_braces(swap_block, brace_start)

                if modifier_block:
                    modifiers = self.extract_modifiers_from_block_content(modifier_block, tradition_data['key'])
                    swap_data['modifiers'] = modifiers

        return swap_data

    def extract_modifiers_from_block_content(self, block_content, tradition_key):
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

            effect_description = self.process_tradition_text(effect_description, key)

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

    def parse_all_traditions(self):
        """解析所有传统文件 - 恢复：处理always=no的传统"""
        print("正在解析传统文件...")

        if not os.path.exists(TRADITIONS_DIR):
            print(f"传统目录不存在: {TRADITIONS_DIR}")
            return

        for file_name in os.listdir(TRADITIONS_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(TRADITIONS_DIR, file_name)
                self.parse_tradition_file(file_path)

        # 恢复：将传统分为两组：正常传统和因always = no被跳过的传统
        normal_traditions = [t for t in self.traditions if not t['skipped_due_to_always_no']]
        skipped_traditions = [t for t in self.traditions if t['skipped_due_to_always_no']]

        # 重新组合：正常传统在前，跳过的传统在后
        self.traditions = normal_traditions + skipped_traditions

        print(f"已解析 {len(self.traditions)} 个传统（其中 {len(skipped_traditions)} 个因always = no被标记）")

    def process_text_with_color_markers(self, text, workbook):
        """处理文本中的颜色标记符，返回富文本格式 - 新增方法"""
        if not text:
            return [], workbook.add_format({'font_color': 'black', 'text_wrap': True})

        # 创建颜色格式
        red_format = workbook.add_format({'font_color': 'red', 'text_wrap': True})
        green_format = workbook.add_format({'font_color': 'green', 'text_wrap': True})
        black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})

        # 如果没有颜色标记，直接返回黑色文本
        if '§R' not in text and '§G' not in text:
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

            else:
                current_text += text[i]
                i += 1

        # 添加最后一段文本
        if current_text:
            rich_text_parts.append((current_format, current_text))

        return rich_text_parts, black_format

    def generate_excel(self):
        """生成Excel文件 - 恢复：处理always=no的传统"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}传统信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('传统信息')

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
        worksheet.set_column('A:A', 15)  # 图片
        worksheet.set_column('B:B', 30)  # 传统名称
        worksheet.set_column('C:C', 30)  # 传统KEY
        worksheet.set_column('D:D', 50)  # 传统描述
        worksheet.set_column('E:E', 30)  # 解锁
        worksheet.set_column('F:F', 80)  # 传统效果（合并了传统修正）
        worksheet.set_column('G:G', 0)  # 传统修正（留空，宽度为0）
        worksheet.set_column('H:H', 80)  # 传统其他效果
        worksheet.set_column('I:I', 50)  # 传统其他名称
        worksheet.set_column('J:J', 20)  # 来源文件

        # 设置行高，为图片留出空间
        worksheet.set_default_row(100)

        # 写入表头
        headers = ['图片', '传统名称', '传统KEY', '传统描述', '解锁', '传统效果', '传统修正', '传统其他效果',
                   '传统其他名称', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for tradition in self.traditions:
            # 恢复：如果当前传统是跳过的，且前一个传统不是跳过的，则在中间插入一个空行
            if tradition['skipped_due_to_always_no'] and row > 1:
                # 检查前一个传统是否不是跳过的
                prev_tradition = self.traditions[self.traditions.index(tradition) - 1]
                if not prev_tradition['skipped_due_to_always_no']:
                    row += 1  # 插入空行

            # 插入图片
            if tradition['image_path']:
                try:
                    # 插入图片，调整大小适应单元格
                    worksheet.insert_image(row, 0, tradition['image_path'], {
                        'x_scale': 0.8,
                        'y_scale': 0.8,
                        'x_offset': 5,
                        'y_offset': 5,
                        'object_position': 1  # 移动和调整大小单元格
                    })
                except Exception as e:
                    print(f"插入图片失败: {tradition['image_path']}, 错误: {e}")
            else:
                # 如果没有图片，留空
                worksheet.write(row, 0, "无图片", cell_format)

            worksheet.write(row, 1, tradition['name'], cell_format)
            worksheet.write(row, 2, tradition['key'], cell_format)

            # 处理传统描述中的颜色标记
            if tradition['description']:
                rich_text_parts, _ = self.process_text_with_color_markers(tradition['description'], workbook)
                if len(rich_text_parts) == 1:
                    worksheet.write(row, 3, tradition['description'], description_format)
                else:
                    # 构建富文本
                    rich_text_args = []
                    for format_obj, text in rich_text_parts:
                        rich_text_args.append(format_obj)
                        rich_text_args.append(text)
                    worksheet.write_rich_string(row, 3, *rich_text_args)
            else:
                worksheet.write(row, 3, tradition['description'], description_format)

            worksheet.write(row, 4, tradition['unlocks'], cell_format)

            # 构建合并的效果文本（传统效果和传统修正）
            combined_effect_text = ""

            # 添加传统名称和描述
            combined_effect_text += f"{tradition['name']}：\n"
            if tradition['description']:
                combined_effect_text += f"{tradition['description']}\n"
            combined_effect_text += "\n效果：\n"

            # 添加传统效果
            if tradition['effects']:
                # 在传统效果每行前添加空格
                effect_lines = tradition['effects'].split('\n')
                for line in effect_lines:
                    if line.strip():  # 只处理非空行
                        combined_effect_text += f" {line}\n"
                combined_effect_text += "\n"

            # 添加传统修正
            if tradition['modifiers']:
                for modifier in tradition['modifiers']:
                    # 根据值的颜色添加颜色标记
                    value_color_code = ''
                    if modifier['value_color'] == 'red':
                        value_color_code = 'R'
                    elif modifier['value_color'] == 'green':
                        value_color_code = 'G'

                    # 如果值有颜色，则用颜色标记包裹值
                    if value_color_code:
                        value_text = f"§{value_color_code}{modifier['formatted_value']}§!"
                    else:
                        value_text = modifier['formatted_value']

                    # 描述部分我们不做颜色处理，因为描述可能已经包含颜色标记，我们希望保留
                    combined_effect_text += f" {modifier['description']}{value_text}\n"

            # 处理合并文本中的颜色标记
            rich_text_parts, _ = self.process_text_with_color_markers(combined_effect_text, workbook)
            if len(rich_text_parts) == 1:
                worksheet.write(row, 5, combined_effect_text, description_format)
            else:
                # 构建富文本
                rich_text_args = []
                for format_obj, text in rich_text_parts:
                    rich_text_args.append(format_obj)
                    rich_text_args.append(text)
                worksheet.write_rich_string(row, 5, *rich_text_args)

            # 传统修正一栏留空
            worksheet.write(row, 6, "", cell_format)

            worksheet.write(row, 9, tradition['source'], cell_format)

            # 构建其他效果文本
            if tradition['other_effects']:
                effect_texts = []
                for swap in tradition['other_effects']:
                    effect_parts = []
                    effect_parts.append(f"{swap['name']}（{'，'.join(swap['trigger_conditions'])}）：")

                    # 添加描述
                    if swap['delayed_description']:
                        effect_parts.append(f"描述：{swap['delayed_description']}")
                        effect_parts.append("")

                    effect_parts.append("效果：")

                    # 添加描述文本
                    if swap['description']:
                        effect_parts.append(f" {swap['description']}")
                    elif swap['effects']:  # 新增：如果有继承的效果
                        effect_parts.append(f" {swap['effects']}")

                    # 添加修正
                    if swap['modifiers']:
                        effect_parts.append("修正：")
                        for modifier in swap['modifiers']:
                            # 为每个修正构建完整的富文本字符串
                            color_code = ''
                            if modifier['value_color'] == 'red':
                                color_code = 'R'
                            elif modifier['value_color'] == 'green':
                                color_code = 'G'

                            if color_code:
                                # 如果有颜色，构建带颜色标记的字符串
                                modifier_text = f" {modifier['description']}§{color_code}{modifier['formatted_value']}§!"
                            else:
                                # 如果没有颜色，使用普通文本
                                modifier_text = f" {modifier['description']}{modifier['formatted_value']}"

                            effect_parts.append(modifier_text)

                    effect_parts.append("")
                    effect_texts.append('\n'.join(effect_parts))

                # 处理其他效果中的颜色标记
                other_effects_text = '\n'.join(effect_texts)
                rich_text_parts, _ = self.process_text_with_color_markers(other_effects_text, workbook)
                if len(rich_text_parts) == 1:
                    worksheet.write(row, 7, other_effects_text, description_format)
                else:
                    # 构建富文本
                    rich_text_args = []
                    for format_obj, text in rich_text_parts:
                        rich_text_args.append(format_obj)
                        rich_text_args.append(text)
                    worksheet.write_rich_string(row, 7, *rich_text_args)
            else:
                worksheet.write(row, 7, "无其他效果", cell_format)

            # 构建其他名称文本
            if tradition['other_names']:
                worksheet.write(row, 8, '，'.join(tradition['other_names']), description_format)
            else:
                worksheet.write(row, 8, "无其他名称", cell_format)

            row += 1

        workbook.close()
        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息 - 恢复：包含跳过的传统信息"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}传统信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 传统信息调试信息 ===\n\n")

            if self.debug_info['missing_tooltip']:
                f.write("=== 缺失传统效果 ===\n")
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

            # 恢复：跳过的传统信息
            if self.debug_info['skipped_traditions']:
                f.write("=== 因always = no被标记的传统 ===\n")
                for info in self.debug_info['skipped_traditions']:
                    f.write(f"{info}\n")
                f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取传统信息...")

        self.load_localization_files()
        self.load_scripted_variables()
        self.load_scripted_loc()
        self.load_tradition_images()
        self.parse_all_traditions()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = TraditionExtractor()
    extractor.run()