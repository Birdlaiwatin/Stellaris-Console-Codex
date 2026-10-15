import os
import re
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from DLC翻译字典 import DLC_TRANSLATIONS
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 颜色转换字典 import SPECIAL_SUFFIXES
from 修正翻译字典 import MODIFIER_TRANSLATIONS
from 条件翻译字典 import CONDITION_TRANSLATIONS


class AscensionPerkExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_variables_dict = {}
        self.ascension_perks = []
        self.insensitive_localization_dict = {}
        self.perk_image_map = {}

        # 调试信息 - 记录缺失的内容
        self.debug_info = {
            'missing_perk_name': set(), # 缺失的飞升名称
            'missing_tooltip': set(), # 缺失的本地化
            'missing_scripted_variable': set(), # 缺失的具体值
            'missing_images': [], # 缺失的图片文件
            'missing_condition_translations': set(),  # 缺失的条件翻译
            'missing_modifier_translations': set()  # 缺失的修正翻译
        }

        # 确保输出目录存在，如果没有将会新建一个
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(PERK_IMAGE_DIR, exist_ok=True)

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

    def get_modifier_description(self, modifier_key):
        """获取修正效果的本地化描述，优先从内置字典查找，在查找不到时，按照以下规则查找"""
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

        # 记录缺失的修正翻译
        if modifier_key not in self.debug_info['missing_modifier_translations']:
            self.debug_info['missing_modifier_translations'].add(modifier_key)

        self.debug_info['missing_tooltip'].add(modifier_key)
        return modifier_key

    def preprocess_localization_text(self, text):
        """预处理本地化文本 - 保留§R和§G颜色标记符，删除其他颜色标记符"""
        if not text:
            return text

        # 1. 递归处理$$替换符
        text, _ = self.process_dollar_replacement(text)

        # 2. 处理颜色标识符：保留§R和§G，删除其他颜色标识符
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
                if abs(result - round(result)) < 1e-10:
                    result = round(result)

                # 如果结果是整数，返回整数形式
                if result == int(result):
                    return str(int(result))
                else:
                    # 对于非整数，返回原始浮点数
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
                self.debug_info['missing_scripted_variable'].add(var_name)
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

        # 从ascension_perks目录加载
        if os.path.exists(ASCENSION_PERKS_DIR):
            for file_name in os.listdir(ASCENSION_PERKS_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(ASCENSION_PERKS_DIR, file_name)
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
                        print(f"加载飞升文件中的脚本变量失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.scripted_variables_dict)} 个脚本变量")

    def extract_modifier_value(self, value_str):
        """提取修正值，处理@变量"""
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
                self.debug_info['missing_scripted_variable'].add(var_name)
                return value_str

        # 处理$$替换符
        if isinstance(value_str, str) and '$' in value_str:
            value_str, _ = self.process_dollar_replacement(value_str)

        return value_str

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
                    if job_key in self.localization_dict:
                        result = self.localization_dict[job_key]

            # 4. 优先级再次：处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith(
                    "'") and ':' not in inner_text and ',' not in inner_text:
                concept_key = inner_text[1:-1]  # 去掉单引号
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]

            # 5. 优先级再次：处理 ['xxx:yyy'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)  # 去掉单引号并按冒号分割
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localization_dict:
                        result = self.localization_dict[building_key]

            # 递归处理替换后的文本中的[]替换符
            if current_depth < max_depth - 1:
                result = self.process_bracket_replacement(result, max_depth, current_depth + 1)

            # 在递归处理后处理$@替换符
            result = self.process_special_dollar_at_replacement(result)

            return result

        return re.sub(r'\[([^\]]+)\]', replace_brackets, text)

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
                    return ""

        text = re.sub(r'£([^£]+)£', replace_pounds, text)

        # 替换后继续处理其他替换符
        text, _ = self.process_dollar_replacement(text)
        text = self.process_bracket_replacement(text)

        return text

    def process_special_pound_replacement(self, text):
        """处理特殊的£xxx£替换符 - 删除特殊替换，只保留清除£blocker£字段"""
        if not text:
            return text

        # 固定清除£blocker£字段
        text = text.replace('£blocker£', '')

        return text

    def process_perk_text(self, text, modifier_key=""):
        """通用文本处理方法"""
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

        # 8. 处理换行符
        text = self.process_newlines(text)

        # 9. 最终清理空白（保留换行符）
        text = self.clean_whitespace_text_with_newlines(text)

        return text

    def process_newlines(self, text):
        """处理换行符"""
        if not text:
            return text

        # 将\\n替换为实际的换行符\n
        text = text.replace('\\n', '\n')

        return text

    def clean_whitespace_text_with_newlines(self, text):
        """清理空白文本但保留换行符"""
        if not text:
            return text

        # 按行分割，清理每行的空白，然后重新组合
        lines = text.split('\n')
        cleaned_lines = []
        for line in lines:
            # 清理每行内的空白（保留换行符）
            cleaned_line = re.sub(r'[ \t]+', ' ', line)
            cleaned_line = cleaned_line.strip()
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

    def load_perk_images(self):
        """加载飞升天赋图片映射 - 从ascension_perks.gfx文件解析图片映射"""
        self.perk_image_map = {}

        if not os.path.exists(PERK_GFX_FILE_PATH):
            print(f"飞升天赋图片映射文件不存在: {PERK_GFX_FILE_PATH}")
            return

        try:
            with open(PERK_GFX_FILE_PATH, 'r', encoding='utf-8') as f:
                content = f.read()

            # 提取所有spriteType块
            sprite_blocks = re.findall(r'spriteType\s*=\s*\{([^}]+)\}', content)

            for block in sprite_blocks:
                # 提取name值
                name_match = re.search(r'name\s*=\s*"([^"]+)"', block)
                if not name_match:
                    continue

                name = name_match.group(1)

                # 检查是否是飞升天赋图片（以GFX_ap_开头）
                if not name.startswith('GFX_ap_'):
                    continue

                # 提取飞升key（去掉GFX_前缀）
                perk_key = name[4:]  # 去掉"GFX_"前缀

                # 提取texturefile值
                texture_match = re.search(r'texturefile\s*=\s*"([^"]+)"', block)
                if not texture_match:
                    continue

                texture_path = texture_match.group(1)

                # 构建完整的DDS文件路径（基础路径 + texturefile的值）
                # 注意：texturefile中的路径使用正斜杠，需要转换为系统路径分隔符
                texture_path = texture_path.replace('/', os.path.sep)
                absolute_dds_path = os.path.join(BASE_GAME_DIR, texture_path)

                # 添加到映射字典
                self.perk_image_map[perk_key] = absolute_dds_path

            print(f"已加载 {len(self.perk_image_map)} 个飞升天赋图片映射")

        except Exception as e:
            print(f"加载飞升天赋图片映射失败: {PERK_GFX_FILE_PATH}, 错误: {e}")

    def get_perk_image_path(self, perk_key):
        """获取飞升天赋图片路径，如果是DDS则自动转换为PNG"""
        if hasattr(self, 'perk_image_map') and perk_key in self.perk_image_map:
            dds_path = self.perk_image_map[perk_key]

            # 生成PNG文件路径（在IMAGE_DIR中，使用相同的文件名但扩展名为.png）
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(PERK_IMAGE_DIR, png_filename)

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
                        self.debug_info['missing_images'].append(f"{perk_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{perk_key} -> DDS文件不存在: {dds_path}")
                    return None
        else:
            # 如果没有找到映射，记录调试信息
            if perk_key not in self.debug_info['missing_images']:
                self.debug_info['missing_images'].append(f"{perk_key} -> 无映射")
            return None

    def parse_ascension_perks_files(self):
        """解析飞升天赋文件目录"""
        print("正在解析飞升天赋文件...")

        if not os.path.exists(ASCENSION_PERKS_DIR):
            print(f"飞升天赋目录不存在: {ASCENSION_PERKS_DIR}")
            return

        for file_name in os.listdir(ASCENSION_PERKS_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(ASCENSION_PERKS_DIR, file_name)
                print(f"正在解析文件: {file_name}")
                self.parse_ascension_perks_file(file_path)

        print(f"已解析 {len(self.ascension_perks)} 个飞升天赋")

    def parse_ascension_perks_file(self, file_path):
        """解析单个飞升天赋文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 使用更强大的方法提取完整的飞升天赋定义块
            perks = self.extract_complete_perk_blocks(content)

            for perk_key, perk_content in perks:
                print(f"  解析飞升天赋: {perk_key}")
                perk_data = self.parse_single_perk(perk_key, perk_content, os.path.basename(file_path))
                if perk_data:
                    self.ascension_perks.append(perk_data)

        except Exception as e:
            print(f"解析飞升天赋文件失败: {file_path}, 错误: {e}")

    def extract_complete_perk_blocks(self, content):
        """提取完整的飞升天赋定义块，正确处理嵌套花括号"""
        blocks = []
        lines = content.split('\n')

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 寻找飞升天赋定义开始（以ap_开头）
            if line.startswith('ap_') and '=' in line and '{' in line:
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

    def parse_single_perk(self, perk_key, perk_content, source_file):
        """解析单个飞升天赋定义"""
        perk_data = {
            'key': perk_key,
            'name': '',
            'description': '',
            'requirements': [],
            'possible_conditions': [],  # 新增：存储possible块中的条件
            'effects': '',
            'modifiers': [],
            'image_path': self.get_perk_image_path(perk_key),
            'combined_output': '',  # 新增：合并输出字段
            'source': source_file
        }

        # 获取飞升天赋名称
        raw_name = self.localization_dict.get(perk_key, perk_key)
        perk_data['name'] = self.process_perk_text(raw_name)

        if perk_key not in self.localization_dict:
            self.debug_info['missing_perk_name'].add(perk_key)

        # 获取飞升天赋描述（飞升key + "_desc"后缀）
        desc_key = f"{perk_key}_desc"
        raw_description = self.localization_dict.get(desc_key, "")
        perk_data['description'] = self.process_perk_text(raw_description)

        # 提取飞升要求（potential块）
        perk_data['requirements'] = self.extract_requirements(perk_content, perk_key)

        # 提取possible块中的条件
        perk_data['possible_conditions'] = self.extract_possible_conditions(perk_content)

        # 提取飞升效果和修正（屏蔽ai_weight块中的modifier）
        self.extract_perk_effects_and_modifiers(perk_content, perk_data)

        # 生成合并输出
        perk_data['combined_output'] = self.generate_combined_output(perk_data)

        return perk_data

    def extract_requirements(self, perk_content, perk_key):
        """提取飞升要求（potential块）"""
        requirements = []

        # 查找potential块
        potential_match = re.search(r'potential\s*=\s*\{', perk_content)
        if potential_match:
            brace_start = potential_match.end() - 1
            potential_block, _ = self.extract_balanced_braces(perk_content, brace_start)

            if potential_block:
                # 提取potential块中的条件
                conditions = self.extract_conditions_from_potential(potential_block, perk_key)
                requirements.extend(conditions)

        return requirements

    def extract_possible_conditions(self, perk_content):
        """提取possible块中的条件"""
        possible_conditions = []

        # 查找possible块
        possible_match = re.search(r'possible\s*=\s*\{', perk_content)
        if possible_match:
            brace_start = possible_match.end() - 1
            possible_block, _ = self.extract_balanced_braces(perk_content, brace_start)

            if possible_block:
                # 提取possible块中的所有custom_tooltip
                possible_conditions = self.extract_custom_tooltips_from_possible(possible_block)

        return possible_conditions

    def extract_custom_tooltips_from_possible(self, possible_block):
        """从possible块中提取所有custom_tooltip的fail_text"""
        tooltips = []

        # 查找所有custom_tooltip块
        pos = 0
        while pos < len(possible_block):
            tooltip_match = re.search(r'custom_tooltip\s*=\s*\{', possible_block[pos:])
            if not tooltip_match:
                break

            brace_start = pos + tooltip_match.end() - 1
            tooltip_block, new_pos = self.extract_balanced_braces(possible_block, brace_start)

            if tooltip_block:
                # 提取fail_text
                fail_text_match = re.search(r'fail_text\s*=\s*"([^"]+)"', tooltip_block)
                if fail_text_match:
                    fail_text_key = fail_text_match.group(1)
                    # 查找本地化翻译
                    localized_text = self.localization_dict.get(fail_text_key, fail_text_key)
                    localized_text = self.process_perk_text(localized_text)
                    tooltips.append(localized_text)

                pos = new_pos
            else:
                pos += tooltip_match.end()

        return tooltips

    def extract_conditions_from_potential(self, potential_block, perk_key):
        """从potential块中提取条件"""
        conditions = []

        # 移除最外层花括号
        if potential_block.startswith('{') and potential_block.endswith('}'):
            potential_content = potential_block[1:-1].strip()
        else:
            potential_content = potential_block

        # 移除exists = owner条件
        potential_content = re.sub(r'exists\s*=\s*owner', '', potential_content)

        # 使用递归解析嵌套条件
        self.parse_nested_conditions(potential_content, conditions, None, None, 0, perk_key)

        return conditions

    def parse_nested_conditions(self, content, conditions, logic_context=None, parent_key=None, depth=0, perk_key=None):
        """解析嵌套条件"""
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

                # 特殊处理NOR块 - 修正：直接翻译NOR块中的每个条件为否定形式
                if logic_type == 'NOR':
                    nested_conditions = []
                    # 递归解析NOR块内的条件，传递in_not_block=True
                    self.parse_nested_conditions(block_content, nested_conditions, 'NOR', parent_key, depth + 1,
                                                 perk_key)

                    # 将NOR块中的每个条件转换为否定形式
                    for condition in nested_conditions:
                        # 如果条件已经是"没有"开头，保持原样
                        if condition.startswith("没有"):
                            conditions.append(condition)
                        # 如果条件是"非"开头，保持原样
                        elif condition.startswith("非"):
                            conditions.append(condition)
                        else:
                            # 其他情况添加"没有"前缀
                            conditions.append(f"没有{condition}")

                # 特殊处理NOT块 - 修正：只添加一次"没有"前缀
                elif logic_type == 'NOT':
                    nested_conditions = []
                    # 递归解析NOT块内的条件，传递in_not_block=True
                    self.parse_nested_conditions(block_content, nested_conditions, 'NOT', parent_key, depth + 1,
                                                 perk_key)

                    # NOT块中只有一个条件，直接添加"没有"前缀
                    if nested_conditions:
                        condition = nested_conditions[0]
                        # 确保不会重复添加"没有"
                        if condition.startswith("没有"):
                            conditions.append(condition)
                        else:
                            conditions.append(f"没有{condition}")

                else:
                    # 其他逻辑块正常处理
                    nested_conditions = []
                    self.parse_nested_conditions(block_content, nested_conditions, logic_type, parent_key, depth + 1,
                                                 perk_key)

                    # 根据逻辑类型添加条件
                    if logic_type == 'OR':
                        conditions.append(f"满足以下任意条件：")
                        for condition in nested_conditions:
                            conditions.append(f"  {condition}")
                    elif logic_type == 'AND':
                        conditions.append(f"同时满足以下条件：")
                        for condition in nested_conditions:
                            conditions.append(f"  {condition}")
                    elif logic_type == 'NAND':
                        conditions.append(f"不同时满足以下条件：")
                        for condition in nested_conditions:
                            conditions.append(f"  {condition}")

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
                        # 特殊处理：如果条件是关于当前飞升天赋的，则忽略
                        if key == 'has_ascension_perk' and value == perk_key:
                            # 忽略这个条件，不添加到要求列表中
                            i += 1
                            continue

                        # 根据当前是否在NOT/NOR块中决定翻译方式
                        is_in_negative_block = (logic_context in ['NOT', 'NOR'] or in_not_block)
                        condition_text = self.translate_condition(key, value, is_in_negative_block)
                        conditions.append(condition_text)

            i += 1

    def translate_condition(self, key, value, in_not_block=False):
        """翻译条件"""
        # 检查是否在条件翻译字典中
        if key in CONDITION_TRANSLATIONS:
            base_translation = CONDITION_TRANSLATIONS[key]
        else:
            # 如果不在字典中，尝试获取本地化翻译
            base_translation = self.get_localization_value(key)
            if base_translation == key:
                # 如果都找不到，使用原始键，并记录缺失的翻译
                if key not in self.debug_info['missing_condition_translations']:
                    self.debug_info['missing_condition_translations'].add(key)
                base_translation = key

        # 处理值中的双引号
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]  # 去掉双引号

        # 特殊处理host_has_dlc条件
        if key == 'host_has_dlc':
            # 使用DLC翻译字典
            if value in DLC_TRANSLATIONS:
                value = DLC_TRANSLATIONS[value]

        # 处理布尔条件（yes/no）
        if value in ['yes', 'no']:
            # 如果在否定块中，反转布尔值的含义
            if in_not_block:
                value = 'no' if value == 'yes' else 'yes'

            if value == 'yes':
                return base_translation
            else:  # value == 'no'
                return f"非{base_translation}"

        # 处理具体值条件
        else:
            # 尝试本地化值
            value_localized = self.get_localization_value(value)
            if value_localized == value:
                value_localized = value

            # 对具体值添加蓝色下划线标记
            value_with_format = f"§B{value_localized}§!"

            # 根据是否在NOT块中决定翻译
            if in_not_block:
                return f"没有{base_translation}{value_with_format}"
            else:
                return f"拥有{base_translation}{value_with_format}"

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

    def extract_perk_effects_and_modifiers(self, perk_content, perk_data):
        """提取飞升效果和修正（屏蔽ai_weight块中的modifier）"""
        # 首先移除ai_weight块，确保不提取其中的modifier
        content_without_ai_weight = self.remove_ai_weight_block(perk_content)

        # 检查是否有tradition_swap块且满足替换条件
        has_valid_tradition_swap = self.check_and_process_tradition_swap(content_without_ai_weight, perk_data)

        if not has_valid_tradition_swap:
            # 如果没有有效的tradition_swap，提取主飞升的效果和修正
            self.extract_main_perk_effects_and_modifiers(content_without_ai_weight, perk_data)

    def remove_ai_weight_block(self, content):
        """移除ai_weight块，确保不提取其中的modifier"""
        result = content
        pos = 0

        while pos < len(result):
            ai_weight_match = re.search(r'ai_weight\s*=\s*\{', result[pos:])
            if not ai_weight_match:
                break

            brace_start = pos + ai_weight_match.end() - 1
            ai_weight_block, new_pos = self.extract_balanced_braces(result, brace_start)

            if ai_weight_block:
                # 移除这个ai_weight块
                result = result[:pos + ai_weight_match.start()] + result[new_pos:]
            else:
                pos += ai_weight_match.end()

        return result

    def check_and_process_tradition_swap(self, perk_content, perk_data):
        """检查并处理tradition_swap块"""
        tradition_swap_match = re.search(r'tradition_swap\s*=\s*\{', perk_content)
        if not tradition_swap_match:
            return False

        brace_start = tradition_swap_match.end() - 1
        tradition_swap_block, _ = self.extract_balanced_braces(perk_content, brace_start)

        if not tradition_swap_block:
            return False

        # 检查trigger条件是否只有has_xxx_dlc=yes
        trigger_match = re.search(r'trigger\s*=\s*\{', tradition_swap_block)
        if trigger_match:
            trigger_brace_start = trigger_match.end() - 1
            trigger_block, _ = self.extract_balanced_braces(tradition_swap_block, trigger_brace_start)

            if trigger_block:
                # 检查是否只有has_xxx_dlc=yes条件
                if self.has_only_dlc_condition(trigger_block):
                    # 使用tradition_swap替换主飞升效果
                    self.extract_tradition_swap_effects_and_modifiers(tradition_swap_block, perk_data)
                    return True

        return False

    def has_only_dlc_condition(self, trigger_block):
        """检查trigger块是否只有has_xxx_dlc=yes条件"""
        # 移除花括号
        if trigger_block.startswith('{') and trigger_block.endswith('}'):
            trigger_content = trigger_block[1:-1].strip()
        else:
            trigger_content = trigger_block

        # 检查是否只有has_xxx_dlc=yes
        dlc_pattern = r'has_\w+_dlc\s*=\s*yes'
        dlc_matches = re.findall(dlc_pattern, trigger_content)

        # 如果找到DLC条件，检查是否还有其他条件
        if dlc_matches:
            # 移除DLC条件后检查是否还有内容
            remaining_content = re.sub(dlc_pattern, '', trigger_content)
            remaining_content = re.sub(r'\s+', '', remaining_content)
            return remaining_content == ''

        return False

    def extract_tradition_swap_effects_and_modifiers(self, tradition_swap_block, perk_data):
        """提取tradition_swap块中的效果和修正"""
        # 提取custom_tooltip
        tooltip_match = re.search(r'custom_tooltip\s*=\s*(\w+)', tradition_swap_block)
        if tooltip_match:
            tooltip_key = tooltip_match.group(1)
            tooltip_text = self.localization_dict.get(tooltip_key, tooltip_key)
            tooltip_text = self.process_perk_text(tooltip_text)
            perk_data['effects'] = tooltip_text

        # 提取on_enabled中的custom_tooltip
        on_enabled_match = re.search(r'on_enabled\s*=\s*\{', tradition_swap_block)
        if on_enabled_match:
            on_enabled_brace_start = on_enabled_match.end() - 1
            on_enabled_block, _ = self.extract_balanced_braces(tradition_swap_block, on_enabled_brace_start)

            if on_enabled_block:
                on_enabled_tooltip_match = re.search(r'custom_tooltip\s*=\s*"([^"]+)"', on_enabled_block)
                if on_enabled_tooltip_match:
                    on_enabled_tooltip_key = on_enabled_tooltip_match.group(1)
                    on_enabled_tooltip_text = self.localization_dict.get(on_enabled_tooltip_key, on_enabled_tooltip_key)
                    on_enabled_tooltip_text = self.process_perk_text(on_enabled_tooltip_text)

                    if perk_data['effects']:
                        perk_data['effects'] += f"\n{on_enabled_tooltip_text}"
                    else:
                        perk_data['effects'] = on_enabled_tooltip_text

        # 提取modifier
        modifier_match = re.search(r'modifier\s*=\s*\{', tradition_swap_block)
        if modifier_match:
            modifier_brace_start = modifier_match.end() - 1
            modifier_block, _ = self.extract_balanced_braces(tradition_swap_block, modifier_brace_start)

            if modifier_block:
                modifiers = self.extract_modifiers_from_block(modifier_block, perk_data['key'])
                perk_data['modifiers'] = modifiers

    def extract_main_perk_effects_and_modifiers(self, perk_content, perk_data):
        """提取主飞升的效果和修正"""
        # 提取最外层的custom_tooltip
        tooltip_match = re.search(r'custom_tooltip\s*=\s*(\w+)', perk_content)
        if tooltip_match:
            tooltip_key = tooltip_match.group(1)
            tooltip_text = self.localization_dict.get(tooltip_key, tooltip_key)
            tooltip_text = self.process_perk_text(tooltip_text)
            perk_data['effects'] = tooltip_text

        # 提取on_enabled中的custom_tooltip
        on_enabled_match = re.search(r'on_enabled\s*=\s*\{', perk_content)
        if on_enabled_match:
            on_enabled_brace_start = on_enabled_match.end() - 1
            on_enabled_block, _ = self.extract_balanced_braces(perk_content, on_enabled_brace_start)

            if on_enabled_block:
                on_enabled_tooltip_match = re.search(r'custom_tooltip\s*=\s*"([^"]+)"', on_enabled_block)
                if on_enabled_tooltip_match:
                    on_enabled_tooltip_key = on_enabled_tooltip_match.group(1)
                    on_enabled_tooltip_text = self.localization_dict.get(on_enabled_tooltip_key, on_enabled_tooltip_key)
                    on_enabled_tooltip_text = self.process_perk_text(on_enabled_tooltip_text)

                    if perk_data['effects']:
                        perk_data['effects'] += f"\n{on_enabled_tooltip_text}"
                    else:
                        perk_data['effects'] = on_enabled_tooltip_text

        # 提取modifier
        modifier_match = re.search(r'modifier\s*=\s*\{', perk_content)
        if modifier_match:
            modifier_brace_start = modifier_match.end() - 1
            modifier_block, _ = self.extract_balanced_braces(perk_content, modifier_brace_start)

            if modifier_block:
                modifiers = self.extract_modifiers_from_block(modifier_block, perk_data['key'])
                perk_data['modifiers'] = modifiers

    def extract_modifiers_from_block(self, modifier_block, perk_key):
        """从modifier块中提取修正效果"""
        modifiers = []

        # 分割成行处理
        lines = modifier_block.split('\n')

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

            # 处理修正值（包括@替换符）
            processed_value = self.extract_modifier_value(value)

            # 获取修正描述
            effect_description = self.get_modifier_description(key)
            effect_description = self.process_perk_text(effect_description, key)

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

    def generate_combined_output(self, perk_data):
        """生成合并输出"""
        output = []

        # 飞升名称
        output.append(f"{perk_data['name']}：")
        output.append("")

        # 飞升描述
        if perk_data['description']:
            output.append(f"{perk_data['description']}")
            output.append("")

        # 效果
        output.append("效果：")
        if perk_data['effects']:
            effect_lines = perk_data['effects'].split('\n')
            for line in effect_lines:
                if line.strip():
                    output.append(f" {line}")

        # 修正效果
        if perk_data['modifiers']:
            for modifier in perk_data['modifiers']:
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

                output.append(f" {modifier['description']}{value_text}")

        output.append("")

        # 要求 - 先显示possible条件，再显示potential条件
        output.append("要求：")

        # 显示possible块中的条件
        if perk_data['possible_conditions']:
            for condition in perk_data['possible_conditions']:
                output.append(f" {condition}")

        # 显示potential块中的条件
        if perk_data['requirements']:
            for requirement in perk_data['requirements']:
                # 处理要求文本中的颜色标记
                requirement_text = self.process_requirement_text(requirement)
                output.append(f" {requirement_text}")

        # 如果没有任何条件
        if not perk_data['possible_conditions'] and not perk_data['requirements']:
            output.append(" 无特殊要求")

        return '\n'.join(output)

    def process_requirement_text(self, requirement_text):
        """处理要求文本中的颜色标记"""
        if not requirement_text:
            return requirement_text

        # 查找需要添加蓝色标记的具体值
        # 这些通常是条件值，如飞升天赋、国民理念、起源等
        # 我们使用正则表达式来识别这些值并添加蓝色标记

        # 匹配条件值模式
        patterns = [
            r'(拥有|没有)([^§]+)(\w+)$',  # 匹配以具体值结尾的条件
        ]

        for pattern in patterns:
            matches = re.findall(pattern, requirement_text)
            for match in matches:
                prefix = match[0]  # "拥有"或"没有"
                base = match[1]  # 基础描述
                value = match[2]  # 具体值

                # 尝试本地化值
                localized_value = self.get_localization_value(value)
                if localized_value == value:
                    localized_value = value

                # 替换为带蓝色标记的版本
                new_text = f"{prefix}{base}§B{localized_value}§!"
                requirement_text = requirement_text.replace(f"{prefix}{base}{value}", new_text)

        return requirement_text

    def process_text_with_color_markers(self, text, workbook):
        """处理文本中的颜色标记符，返回富文本格式"""
        if not text:
            black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})
            return [(black_format, "")], black_format

        # 创建颜色格式
        red_format = workbook.add_format({'font_color': 'red', 'text_wrap': True})
        green_format = workbook.add_format({'font_color': 'green', 'text_wrap': True})
        blue_format = workbook.add_format({'font_color': 'blue', 'underline': True, 'text_wrap': True})
        black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})

        # 如果没有颜色标记，直接返回黑色文本
        if '§R' not in text and '§G' not in text and '§B' not in text:
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

    def generate_excel(self):
        """生成Excel文件"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}飞升提取.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('飞升提取')

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
        worksheet.set_column('B:B', 30)  # 飞升名称
        worksheet.set_column('C:C', 30)  # 飞升KEY
        worksheet.set_column('D:D', 50)  # 飞升描述
        worksheet.set_column('E:E', 50)  # 飞升要求
        worksheet.set_column('F:F', 80)  # 飞升效果（包含修正）
        worksheet.set_column('G:G', 80)  # 飞升合并输出
        worksheet.set_column('H:H', 20)  # 来源文件

        # 设置行高，为图片留出空间
        worksheet.set_default_row(100)

        # 写入表头
        headers = ['图片', '飞升名称', '飞升KEY', '飞升描述', '飞升要求', '飞升效果', '飞升合并输出', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for perk in self.ascension_perks:
            # 插入图片
            if perk['image_path']:
                try:
                    # 插入图片，调整大小适应单元格
                    worksheet.insert_image(row, 0, perk['image_path'], {
                        'x_scale': 0.8,
                        'y_scale': 0.8,
                        'x_offset': 5,
                        'y_offset': 5,
                        'object_position': 1
                    })
                except Exception as e:
                    print(f"插入图片失败: {perk['image_path']}, 错误: {e}")
            else:
                # 如果没有图片，留空
                worksheet.write(row, 0, "无图片", cell_format)

            worksheet.write(row, 1, perk['name'], cell_format)
            worksheet.write(row, 2, perk['key'], cell_format)

            # 处理飞升描述中的颜色标记
            if perk['description']:
                rich_text_parts, default_format = self.process_text_with_color_markers(perk['description'], workbook)
                if len(rich_text_parts) == 1:
                    # 只有一个部分，直接写入
                    format_obj, text = rich_text_parts[0]
                    worksheet.write(row, 3, text, format_obj)
                else:
                    # 构建富文本
                    rich_text_args = []
                    for format_obj, text in rich_text_parts:
                        rich_text_args.append(format_obj)
                        rich_text_args.append(text)
                    worksheet.write_rich_string(row, 3, *rich_text_args)
            else:
                worksheet.write(row, 3, perk['description'], description_format)

            # 处理飞升要求 - 修复：应用颜色标记处理
            # 合并possible条件和potential条件
            all_requirements = []

            # 添加possible条件
            if perk['possible_conditions']:
                all_requirements.extend(perk['possible_conditions'])

            # 添加potential条件
            if perk['requirements']:
                for requirement in perk['requirements']:
                    # 处理要求文本中的颜色标记
                    requirement_text = self.process_requirement_text(requirement)
                    all_requirements.append(requirement_text)

            requirements_text = '\n'.join(all_requirements) if all_requirements else "无特殊要求"

            # 处理飞升要求中的颜色标记
            if requirements_text and requirements_text != "无特殊要求":
                rich_text_parts, default_format = self.process_text_with_color_markers(requirements_text, workbook)
                if len(rich_text_parts) == 1:
                    # 只有一个部分，直接写入
                    format_obj, text = rich_text_parts[0]
                    worksheet.write(row, 4, text, format_obj)
                else:
                    # 构建富文本
                    rich_text_args = []
                    for format_obj, text in rich_text_parts:
                        rich_text_args.append(format_obj)
                        rich_text_args.append(text)
                    # 确保rich_text_args不为空
                    if rich_text_args:
                        worksheet.write_rich_string(row, 4, *rich_text_args)
                    else:
                        worksheet.write(row, 4, requirements_text, description_format)
            else:
                worksheet.write(row, 4, requirements_text, description_format)

            # 构建合并的效果文本（飞升效果和飞升修正）
            combined_effect_text = ""

            # 添加飞升效果
            if perk['effects']:
                combined_effect_text += f"{perk['effects']}\n\n"

            # 添加飞升修正
            if perk['modifiers']:
                combined_effect_text += "修正效果：\n"
                for modifier in perk['modifiers']:
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

                    combined_effect_text += f" {modifier['description']}{value_text}\n"

            # 处理合并文本中的颜色标记
            if combined_effect_text:
                rich_text_parts, default_format = self.process_text_with_color_markers(combined_effect_text, workbook)
                if len(rich_text_parts) == 1:
                    # 只有一个部分，直接写入
                    format_obj, text = rich_text_parts[0]
                    worksheet.write(row, 5, text, format_obj)
                else:
                    # 构建富文本
                    rich_text_args = []
                    for format_obj, text in rich_text_parts:
                        rich_text_args.append(format_obj)
                        rich_text_args.append(text)
                    # 确保rich_text_args不为空
                    if rich_text_args:
                        worksheet.write_rich_string(row, 5, *rich_text_args)
                    else:
                        worksheet.write(row, 5, combined_effect_text, description_format)
            else:
                worksheet.write(row, 5, "", description_format)

            # 处理飞升合并输出中的颜色标记
            if perk['combined_output']:
                rich_text_parts, default_format = self.process_text_with_color_markers(perk['combined_output'],
                                                                                       workbook)
                if len(rich_text_parts) == 1:
                    # 只有一个部分，直接写入
                    format_obj, text = rich_text_parts[0]
                    worksheet.write(row, 6, text, format_obj)
                else:
                    # 构建富文本
                    rich_text_args = []
                    for format_obj, text in rich_text_parts:
                        rich_text_args.append(format_obj)
                        rich_text_args.append(text)
                    # 确保rich_text_args不为空
                    if rich_text_args:
                        worksheet.write_rich_string(row, 6, *rich_text_args)
                    else:
                        worksheet.write(row, 6, perk['combined_output'], description_format)
            else:
                worksheet.write(row, 6, "", description_format)

            worksheet.write(row, 7, perk['source'], cell_format)

            row += 1

        workbook.close()
        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息 - 新增缺失翻译记录"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}飞升调试信息.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 飞升天赋调试信息 ===\n\n")

            if self.debug_info['missing_perk_name']:
                f.write("=== 缺失飞升名称 ===\n")
                for info in sorted(self.debug_info['missing_perk_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_tooltip']:
                f.write("=== 缺失飞升效果 ===\n")
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

            # 新增：缺失的条件翻译
            if self.debug_info['missing_condition_translations']:
                f.write("=== 缺失条件翻译 ===\n")
                for info in sorted(self.debug_info['missing_condition_translations']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 新增：缺失的修正翻译
            if self.debug_info['missing_modifier_translations']:
                f.write("=== 缺失修正翻译 ===\n")
                for info in sorted(self.debug_info['missing_modifier_translations']):
                    f.write(f"{info}\n")
                f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取飞升天赋信息...")

        self.load_localization_files()
        self.load_scripted_variables()
        self.load_perk_images()  # 加载图片映射
        self.parse_ascension_perks_files()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = AscensionPerkExtractor()
    extractor.run()