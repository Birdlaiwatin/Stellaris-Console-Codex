import os
import re
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *

# 导入百分比转换字典
try:
    from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
except ImportError:
    NON_PERCENTAGE_KEYS = []
    FORCE_PERCENTAGE_KEYS = []

# 导入颜色转换字典
try:
    from 颜色转换字典 import SPECIAL_SUFFIXES
except ImportError:
    SPECIAL_SUFFIXES = []

# 导入修正翻译字典
try:
    from 修正翻译字典 import MODIFIER_TRANSLATIONS
except ImportError:
    MODIFIER_TRANSLATIONS = {}


class StellarisModifierExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_variables_dict = {}
        self.scripted_loc_dict = {}
        self.static_modifiers = []
        self.insensitive_localization_dict = {}

        # 分类调试信息
        self.debug_info = {
            'missing_modifier_name': set(),  # 缺失修正名称
            'missing_modifier_description': set(),  # 缺失修正说明
            'missing_scripted_variable': set(),  # 缺失脚本变量
            'max_recursion_depth_exceeded': [],  # 超过最大递归深度
            'missing_bracket_replacement': set(),  # 缺失[]替换符
            'missing_pound_replacement': set(),  # 缺失£xxx£替换符
            'missing_images': []  # 缺失图片
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(STATIC_MODIFIER_IMAGES_DIR, exist_ok=True)

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

    def get_modifier_image_path(self, modifier_key, modifier_content):
        """获取修正图片路径，如果是DDS则自动转换为PNG"""
        icon_path = None

        # 查找icon字段，考虑带引号和不带引号的情况
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', modifier_content)
        if icon_match:
            icon_path = icon_match.group(1)
        else:
            icon_match = re.search(r'icon\s*=\s*([^\s{]+)', modifier_content)
            if icon_match:
                icon_path = icon_match.group(1)

        if icon_path:
            # 构建完整的DDS文件路径
            dds_path = os.path.join(BASE_GAME_DIR, icon_path.replace('/', os.path.sep))
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(STATIC_MODIFIER_IMAGES_DIR, png_filename)

            if os.path.exists(png_path):
                return png_path
            else:
                if os.path.exists(dds_path):
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        self.debug_info['missing_images'].append(f"{modifier_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    self.debug_info['missing_images'].append(f"{modifier_key} -> DDS文件不存在: {dds_path}")
                    return None

        # 如果找不到icon字段，尝试使用默认路径
        default_dds_path = os.path.join(STATIC_MODIFIER_GFX_FILE_PATH, f"{modifier_key}.dds")
        png_filename = f"{modifier_key}.png"
        png_path = os.path.join(STATIC_MODIFIER_IMAGES_DIR, png_filename)

        if os.path.exists(png_path):
            return png_path
        else:
            if os.path.exists(default_dds_path):
                success = self.convert_dds_to_png(default_dds_path, png_path)
                if success:
                    return png_path
                else:
                    self.debug_info['missing_images'].append(f"{modifier_key} -> 转换失败: {modifier_key}.dds")
                    return None
            else:
                self.debug_info['missing_images'].append(f"{modifier_key} -> 默认DDS文件不存在: {default_dds_path}")
                return None

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
            return 'black'  # 无法解析的值默认为黑色

        # 特殊后缀：正数为红色，负数为绿色
        if any(key.endswith(suffix) for suffix in SPECIAL_SUFFIXES):
            return 'red' if value_float > 0 else 'green'

        # 一般规则：正数为绿色，负数为红色
        return 'green' if value_float > 0 else 'red'

    def process_per_pop_modifier(self, key, value, description):
        """处理_per_pop和_per_crime后缀的修正，计算每多少人口/犯罪度提供一个岗位"""
        try:
            value_float = float(value)
            if value_float == 0:
                return description, value, 'black'

            # 计算每多少人口/犯罪度提供一个岗位
            per_value = 1 / value_float

            # 如果per_value不是整数，保留整数位（向下取整）
            if per_value != int(per_value):
                per_value = int(per_value)  # 向下取整
            else:
                per_value = int(per_value)  # 确保是整数类型而不是浮点数

            # 确定修正类型和对应的描述
            if key.endswith('_per_pop'):
                per_type = "人口"
            elif key.endswith('_per_crime'):
                per_type = "点犯罪度"
            else:
                return description, value, 'black'

            # 使用process_bracket_replacement处理描述中的[]替换符
            processed_description = self.process_bracket_replacement(description)

            # 从处理后的描述中提取岗位名称
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
                    else:
                        job_name = job_match.group(1)
                else:
                    job_name = "岗位"

            # 构建新的描述 - 确保使用整数，并且在描述中包含"+1"
            new_description = f"每{per_value}{per_type}提供+1个{job_name}岗位"

            # 对于_per_pop和_per_crime修正，返回空字符串作为格式化值，避免重复显示"+1"
            return new_description, "", 'green'

        except (ValueError, TypeError, ZeroDivisionError):
            return description, value, 'black'

    def preprocess_localization_text(self, text):
        """预处理本地化文本：处理$$替换符并清理格式"""
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
                # 如果找不到变量，使用默认值
                if has_percent:
                    return f"{default_value}%"
                else:
                    return default_value

        text = re.sub(pattern, replace_special_dollar_at, text)

        return text

    def load_localization_files(self):
        """加载所有本地化文件并进行预处理"""
        print("正在加载本地化文件...")
        localization_files = []

        # 遍历本地化目录及其子文件夹
        for root, dirs, files in os.walk(LOCALISATION_DIR):
            for file in files:
                if file.endswith('.yml'):
                    localization_files.append(os.path.join(root, file))

        for file_path in localization_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                # 解析YAML格式的本地化文件
                lines = content.split('\n')
                current_key = None
                current_value = []

                for line in lines:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue

                    if ':' in line and not line.startswith(' '):
                        # 保存前一个键值对
                        if current_key and current_value:
                            # 去除双引号
                            value_str = ' '.join(current_value).strip()
                            # 只提取双引号内的内容，忽略#注释
                            if '"' in value_str:
                                value_str = value_str.split('"')[1] if value_str.count('"') >= 2 else value_str
                            # 预处理文本
                            value_str = self.preprocess_localization_text(value_str)
                            self.localization_dict[current_key] = value_str

                        # 开始新的键值对
                        parts = line.split(':', 1)
                        current_key = parts[0].strip()
                        current_value = [parts[1].strip()] if len(parts) > 1 else []
                    elif current_key and line:
                        current_value.append(line.strip())

                # 保存最后一个键值对
                if current_key and current_value:
                    # 去除双引号
                    value_str = ' '.join(current_value).strip()
                    # 只提取双引号内的内容，忽略#注释
                    if '"' in value_str:
                        value_str = value_str.split('"')[1] if value_str.count('"') >= 2 else value_str
                    # 预处理文本
                    value_str = self.preprocess_localization_text(value_str)
                    self.localization_dict[current_key] = value_str

            except Exception as e:
                print(f"加载本地化文件失败: {file_path}, 错误: {e}")

        # 构建大小写不敏感的本地化字典
        print("正在构建大小写不敏感的本地化字典...")
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

        # 从static_modifiers目录加载
        if os.path.exists(STATIC_MODIFIER_DIR):
            for file_name in os.listdir(STATIC_MODIFIER_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(STATIC_MODIFIER_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        # 匹配文件开头的 @variable = value 模式
                        # 查找文件开头（不在任何块内的）@变量定义
                        lines = content.split('\n')
                        for line in lines:
                            line = line.strip()
                            # 跳过注释和空行
                            if not line or line.startswith('#'):
                                continue
                            # 如果遇到第一个块定义，停止查找
                            if re.match(r'^\w+\s*=', line) and not line.startswith('@'):
                                break

                            # 匹配 @variable = value 模式
                            pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                            match = re.match(pattern, line)
                            if match:
                                var_name, var_value = match.groups()
                                self.scripted_variables_dict[var_name] = float(var_value)

                    except Exception as e:
                        print(f"加载静态修正文件中的脚本变量失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.scripted_variables_dict)} 个脚本变量")

    def load_scripted_loc(self):
        """加载脚本本地化 - 修复版本，正确处理嵌套花括号"""
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

    def process_dollar_replacement(self, text, max_depth=5, current_depth=0):
        """处理$xxx$替换符，递归查找直到没有$xxx$或达到最大深度"""
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

    def process_pound_replacement(self, text):
        """处理£xxx£替换符"""
        if not text:
            return text

        # 查找所有£xxx£替换符
        def replace_pounds(match):
            inner_text = match.group(1)
            concept_key = f"concept_{inner_text}"

            # 首先尝试查找 concept_xxx
            if concept_key in self.localization_dict:
                return self.localization_dict[concept_key]
            else:
                # 如果找不到concept_xxx，尝试直接查找xxx
                if inner_text in self.localization_dict:
                    return self.localization_dict[inner_text]
                else:
                    self.debug_info['missing_pound_replacement'].add(
                        f"概念本地化 '{concept_key}' 和 '{inner_text}' 均未找到"
                    )
                    return ""  # 如果都找不到，删除整个£xxx£

        text = re.sub(r'£([^£]+)£', replace_pounds, text)

        return text

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符，支持递归处理嵌套的[]"""
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

            # 4. 常规[]替换符处理
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

    def process_localization_text(self, text, modifier_key=""):
        """处理本地化文本中的替换符"""
        if not text:
            return text

        # 1. 处理$xxx$替换符（虽然预处理过，但可能还有遗漏）
        text, exceeded_depth = self.process_dollar_replacement(text)
        if exceeded_depth:
            self.debug_info['max_recursion_depth_exceeded'].append(
                f"修正 {modifier_key}: 超过最大递归深度，文本: {text}"
            )

        # 2. 处理特殊$@xxx|任意数值%$替换符
        text = self.process_special_dollar_at_replacement(text)

        # 3. 处理[]替换符
        text = self.process_bracket_replacement(text)

        # 4. 处理£xxx£替换符
        text = self.process_pound_replacement(text)

        # 5. 最后清除重复文本
        text = self.remove_duplicate_text(text)

        # 6. 处理换行符 - 替换为Excel可识别的换行
        text = text.replace('\\n', '\n')

        return text

    def extract_modifier_value(self, value_str):
        """提取修正值，处理@变量"""
        # 忽略#注释
        if isinstance(value_str, str) and '#' in value_str:
            value_str = value_str.split('#')[0].strip()

        # 处理特殊@[ xxx * 数值] 格式
        value_str = self.process_special_at_replacement(value_str)

        if isinstance(value_str, str) and value_str.startswith('@'):
            var_name = value_str[1:]
            if var_name in self.scripted_variables_dict:
                return self.scripted_variables_dict[var_name]
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                return value_str
        return value_str

    def get_modifier_description(self, modifier_key):
        """获取修正效果的本地化描述 - 改进版本，使用预先构建的大小写不敏感字典"""
        # 先查找自定义翻译字典
        if modifier_key in MODIFIER_TRANSLATIONS:
            return MODIFIER_TRANSLATIONS[modifier_key]

        search_key = modifier_key.lower()

        # 按照指定顺序查找
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

        # 如果都没找到，返回原始键
        self.debug_info['missing_modifier_description'].add(modifier_key)
        return modifier_key

    def parse_static_modifiers(self):
        """解析静态修正文件"""
        print("正在解析静态修正文件...")

        if not os.path.exists(STATIC_MODIFIER_DIR):
            print(f"静态修正目录不存在: {STATIC_MODIFIER_DIR}")
            return

        # 第一阶段：收集所有修正块
        raw_modifier_blocks = {}
        for file_name in os.listdir(STATIC_MODIFIER_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(STATIC_MODIFIER_DIR, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    # 移除注释
                    content = re.sub(r'#.*', '', content)
                    # 查找修正块
                    modifier_blocks = re.findall(r'(\w+)\s*=\s*\{([^}]+)\}', content)

                    for modifier_key, block_content in modifier_blocks:
                        raw_modifier_blocks[modifier_key] = {
                            'content': block_content,
                            'source': file_name
                        }

                except Exception as e:
                    print(f"解析静态修正文件失败: {file_path}, 错误: {e}")

        # 第二阶段：解析所有修正块
        self.parse_modifier_blocks_recursive(raw_modifier_blocks)

        print(f"已解析 {len(self.static_modifiers)} 个静态修正")

    def parse_modifier_blocks_recursive(self, raw_modifier_blocks):
        """递归解析修正块，处理修正引用"""
        # 用于跟踪已解析的修正，避免无限递归
        resolved_modifiers = {}

        def parse_block(modifier_key, multiplier=1, visited=None):
            """递归解析单个修正块"""
            if visited is None:
                visited = set()

            if modifier_key in visited:
                print(f"警告: 检测到循环引用: {modifier_key}")
                return [], ""  # 返回空效果和空描述

            if modifier_key not in raw_modifier_blocks:
                print(f"警告: 未找到引用的修正: {modifier_key}")
                return [], ""  # 返回空效果和空描述

            visited.add(modifier_key)

            # 如果这个修正已经解析过，直接使用缓存的结果
            if modifier_key in resolved_modifiers:
                effects, description = resolved_modifiers[modifier_key]
                # 应用乘数到效果值，但不应用到描述
                multiplied_effects = []
                for effect in effects:
                    # 复制效果
                    multiplied_effect = effect.copy()

                    # 只有当值是数值类型时才应用乘数
                    if isinstance(multiplied_effect['value'], (int, float)):
                        multiplied_effect['value'] = multiplied_effect['value'] * multiplier

                        # 重新格式化修正值
                        if multiplied_effect['key'].endswith('_per_pop') or multiplied_effect['key'].endswith(
                                '_per_crime'):
                            # 对于_per_pop和_per_crime修正，重新处理
                            new_description, new_value, value_color = self.process_per_pop_modifier(
                                multiplied_effect['key'],
                                multiplied_effect['value'],
                                multiplied_effect['description']
                            )
                            multiplied_effect['description'] = new_description
                            multiplied_effect['formatted_value'] = new_value
                            multiplied_effect['value_color'] = value_color
                        else:
                            # 对于其他修正，重新格式化值
                            multiplied_effect['formatted_value'] = self.format_modifier_value(
                                multiplied_effect['key'],
                                multiplied_effect['value']
                            )
                            multiplied_effect['value_color'] = self.get_value_color(
                                multiplied_effect['key'],
                                multiplied_effect['value']
                            )

                    multiplied_effects.append(multiplied_effect)

                # 描述直接复制，不应用乘数
                return multiplied_effects, description

            # 解析修正块
            block_data = raw_modifier_blocks[modifier_key]
            block_content = block_data['content']
            source_file = block_data['source']

            # 获取修正图片路径
            icon_path = self.get_modifier_image_path(modifier_key, block_content)

            lines = block_content.split('\n')
            effects = []
            description = ""

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # 解析键值对
                if '=' in line:
                    # 忽略#注释
                    if '#' in line:
                        line = line.split('#')[0].strip()

                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()

                    # 处理 custom_tooltip - 作为描述
                    if key == 'custom_tooltip':
                        # 清除双引号
                        custom_tooltip = value.strip('"')
                        raw_description = self.localization_dict.get(custom_tooltip, custom_tooltip)
                        description = self.process_localization_text(raw_description, modifier_key)
                        if custom_tooltip not in self.localization_dict:
                            self.debug_info['missing_modifier_description'].add(custom_tooltip)
                        continue

                    # 跳过不需要提取的键
                    if key in ['icon', 'icon_frame', 'hide_from_country_list', 'show_only_custom_tooltip', 'important',
                               'apply_modifier_to_other_planets']:
                        continue

                    # 如果这个键是另一个修正，递归解析
                    if key in raw_modifier_blocks:
                        try:
                            sub_multiplier = float(value)
                            sub_effects, sub_description = parse_block(key, sub_multiplier * multiplier, visited.copy())
                            effects.extend(sub_effects)
                            # 不再使用子修正的描述
                        except ValueError:
                            print(f"警告: 无法解析修正引用值: {key} = {value}")
                    else:
                        # 处理普通修正效果
                        processed_value = self.extract_modifier_value(value)
                        # 应用乘数，但只对数值类型的值应用乘数
                        if isinstance(processed_value, (int, float)):
                            processed_value *= multiplier

                        effect_description = self.get_modifier_description(key)
                        effect_description = self.process_localization_text(effect_description, modifier_key)

                        # 特殊处理_per_pop和_per_crime后缀的修正
                        if key.endswith('_per_pop') or key.endswith('_per_crime'):
                            new_description, new_value, value_color = self.process_per_pop_modifier(key,
                                                                                                    processed_value,
                                                                                                    effect_description)
                            effect_description = new_description
                            formatted_value = new_value
                        else:
                            # 格式化修正值
                            formatted_value = self.format_modifier_value(key, processed_value)
                            # 获取数值颜色
                            value_color = self.get_value_color(key, processed_value)

                        effects.append({
                            'key': key,
                            'value': processed_value,
                            'formatted_value': formatted_value,
                            'description': effect_description,
                            'value_color': value_color
                        })

            # 缓存解析结果
            resolved_modifiers[modifier_key] = (effects, description)

            return effects, description

        # 解析所有修正块
        for modifier_key in raw_modifier_blocks:
            if modifier_key not in resolved_modifiers:
                effects, description = parse_block(modifier_key)

                # 获取修正名称
                raw_name = self.localization_dict.get(modifier_key, modifier_key)
                name = self.process_localization_text(raw_name, modifier_key)

                # 获取修正图片路径
                icon_path = self.get_modifier_image_path(modifier_key, raw_modifier_blocks[modifier_key]['content'])

                # 不再使用修正名称作为描述的回退，保持描述为空
                self.static_modifiers.append({
                    'key': modifier_key,
                    'name': name,
                    'description': description,  # 可能为空
                    'effects': effects,
                    'source': raw_modifier_blocks[modifier_key]['source'],
                    'icon': icon_path  # 添加图标路径
                })

    def generate_excel(self):
        """生成Excel文件，使用xlsxwriter设置颜色"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}全局修正.xlsx')

        # 创建Excel工作簿
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('全局修正')

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
            'text_wrap': True  # 允许文本换行
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
        worksheet.set_column('A:A', 15)  # 修正图片
        worksheet.set_column('B:B', 30)  # 修正名称
        worksheet.set_column('C:C', 30)  # 修正KEY
        worksheet.set_column('D:D', 50)  # 修正说明
        worksheet.set_column('E:E', 80)  # 修正效果
        worksheet.set_column('F:F', 30)  # 修正来源

        # 设置行高
        worksheet.set_default_row(40)

        # 写入表头
        headers = ['图片', '修正名称', '修正KEY', '修正说明', '修正效果', '修正来源']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for modifier in self.static_modifiers:
            # 写入修正数据
            worksheet.write(row, 0, '', cell_format)  # 图片单元格占位

            # 插入图片
            if modifier['icon'] and os.path.exists(modifier['icon']):
                try:
                    worksheet.insert_image(row, 0, modifier['icon'], {
                        'x_offset': 5,
                        'y_offset': 5,
                        'x_scale': 0.8,
                        'y_scale': 0.8,
                        'object_position': 1
                    })
                except Exception as e:
                    print(f"插入图片失败 {modifier['key']}: {e}")

            worksheet.write(row, 1, modifier['name'], cell_format)
            worksheet.write(row, 2, modifier['key'], cell_format)
            worksheet.write(row, 3, modifier['description'], description_format)  # 使用允许换行的格式
            worksheet.write(row, 5, modifier['source'], cell_format)  # 修正来源

            # 构建修正效果文本 - 使用逗号连接不同条目
            effects_parts = []
            for effect in modifier['effects']:
                effect_text = f"{effect['description']}{effect['formatted_value']}"
                effects_parts.append(effect_text)

            # 使用逗号连接所有效果，去掉逗号后的空格
            full_effect_text = ','.join(effects_parts)

            # 使用富文本格式写入修正效果
            if effects_parts:
                # 创建富文本字符串
                rich_text = []

                for i, effect in enumerate(modifier['effects']):
                    if i > 0:
                        rich_text.append(',')  # 效果之间的逗号分隔，去掉空格

                    # 确保描述和数值都不是空字符串
                    description_text = effect['description'] if effect['description'] else " "
                    formatted_value = effect['formatted_value'] if effect['formatted_value'] else " "

                    # 添加描述部分（黑色）
                    rich_text.append(black_format)
                    rich_text.append(description_text)

                    # 添加数值部分（根据颜色）
                    if effect['value_color'] == 'green':
                        rich_text.append(green_format)
                    elif effect['value_color'] == 'red':
                        rich_text.append(red_format)
                    else:
                        rich_text.append(black_format)

                    # 添加数值部分
                    rich_text.append(formatted_value)

                # 写入富文本
                if rich_text:  # 确保rich_text不为空
                    worksheet.write_rich_string(row, 4, *rich_text)

            row += 1

        # 关闭工作簿
        workbook.close()
        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}全局修正调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 全局修正调试信息 ===\n\n")

            # 输出缺失修正名称
            if self.debug_info['missing_modifier_name']:
                f.write("=== 缺失修正名称 ===\n")
                for info in sorted(self.debug_info['missing_modifier_name']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 输出缺失修正说明
            if self.debug_info['missing_modifier_description']:
                f.write("=== 缺失修正说明 ===\n")
                for info in sorted(self.debug_info['missing_modifier_description']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 输出缺失脚本变量
            if self.debug_info['missing_scripted_variable']:
                f.write("=== 缺失脚本变量 ===\n")
                for info in sorted(self.debug_info['missing_scripted_variable']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 输出超过最大递归深度
            if self.debug_info['max_recursion_depth_exceeded']:
                f.write("=== 超过最大递归深度 ===\n")
                for info in self.debug_info['max_recursion_depth_exceeded']:
                    f.write(f"{info}\n")
                f.write("\n")

            # 输出缺失[]替换符
            if self.debug_info['missing_bracket_replacement']:
                f.write("=== 缺失[]替换符 ===\n")
                for info in sorted(self.debug_info['missing_bracket_replacement']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 输出缺失£xxx£替换符
            if self.debug_info['missing_pound_replacement']:
                f.write("=== 缺失£xxx£替换符 ===\n")
                for info in sorted(self.debug_info['missing_pound_replacement']):
                    f.write(f"{info}\n")
                f.write("\n")

            # 输出缺失图片信息
            if self.debug_info['missing_images']:
                f.write("=== 缺失图片信息 ===\n")
                for info in self.debug_info['missing_images']:
                    f.write(f"{info}\n")
                f.write("\n")

            # 如果没有调试信息
            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取全局修正...")

        # 按顺序执行各个步骤
        self.load_localization_files()
        self.load_scripted_variables()
        self.load_scripted_loc()
        self.parse_static_modifiers()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = StellarisModifierExtractor()
    extractor.run()