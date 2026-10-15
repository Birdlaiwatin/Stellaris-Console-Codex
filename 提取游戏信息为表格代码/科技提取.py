import os
import re
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS

# 固定£替换符翻译字典
SPECIAL_POUND_TRANSLATIONS = {
    'society': '社会学研究',
    'physics': '物理学研究',
    'engineering': '工程学研究',
    'trade_value': '贸易'
}


class TechnologyExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_variables_dict = {}
        self.scripted_loc_dict = {}
        self.technologies = []
        self.insensitive_localization_dict = {}

        # 调试信息
        self.debug_info = {
            'missing_tech_name': set(),
            'missing_scripted_variable': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
            'max_recursion_depth_exceeded': [],
            'debug_technologies': {},
            'missing_images': []  # 新增：缺失图片信息
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(TECHNO_IMAGES_DIR, exist_ok=True)

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

    def debug_log(self, tech_key, message):
        """记录特定科技的调试信息"""
        if tech_key in ['tech_arcane_deciphering', 'tech_archeology_lab_ancrel', 'tech_archaeostudies']:
            if tech_key not in self.debug_info['debug_technologies']:
                self.debug_info['debug_technologies'][tech_key] = []
            self.debug_info['debug_technologies'][tech_key].append(message)
            print(f"[DEBUG {tech_key}] {message}")

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

        # 从technology目录加载
        if os.path.exists(TECHNOLOGY_DIR):
            for file_name in os.listdir(TECHNOLOGY_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(TECHNOLOGY_DIR, file_name)
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
                        print(f"加载科技文件中的脚本变量失败: {file_path}, 错误: {e}")

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
                self.debug_info['missing_scripted_variable'].add(var_name)
                return value_str

        # 处理$$替换符
        if isinstance(value_str, str) and '$' in value_str:
            value_str, _ = self.process_dollar_replacement(value_str)

        return value_str

    def get_modifier_description(self, modifier_key):
        """获取修正效果的本地化描述"""
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

        for pound_key, translation in SPECIAL_POUND_TRANSLATIONS.items():
            text = text.replace(f'£{pound_key}£', translation)

        return text

    def process_tech_name(self, text):
        """专门处理科技名称"""
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

    def process_localization_text(self, text, modifier_key=""):
        """处理本地化文本中的替换符"""
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

    def get_tech_localization(self, tech_key):
        """获取科技的本地化名称"""
        # 尝试直接查找
        if tech_key in self.localization_dict:
            return self.process_tech_name(self.localization_dict[tech_key])

        # 尝试查找tech_前缀
        tech_name_key = f"{tech_key}_name"
        if tech_name_key in self.localization_dict:
            return self.process_tech_name(self.localization_dict[tech_name_key])

        # 如果都找不到，返回原始键
        self.debug_info['missing_tech_name'].add(tech_key)
        return tech_key

    def get_category_localization(self, category_key):
        """获取科技分类的本地化名称"""
        # 尝试直接查找
        if category_key in self.localization_dict:
            return self.process_localization_text(self.localization_dict[category_key])

        # 尝试查找tech_category_前缀
        category_name_key = f"tech_category_{category_key}"
        if category_name_key in self.localization_dict:
            return self.process_localization_text(self.localization_dict[category_name_key])

        # 如果都找不到，返回原始键
        return category_key

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

    def get_tech_image_path(self, tech_key, tech_content):
        """获取科技图片路径，如果是DDS则自动转换为PNG"""
        icon_path = None

        # 优化正则表达式，避免匹配到inherit_icon
        icon_match = re.search(r'(?<!inherit_)icon\s*=\s*"([^"]+)"', tech_content)
        if icon_match:
            icon_path = icon_match.group(1)
        else:
            icon_match = re.search(r'(?<!inherit_)icon\s*=\s*([^\s{]+)', tech_content)
            if icon_match:
                icon_path = icon_match.group(1)

        if icon_path:
            # 先尝试直接使用icon路径
            dds_path = os.path.join(BASE_GAME_DIR, icon_path.replace('/', os.path.sep))
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(TECHNO_IMAGES_DIR, png_filename)

            if os.path.exists(png_path):
                return png_path
            else:
                if os.path.exists(dds_path):
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path
                    else:
                        self.debug_info['missing_images'].append(f"{tech_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    # 如果找不到对应图片，尝试使用TECHNO_GFX_FILE_PATH + icon值 + .dds后缀
                    alternative_dds_path = os.path.join(TECHNO_GFX_FILE_PATH, f"{icon_path}.dds")
                    if os.path.exists(alternative_dds_path):
                        success = self.convert_dds_to_png(alternative_dds_path, png_path)
                        if success:
                            return png_path
                        else:
                            self.debug_info['missing_images'].append(f"{tech_key} -> 转换失败: {icon_path}.dds")
                            return None
                    else:
                        self.debug_info['missing_images'].append(f"{tech_key} -> DDS文件不存在: {dds_path} 和 {alternative_dds_path}")
                        return None

        # 如果没有找到icon，尝试使用默认路径
        default_dds_path = os.path.join(TECHNO_GFX_FILE_PATH, f"{tech_key}.dds")
        png_filename = f"{tech_key}.png"
        png_path = os.path.join(TECHNO_IMAGES_DIR, png_filename)

        if os.path.exists(png_path):
            return png_path
        else:
            if os.path.exists(default_dds_path):
                success = self.convert_dds_to_png(default_dds_path, png_path)
                if success:
                    return png_path
                else:
                    self.debug_info['missing_images'].append(f"{tech_key} -> 转换失败: {tech_key}.dds")
                    return None
            else:
                self.debug_info['missing_images'].append(f"{tech_key} -> 默认DDS文件不存在: {default_dds_path}")
                return None

    def parse_technology_file(self, file_path):
        """解析单个科技文件"""
        file_name = os.path.basename(file_path)
        print(f"正在解析文件: {file_name}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 处理inline_script块
            content = self.process_inline_script(content)

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 使用更强大的方法提取完整的科技定义
            technologies = self.extract_complete_blocks(content)

            for tech_key, tech_content in technologies:
                print(f"  解析科技: {tech_key}")
                self.parse_single_technology(tech_key, tech_content, file_name)

        except Exception as e:
            print(f"解析文件失败: {file_path}, 错误: {e}")

    def extract_complete_blocks(self, content):
        """提取完整的科技定义块，正确处理嵌套花括号"""
        blocks = []
        lines = content.split('\n')

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 寻找科技定义开始
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

    def parse_single_technology(self, tech_key, tech_content, source_file):
        """解析单个科技定义"""
        tech_data = {
            'key': tech_key,
            'name': '',
            'cost': '',
            'area': '',
            'category': '',
            'tier': '',
            'weight': '',
            'prerequisites': [],
            'modifiers': [],
            'other_names': [],  # 新增：其他名称
            'attributes': '',  # 新增：属性
            'description': '',  # 新增：科技描述
            'image': f"{tech_key}.png",  # 新增：科技图片
            'image_path': '',  # 新增：图片完整路径
            'source': source_file,
            'parse_order': len(self.technologies)  # 新增：记录解析顺序
        }

        # 获取科技名称
        tech_data['name'] = self.get_tech_localization(tech_key)

        # 获取图片路径
        tech_data['image_path'] = self.get_tech_image_path(tech_key, tech_content)

        # 提取基础信息
        self.extract_basic_info(tech_content, tech_data)

        # 提取前置科技
        self.extract_prerequisites(tech_content, tech_data)

        # 提取修正效果
        self.extract_modifiers(tech_content, tech_data)

        # 提取technology_swap
        self.extract_technology_swap(tech_content, tech_data)

        # 提取属性
        self.extract_attributes(tech_content, tech_data)

        # 提取科技描述
        self.extract_tech_description(tech_content, tech_data)

        print(f"    科技 {tech_key} 的修正数量: {len(tech_data['modifiers'])}")
        self.technologies.append(tech_data)

    def extract_basic_info(self, tech_content, tech_data):
        """提取科技的基本信息"""
        # 提取cost - 修复版本，处理嵌套结构
        tech_data['cost'] = self.extract_cost_value(tech_content, tech_data['key'])

        # 提取area
        area_match = re.search(r'area\s*=\s*(\w+)', tech_content)
        if area_match:
            area_value = area_match.group(1)
            tech_data['area'] = area_value
            self.debug_log(tech_data['key'], f"提取到区域: {area_value} -> {tech_data['area']}")

        # 提取category
        category_match = re.search(r'category\s*=\s*\{([^}]+)\}', tech_content)
        if category_match:
            category_value = category_match.group(1).strip()
            tech_data['category'] = self.get_category_localization(category_value)
            self.debug_log(tech_data['key'], f"提取到分类: {category_value} -> {tech_data['category']}")

        # 提取tier - 修复版本，添加t前缀
        tier_match = re.search(r'tier\s*=\s*([^\s#]+)', tech_content)
        if tier_match:
            tier_value = tier_match.group(1)
            processed_tier = self.extract_modifier_value(tier_value)
            # 添加t前缀并去除小数部分
            if processed_tier:
                try:
                    if float(processed_tier) == int(float(processed_tier)):
                        tech_data['tier'] = f"t{int(float(processed_tier))}"
                    else:
                        tech_data['tier'] = f"t{processed_tier}"
                except (ValueError, TypeError):
                    tech_data['tier'] = f"t{processed_tier}"
            self.debug_log(tech_data['key'], f"提取到等级: {tier_value} -> {tech_data['tier']}")

        # 提取weight - 修复版本，去除小数部分
        tech_data['weight'] = self.extract_weight(tech_content, tech_data['key'])

    def extract_cost_value(self, tech_content, tech_key):
        """提取cost值，处理嵌套结构"""
        # 查找cost
        cost_match = re.search(r'cost\s*=\s*\{', tech_content)
        if cost_match:
            # 处理嵌套的cost结构
            brace_start = cost_match.end() - 1
            cost_block, _ = self.extract_balanced_braces(tech_content, brace_start)
            if cost_block:
                # 在cost块中查找factor
                factor_match = re.search(r'factor\s*=\s*([^\s#]+)', cost_block)
                if factor_match:
                    cost_value = factor_match.group(1)
                    processed_cost = self.extract_modifier_value(cost_value)
                    # 去除小数部分
                    if processed_cost:
                        try:
                            if float(processed_cost) == int(float(processed_cost)):
                                processed_cost = int(float(processed_cost))
                        except (ValueError, TypeError):
                            pass
                    self.debug_log(tech_key, f"从cost块中提取到花费: {cost_value} -> {processed_cost}")
                    return processed_cost
        else:
            # 处理简单的cost
            cost_match_simple = re.search(r'cost\s*=\s*([^\s#]+)', tech_content)
            if cost_match_simple:
                cost_value = cost_match_simple.group(1)
                processed_cost = self.extract_modifier_value(cost_value)
                # 去除小数部分
                if processed_cost:
                    try:
                        if float(processed_cost) == int(float(processed_cost)):
                            processed_cost = int(float(processed_cost))
                    except (ValueError, TypeError):
                        pass
                self.debug_log(tech_key, f"提取到花费: {cost_value} -> {processed_cost}")
                return processed_cost
        return ""

    def extract_weight(self, tech_content, tech_key):
        """提取weight值 - 修复版本，去除小数部分"""
        # 使用简单的正则表达式查找weight
        weight_match = re.search(r'weight\s*=\s*([^\s#\{]+)', tech_content)
        if weight_match:
            weight_value = weight_match.group(1)
            processed_weight = self.extract_modifier_value(weight_value)
            # 去除小数部分
            if processed_weight:
                try:
                    if float(processed_weight) == int(float(processed_weight)):
                        processed_weight = int(float(processed_weight))
                except (ValueError, TypeError):
                    pass
            self.debug_log(tech_key, f"提取到权重: {weight_value} -> {processed_weight}")
            return processed_weight

        return ""

    def extract_prerequisites(self, tech_content, tech_data):
        """提取前置科技"""
        prereq_match = re.search(r'prerequisites\s*=\s*\{([^}]+)\}', tech_content)
        if prereq_match:
            prereq_content = prereq_match.group(1)
            # 提取所有引号内的科技key
            prereq_keys = re.findall(r'"([^"]+)"', prereq_content)

            for prereq_key in prereq_keys:
                prereq_name = self.get_tech_localization(prereq_key)
                tech_data['prerequisites'].append(prereq_name)
                self.debug_log(tech_data['key'], f"提取到前置科技: {prereq_key} -> {prereq_name}")

    def extract_modifiers(self, tech_content, tech_data):
        """提取修正效果 - 修复版本，只提取最外层的modifier"""
        # 使用逐行分析的方法，跟踪花括号层级
        lines = tech_content.split('\n')
        brace_level = 0
        in_outer_modifier = False
        current_modifier_block = []

        for line in lines:
            line = line.strip()

            # 统计花括号
            for char in line:
                if char == '{':
                    brace_level += 1
                elif char == '}':
                    brace_level -= 1

            # 检查是否进入最外层的modifier块
            if brace_level == 1 and 'modifier' in line and '=' in line and '{' in line:
                # 确保不是weight_modifier或其他嵌套的modifier
                if not any(keyword in line for keyword in ['weight_modifier', 'ai_weight', 'technology_swap']):
                    in_outer_modifier = True
                    self.debug_log(tech_data['key'], f"进入最外层modifier块: {line}")
                    continue

            # 如果在最外层的modifier块中
            if in_outer_modifier:
                current_modifier_block.append(line)

                # 如果modifier块结束
                if brace_level == 0 and in_outer_modifier:
                    modifier_block = '\n'.join(current_modifier_block)
                    self.debug_log(tech_data['key'], f"提取到完整的最外层modifier块: {modifier_block}")

                    # 处理modifier块
                    modifiers = self.extract_modifiers_from_block_content(modifier_block, tech_data['key'])
                    tech_data['modifiers'].extend(modifiers)

                    # 重置状态
                    in_outer_modifier = False
                    current_modifier_block = []

    def extract_modifiers_from_block_content(self, block_content, tech_key):
        """从块内容中提取修正效果 - 修复版本，处理custom_tooltip和description"""
        modifiers = []

        # 检查是否有custom_tooltip
        custom_tooltip_match = re.search(r'custom_tooltip\s*=\s*(\w+)', block_content)
        if custom_tooltip_match:
            tooltip_key = custom_tooltip_match.group(1)
            tooltip_localized = self.get_tech_localization(tooltip_key)

            # 如果找到了custom_tooltip，只输出这个tooltip，不提取其他修正
            modifiers.append({
                'key': 'custom_tooltip',
                'value': '',
                'formatted_value': '',
                'description': tooltip_localized,
            })
            self.debug_log(tech_key, f"找到custom_tooltip: {tooltip_key} -> {tooltip_localized}")
            return modifiers

        # 检查是否有description
        description_match = re.search(r'description\s*=\s*(\w+)', block_content)
        if description_match:
            description_key = description_match.group(1)
            description_localized = self.get_tech_localization(description_key)

            # 如果找到了description，只输出这个description，不提取其他修正
            modifiers.append({
                'key': 'description',
                'value': '',
                'formatted_value': '',
                'description': description_localized,
            })
            self.debug_log(tech_key, f"找到description: {description_key} -> {description_localized}")
            return modifiers

        # 如果没有custom_tooltip或description，则正常提取其他修正
        # 先移除description_parameters块
        cleaned_content = self.remove_description_parameters(block_content)

        lines = cleaned_content.split('\n')

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

            # 跳过明显的条件语句和不需要的键
            if any(keyword in line for keyword in
                   ['potential', 'trigger', 'owner', 'exists', 'OR', 'AND', 'NOT', 'NOR',
                    'show_only_custom_tooltip']):
                continue

            # 分割键值对
            parts = line.split('=', 1)
            if len(parts) != 2:
                continue

            key = parts[0].strip()
            value = parts[1].strip()

            # 跳过不需要的键
            if key in ['potential', 'icon', 'icon_frame', 'trigger', 'owner', 'exists',
                       'custom_tooltip', 'description', 'show_only_custom_tooltip', 'description_parameters']:
                continue

            # 跳过以条件关键词开头的键
            if any(key.startswith(prefix) for prefix in ['has_', 'is_', 'can_']):
                continue

            # 处理修正值（包括@替换符和$$替换符）
            processed_value = self.extract_modifier_value(value)
            self.debug_log(tech_key, f"处理修正: {key} = {value} -> {processed_value}")

            effect_description = self.get_modifier_description(key)
            self.debug_log(tech_key, f"修正描述: {key} -> {effect_description}")

            # 如果描述与键相同，说明没有找到翻译，但我们仍然保留
            if effect_description == key:
                self.debug_log(tech_key, f"未找到修正翻译: {key}")

            effect_description = self.process_localization_text(effect_description, key)
            self.debug_log(tech_key, f"处理后修正描述: {effect_description}")

            # 格式化修正值
            formatted_value = self.format_modifier_value(key, processed_value)

            modifiers.append({
                'key': key,
                'value': processed_value,
                'formatted_value': formatted_value,
                'description': effect_description,
            })

        return modifiers

    def remove_description_parameters(self, block_content):
        """移除description_parameters块及其内容"""
        # 查找description_parameters块
        desc_params_match = re.search(r'description_parameters\s*=\s*\{', block_content)
        if desc_params_match:
            brace_start = desc_params_match.end() - 1
            desc_params_block, _ = self.extract_balanced_braces(block_content, brace_start)

            if desc_params_block:
                # 从原内容中移除description_parameters块
                block_content = block_content.replace(desc_params_block, '')

        return block_content

    def extract_technology_swap(self, tech_content, tech_data):
        """提取technology_swap信息"""
        swap_match = re.search(r'technology_swap\s*=\s*\{', tech_content)
        if swap_match:
            brace_start = swap_match.end() - 1
            swap_block, _ = self.extract_balanced_braces(tech_content, brace_start)

            if swap_block:
                self.debug_log(tech_data['key'], f"找到technology_swap块: {swap_block[:100]}...")

                # 提取name
                name_match = re.search(r'name\s*=\s*(\w+)', swap_block)
                if name_match:
                    swap_name = name_match.group(1)
                    swap_localized_name = self.get_tech_localization(swap_name)

                    # 如果其他名称与科技名称不同，才添加
                    if swap_localized_name != tech_data['name']:
                        # 提取trigger条件
                        trigger_text = self.extract_swap_trigger(swap_block)

                        other_name_info = f"（{trigger_text}）{swap_localized_name}"
                        tech_data['other_names'].append(other_name_info)
                        self.debug_log(tech_data['key'], f"提取到其他名称: {other_name_info}")

    def extract_swap_trigger(self, swap_block):
        """提取technology_swap中的trigger条件"""
        trigger_match = re.search(r'trigger\s*=\s*\{', swap_block)
        if trigger_match:
            brace_start = trigger_match.end() - 1
            trigger_block, _ = self.extract_balanced_braces(swap_block, brace_start)

            if trigger_block:
                # 简化处理，只提取关键条件
                # 这里可以根据需要扩展更复杂的条件解析
                if 'country_uses_bio_ships = yes' in trigger_block:
                    return "使用生物舰船的国家"
                elif 'is_wilderness_empire = yes' in trigger_block:
                    return "生体荒野国家"
                # 可以添加更多条件翻译...

        return "特定条件"

    def extract_attributes(self, tech_content, tech_data):
        """提取科技属性 - 简化版本"""
        attributes = []

        # 1. 检查is_insight
        if re.search(r'is_insight\s*=\s*yes', tech_content):
            attributes.append("土著")

        # 2. 检查levels（只要存在就标记为循环）
        if re.search(r'levels\s*=', tech_content) and "土著" not in attributes:
            attributes.append("循环")

        # 3. 检查is_dangerous
        if re.search(r'is_dangerous\s*=\s*yes', tech_content) and "土著" not in attributes and "循环" not in attributes:
            attributes.append("危险")

        # 4. 检查is_rare
        if re.search(r'is_rare\s*=\s*yes',
                     tech_content) and "土著" not in attributes and "循环" not in attributes and "危险" not in attributes:
            attributes.append("稀有")

        # 5. 如果以上条件都不满足，标记为普通
        if not attributes:
            attributes.append("普通")

        # 按优先级取第一个属性
        tech_data['attributes'] = attributes[0] if attributes else "普通"
        self.debug_log(tech_data['key'], f"提取到属性: {tech_data['attributes']}")

    def extract_tech_description(self, tech_content, tech_data):
        """提取科技描述"""
        prereq_match = re.search(r'prereqfor_desc\s*=\s*\{', tech_content)
        if prereq_match:
            brace_start = prereq_match.end() - 1
            prereq_block, _ = self.extract_balanced_braces(tech_content, brace_start)

            if prereq_block:
                # 查找title
                title_match = re.search(r'title\s*=\s*"([^"]+)"', prereq_block)
                if title_match:
                    title_key = title_match.group(1)
                    title_localized = self.get_tech_localization(title_key)
                    tech_data['description'] = title_localized
                    self.debug_log(tech_data['key'], f"提取到科技描述: {title_key} -> {title_localized}")
                    return

    def parse_all_technologies(self):
        """解析所有科技文件"""
        print("正在解析科技文件...")

        if not os.path.exists(TECHNOLOGY_DIR):
            print(f"科技目录不存在: {TECHNOLOGY_DIR}")
            return

        for file_name in os.listdir(TECHNOLOGY_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(TECHNOLOGY_DIR, file_name)
                self.parse_technology_file(file_path)

        print(f"已解析 {len(self.technologies)} 个科技")

    def sort_technologies(self):
        """按照指定分类顺序对科技进行排序，同一分类内保持提取顺序"""
        print("正在对科技进行分类排序...")

        # 定义分类顺序
        category_order = [
            "力场操控", "计算技术", "粒子物理", "军事理论", "新世界理论",
            "治国术", "生物学", "灵能理论", "远古学", "推进力学",
            "宇航技术", "工业技术", "材料科学"
        ]

        # 创建一个字典，将分类映射到排序索引
        category_index = {category: index for index, category in enumerate(category_order)}

        # 对科技进行排序：先按分类排序，同一分类内按解析顺序排序
        self.technologies.sort(key=lambda tech: (
            category_index.get(tech['category'], len(category_order)),  # 按分类排序
            tech['parse_order']  # 分类相同按解析顺序排序
        ))

    def generate_excel(self):
        """生成Excel文件"""
        print("正在生成Excel文件...")

        # 先对科技进行排序
        self.sort_technologies()

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}科技信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('科技信息')

        # 定义基础格式 - 黑体12号字
        base_format = workbook.add_format({
            'font_name': '黑体',
            'font_size': 12,
            'border': 1,
            'valign': 'vcenter'
        })

        # 定义表头格式
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': '黑体',
            'font_size': 12
        })

        # 定义描述格式（文本换行）
        description_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'text_wrap': True,
            'font_name': '黑体',
            'font_size': 12
        })

        # 定义不同区域和属性的颜色格式
        color_formats = {}

        # Physics 区域颜色
        color_formats['physics_土著'] = workbook.add_format({
            'bg_color': '#91AADF', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['physics_循环'] = workbook.add_format({
            'bg_color': '#91AADF', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['physics_危险'] = workbook.add_format({
            'bg_color': '#B5C6EA', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['physics_稀有'] = workbook.add_format({
            'bg_color': '#B5C6EA', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['physics_普通'] = workbook.add_format({
            'bg_color': '#D9E1F4', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })

        # Society 区域颜色
        color_formats['society_土著'] = workbook.add_format({
            'bg_color': '#ADD88D', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['society_循环'] = workbook.add_format({
            'bg_color': '#ADD88D', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['society_危险'] = workbook.add_format({
            'bg_color': '#C9E4B4', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['society_稀有'] = workbook.add_format({
            'bg_color': '#C9E4B4', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['society_普通'] = workbook.add_format({
            'bg_color': '#E3F2D9', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })

        # Engineering 区域颜色
        color_formats['engineering_土著'] = workbook.add_format({
            'bg_color': '#F4B382', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['engineering_循环'] = workbook.add_format({
            'bg_color': '#F4B382', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['engineering_危险'] = workbook.add_format({
            'bg_color': '#F9CBAA', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['engineering_稀有'] = workbook.add_format({
            'bg_color': '#F9CBAA', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })
        color_formats['engineering_普通'] = workbook.add_format({
            'bg_color': '#FCE4D3', 'font_name': '黑体', 'font_size': 12, 'border': 1, 'valign': 'vcenter'
        })

        # 默认格式（用于未知区域或属性）
        default_format = workbook.add_format({
            'font_name': '黑体',
            'font_size': 12,
            'border': 1,
            'valign': 'vcenter'
        })

        # 设置列宽 - 图片列放在第一列
        worksheet.set_column('A:A', 20)  # 科技图片（第一列）
        worksheet.set_column('B:B', 30)  # 科技名称
        worksheet.set_column('C:C', 30)  # 科技KEY
        worksheet.set_column('D:D', 15)  # 花费
        worksheet.set_column('E:E', 15)  # 区域
        worksheet.set_column('F:F', 20)  # 分类
        worksheet.set_column('G:G', 10)  # 等级
        worksheet.set_column('H:H', 15)  # 权重
        worksheet.set_column('I:I', 30)  # 前置科技
        worksheet.set_column('J:J', 80)  # 修正效果
        worksheet.set_column('K:K', 40)  # 其他名称
        worksheet.set_column('L:L', 10)  # 属性
        worksheet.set_column('M:M', 50)  # 科技描述
        worksheet.set_column('N:N', 20)  # 来源文件

        # 设置行高（为图片预留足够高度）
        worksheet.set_default_row(80)

        # 写入表头 - 图片列放在第一列
        headers = ['科技图片', '科技名称', '科技KEY', '花费', '区域', '分类', '等级', '权重', '前置科技', '修正效果', '其他名称',
                   '属性', '科技描述', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for tech in self.technologies:
            # 确定单元格格式
            format_key = f"{tech['area']}_{tech['attributes']}"
            cell_format = color_formats.get(format_key, default_format)

            # 描述列使用特殊的文本换行格式，但保持相同的背景色
            desc_format = workbook.add_format({
                'bg_color': cell_format.bg_color,
                'font_name': '黑体',
                'font_size': 12,
                'border': 1,
                'valign': 'vcenter',
                'text_wrap': True
            })

            # 科技图片 - 放在第一列
            if tech['image_path'] and os.path.exists(tech['image_path']):
                try:
                    # 插入图片到单元格，调整图片大小适应单元格
                    worksheet.insert_image(row, 0, tech['image_path'], {
                        'x_scale': 0.5,
                        'y_scale': 0.5,
                        'x_offset': 5,
                        'y_offset': 5,
                        'object_position': 1  # 移动和调整单元格大小时移动图片
                    })
                    worksheet.write(row, 0, "", cell_format)  # 清空单元格文本内容，但保持格式
                except Exception as e:
                    print(f"插入图片失败 {tech['image_path']}: {e}")
                    worksheet.write(row, 0, tech['image'], cell_format)
            else:
                worksheet.write(row, 0, "无图片", cell_format)

            # 其他列数据
            worksheet.write(row, 1, tech['name'], cell_format)
            worksheet.write(row, 2, tech['key'], cell_format)
            worksheet.write(row, 3, str(tech['cost']), cell_format)
            worksheet.write(row, 4, tech['area'], cell_format)
            worksheet.write(row, 5, tech['category'], cell_format)
            worksheet.write(row, 6, str(tech['tier']), cell_format)
            worksheet.write(row, 7, str(tech['weight']), cell_format)

            # 前置科技
            prereq_text = '，'.join(tech['prerequisites']) if tech['prerequisites'] else "无"
            worksheet.write(row, 8, prereq_text, cell_format)

            # 修正效果
            if tech['modifiers']:
                text_parts = []
                for modifier in tech['modifiers']:
                    text_parts.append(f"{modifier['description']}{modifier['formatted_value']}")
                worksheet.write(row, 9, '；'.join(text_parts), cell_format)
            else:
                worksheet.write(row, 9, "无修正", cell_format)

            # 其他名称
            other_names_text = '；'.join(tech['other_names']) if tech['other_names'] else "无"
            worksheet.write(row, 10, other_names_text, cell_format)

            # 属性
            worksheet.write(row, 11, tech['attributes'], cell_format)

            # 科技描述 - 使用文本换行格式
            worksheet.write(row, 12, tech['description'], desc_format)

            # 来源文件
            worksheet.write(row, 13, tech['source'], cell_format)

            row += 1

        workbook.close()
        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}科技信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 科技信息调试信息 ===\n\n")

            if self.debug_info['missing_tech_name']:
                f.write("=== 缺失科技名称 ===\n")
                for info in sorted(self.debug_info['missing_tech_name']):
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

            # 输出特定科技的调试信息
            if self.debug_info['debug_technologies']:
                f.write("=== 特定科技调试信息 ===\n\n")
                for tech_key, messages in self.debug_info['debug_technologies'].items():
                    f.write(f"科技 {tech_key}:\n")
                    for message in messages:
                        f.write(f"  - {message}\n")
                    f.write("\n")

            # 输出缺失图片信息
            if self.debug_info['missing_images']:
                f.write("=== 缺失图片信息 ===\n")
                for info in self.debug_info['missing_images']:
                    f.write(f"{info}\n")
                f.write(f"总计: {len(self.debug_info['missing_images'])} 个\n")
                f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取科技信息...")

        self.load_localization_files()
        self.load_scripted_variables()
        self.load_scripted_loc()
        self.parse_all_technologies()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = TechnologyExtractor()
    extractor.run()