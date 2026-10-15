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
from 颜色转换字典 import SPECIAL_SUFFIXES  # 添加导入

# 标本类型翻译字典
TYPE_TRANSLATIONS = {
    "aesthetic_wonder": "美学大观",
    "xeno_geology": "异星奇景",
    "historical_item": "星海往事"
}

# 类型排序映射
TYPE_ORDER = {
    "美学大观": 1,
    "星海往事": 2,
    "异星奇景": 3
}


class SpecimenExtractor:
    def __init__(self):
        self.localisation_dict = {}
        self.specimens = []
        self.specimen_image_map = {}

        # 调试信息 - 扩展以包含缺失图片
        self.debug_info = {
            'missing_images': [],
            'missing_modifiers': set()
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(SPECIMENS_IMAGES_DIR, exist_ok=True)

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
                    modifier_key.endswith('_uninhabitable')):
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
            if percentage == int(percentage):
                formatted = f"{int(percentage)}%"
            else:
                formatted = f"{percentage:.1f}%"
        else:
            if value_float == int(value_float):
                formatted = f"{int(value_float)}"
            else:
                formatted = f"{value_float}"

        sign = "+" if value_float >= 0 else ""
        return f"{sign}{formatted}"

    def get_value_color(self, key, value):
        """根据修正键和值确定颜色 - 按照地块提取中的方式"""
        try:
            value_float = float(value)
        except (ValueError, TypeError):
            return 'black'

        # 特殊后缀：正数为红色，负数为绿色
        if any(key.endswith(suffix) for suffix in SPECIAL_SUFFIXES):
            return 'red' if value_float > 0 else 'green'

        # 一般规则：正数为绿色，负数为红色
        return 'green' if value_float > 0 else 'red'

    def clean_localisation_value(self, value):
        """清理本地化值，移除注释和多余引号"""
        # 移除注释（#后面的所有内容）
        if '#' in value:
            value = value.split('#')[0].strip()

        # 移除引号
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]

        return value.strip()

    def preprocess_localisation(self, text):
        """预处理本地化文本 - 清除格式标记、换行符和所有空格"""
        if not text:
            return text

        # 清除$t$, $tt$, $ttt$, $tttt$等格式标记
        text = re.sub(r'\$t+\$', '', text)

        # 删除所有换行符（\n）
        text = re.sub(r'\\n', '', text)

        # 删除所有空格（包括普通空格、制表符等所有空白字符）
        text = re.sub(r'\s+', '', text)

        return text

    def format_resource_value(self, value):
        """格式化资源值，用于resources块"""
        try:
            value_float = float(value)
        except (ValueError, TypeError):
            return str(value)

        if value_float == int(value_float):
            return f"{int(value_float)}"
        else:
            return f"{value_float}"

    def parse_localisation_files(self):
        """解析本地化文件，构建翻译字典"""
        print("正在解析本地化文件...")
        localisation_dict = {}

        for file_path in Path(LOCALISATION_DIR).rglob("*.yml"):
            try:
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    for line in f:
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            # 分割键和值
                            parts = line.split(':', 1)
                            if len(parts) == 2:
                                key = parts[0].strip()
                                value = parts[1].strip()

                                # 清理值（移除注释和引号）
                                value = self.clean_localisation_value(value)

                                # 预处理本地化文本
                                value = self.preprocess_localisation(value)

                                localisation_dict[key] = value
            except Exception as e:
                print(f"解析本地化文件时出错 {file_path}: {e}")

        self.localisation_dict = localisation_dict
        print(f"已加载 {len(localisation_dict)} 个本地化条目")

    def resolve_localisation(self, value, depth=0, max_depth=10):
        """递归解析本地化值中的变量引用"""
        if depth > max_depth:
            return value  # 防止无限递归

        # 预处理文本
        value = self.preprocess_localisation(value)

        # 查找所有 $...$ 格式的变量引用
        pattern = r'\$([^$]+)\$'
        matches = re.findall(pattern, value)

        if not matches:
            return value  # 没有变量引用，直接返回

        # 替换所有变量引用
        for match in matches:
            # 递归解析引用的变量
            resolved_value = self.resolve_localisation(self.localisation_dict.get(match, match), depth + 1, max_depth)
            # 替换原字符串中的引用
            value = value.replace(f'${match}$', resolved_value)

        return value

    def find_modifier_translation(self, modifier_key, max_depth=8):
        """查找修正字段的本地化翻译，并递归解析占位符"""
        # 首先检查手动指定的修正翻译字典
        if modifier_key in MODIFIER_TRANSLATIONS:
            return MODIFIER_TRANSLATIONS[modifier_key], None

        # 尝试不同的键格式
        possible_keys = [
            f"mod_{modifier_key}",
            f"MOD_{modifier_key.upper()}",
            modifier_key
        ]

        translation = None
        for key in possible_keys:
            if key in self.localisation_dict:
                translation = self.localisation_dict[key]
                break

        if not translation:
            return None, modifier_key

        # 递归解析占位符
        return self.resolve_modifier_placeholders(translation, max_depth), None

    def deduplicate_repeated_text(self, text):
        """去除重复的文本，如"灵能使灵能使" -> "灵能使" """
        if not text or len(text) < 4:  # 如果文本太短，不需要处理
            return text

        # 第一步：先处理特定重复文本
        replacements = {
            '军用舰船舰队中': '舰队中'
        }

        result = text
        for old, new in replacements.items():
            result = result.replace(old, new)

        # 第二步：匹配连续重复的中文文本（至少2个字符）
        pattern = r'([\u4e00-\u9fa5]{2,}?)\1+'

        def remove_duplicate(match):
            # 返回单个实例
            return match.group(1)

        # 使用sub方法替换所有匹配的重复文本
        previous_result = result
        result = re.sub(pattern, remove_duplicate, result)

        # 如果进行了替换，递归处理直到没有变化
        if result != previous_result:
            return self.deduplicate_repeated_text(result)

        return result

    def resolve_modifier_placeholders(self, text, max_depth=0):
        """递归解析修正文本中的占位符"""
        if max_depth <= 0:
            return text

        # 预处理文本
        text = self.preprocess_localisation(text)

        # 解析 £...£ 格式的占位符
        pound_pattern = r'£([^£]+)£'
        pound_matches = re.findall(pound_pattern, text)

        for match in pound_matches:
            # 查找概念本地化
            concept_key = f"concept_{match}"
            concept_value = self.localisation_dict.get(concept_key, match)

            # 递归解析概念值中的占位符
            resolved_concept = self.resolve_modifier_placeholders(concept_value, max_depth - 1)

            # 替换占位符
            text = text.replace(f'£{match}£', resolved_concept)

        # 解析 $...$ 格式的占位符
        dollar_pattern = r'\$([^$]+)\$'
        dollar_matches = re.findall(dollar_pattern, text)

        for match in dollar_matches:
            # 直接查找本地化
            resolved_value = self.localisation_dict.get(match, match)

            # 递归解析值中的占位符
            resolved_value = self.resolve_modifier_placeholders(resolved_value, max_depth - 1)

            # 替换占位符
            text = text.replace(f'${match}$', resolved_value)

        # 删除重复文本
        text = self.deduplicate_repeated_text(text)

        return text

    def parse_condition(self, condition_content):
        """解析条件并返回翻译后的文本"""
        # 查找所有条件键值对
        condition_pattern = r'(\w+)\s*=\s*(\w+)'
        condition_matches = re.findall(condition_pattern, condition_content)

        conditions = []
        for condition_key, condition_value in condition_matches:
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

        return "；".join(conditions) if conditions else "特殊条件效果"

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

    def extract_image_name_from_icon(self, icon_value):
        """从icon值中提取图片名称"""
        if not icon_value:
            return None

        # 移除可能的引号
        icon_value = icon_value.strip('"\'')

        return icon_value

    def load_specimen_images(self):
        """加载标本图片映射"""
        self.specimen_image_map = {}

        if not os.path.exists(SPECIMENS_GFX_FILE_PATH):
            print(f"标本图片映射文件不存在: {SPECIMENS_GFX_FILE_PATH}")
            return

        try:
            with open(SPECIMENS_GFX_FILE_PATH, 'r', encoding='utf-8') as f:
                content = f.read()

            # 提取所有spriteType块
            sprite_blocks = re.findall(r'spriteType\s*=\s*\{([^}]*?(?:\{[^}]*\}[^}]*?)*)\}', content)

            for block in sprite_blocks:
                # 提取name值
                name_match = re.search(r'name\s*=\s*"([^"]+)"', block)
                if not name_match:
                    continue

                name = name_match.group(1)

                # 提取texturefile值（小写）
                texture_match = re.search(r'texturefile\s*=\s*"([^"]+)"', block, re.IGNORECASE)
                if not texture_match:
                    continue

                texture_path = texture_match.group(1)

                # 构建完整的DDS文件路径（基础路径 + texturefile的值）
                # 注意：texturefile中的路径使用正斜杠，需要转换为系统路径分隔符
                texture_path = texture_path.replace('/', os.path.sep)
                absolute_dds_path = os.path.join(BASE_GAME_DIR, texture_path)

                # 添加到映射字典，使用name作为键
                self.specimen_image_map[name] = absolute_dds_path

            print(f"已加载 {len(self.specimen_image_map)} 个标本图片映射")

        except Exception as e:
            print(f"加载标本图片映射失败: {SPECIMENS_GFX_FILE_PATH}, 错误: {e}")

    def get_specimen_image_path(self, specimen_key, icon_value):
        """获取标本图片路径"""
        if not icon_value:
            return None

        # 移除可能的引号
        icon_value = icon_value.strip('"\'')

        # 直接使用icon_value作为映射键
        if icon_value in self.specimen_image_map:
            dds_path = self.specimen_image_map[icon_value]

            # 生成PNG文件路径（在SPECIMENS_IMAGES_DIR中，使用相同的文件名但扩展名为.png）
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(SPECIMENS_IMAGES_DIR, png_filename)

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
                        self.debug_info['missing_images'].append(f"{specimen_key} -> 转换失败: {dds_filename}")
                        return None
                else:
                    # 记录缺失的图片
                    self.debug_info['missing_images'].append(f"{specimen_key} -> DDS文件不存在: {dds_path}")
                    return None
        else:
            # 如果没有找到映射，尝试直接查找图片文件
            image_name = self.extract_image_name_from_icon(icon_value)
            possible_extensions = ['.png', '.jpg', '.jpeg']
            for ext in possible_extensions:
                test_path = os.path.join(SPECIMENS_IMAGES_DIR, f"{image_name}{ext}")
                if os.path.exists(test_path):
                    return test_path

            # 记录缺失的图片
            self.debug_info['missing_images'].append(f"{specimen_key} -> 无映射且文件不存在: {icon_value}")
            return None

    def get_resource_localization(self, resource_key):
        """获取资源的本地化名称 - 按照地块提取中的方式"""
        # 先尝试concept_前缀
        concept_key = f"concept_{resource_key}"
        if concept_key in self.localisation_dict:
            return self.localisation_dict[concept_key]

        # 再尝试直接查找
        if resource_key in self.localisation_dict:
            return self.localisation_dict[resource_key]

        # 如果都找不到，返回原始键
        return resource_key

    def parse_specimen_files(self):
        """解析标本文件，提取标本定义"""
        print("正在解析标本文件...")
        specimens = []

        for file_path in Path(SPECIMENS_DIR).glob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                    # 移除单行注释
                    content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                    # 使用正则表达式找到所有标本定义的开始
                    pattern = r'^(\w+)\s*=\s*\{'
                    matches = list(re.finditer(pattern, content, re.MULTILINE))

                    for match in matches:
                        specimen_key = match.group(1)
                        start_pos = match.start()

                        # 提取完整的标本块（处理嵌套结构）
                        specimen_block = self.extract_complete_block(content, start_pos)

                        if not specimen_block:
                            print(f"警告: 无法找到标本 {specimen_key} 的完整定义块")
                            continue

                        # 提取品质 (RARITY)
                        rarity_match = re.search(r'RARITY\s*=\s*(\w+)', specimen_block)
                        rarity = rarity_match.group(1) if rarity_match else "unknown"

                        # 提取类型 (type)
                        type_match = re.search(r'type\s*=\s*(\w+)', specimen_block)
                        specimen_type = type_match.group(1) if type_match else "unknown"

                        # 提取图标 (icon) - 改进的提取方法
                        icon_value = None

                        # 方法1: 使用更健壮的正则表达式匹配icon字段
                        icon_patterns = [
                            r'icon\s*=\s*"([^"]*)"',  # 带双引号
                            r"icon\s*=\s*'([^']*)'",  # 带单引号
                            r'icon\s*=\s*(\S+)'  # 不带引号
                        ]

                        for pattern in icon_patterns:
                            icon_match = re.search(pattern, specimen_block)
                            if icon_match:
                                icon_value = icon_match.group(1)
                                break

                        # 方法2: 如果正则表达式失败，使用逐行搜索
                        if not icon_value:
                            lines = specimen_block.split('\n')
                            for line in lines:
                                line = line.strip()
                                if line.startswith('icon'):
                                    # 尝试提取icon值
                                    parts = re.split(r'\s*=\s*', line, 1)
                                    if len(parts) > 1:
                                        icon_value = parts[1].strip()
                                        break

                        # 获取图片路径
                        image_path = self.get_specimen_image_path(specimen_key, icon_value)

                        # 提取标本描述和详情
                        desc_short_key = f"{specimen_key}_desc_short"
                        desc_details_key = f"{specimen_key}_desc_details"

                        desc_short = self.resolve_localisation(self.localisation_dict.get(desc_short_key, ""))
                        desc_details = self.resolve_localisation(self.localisation_dict.get(desc_details_key, ""))

                        # 提取效果
                        effects = []
                        missing_modifiers = set()

                        # 提取triggered_country_modifier效果 - 修复没有potential块的情况
                        modifier_pattern = r'triggered_country_modifier\s*=\s*\{([^{}]*?(?:\{[^{}]*\}[^{}]*?)*)\}'
                        modifier_blocks = re.findall(modifier_pattern, specimen_block)

                        for modifier_block in modifier_blocks:
                            # 提取限制条件（如果有potential块）
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
                                translation, missing_key = self.find_modifier_translation(modifier_key)

                                if translation:
                                    # 使用智能格式化修正值
                                    value_str = self.format_modifier_value(modifier_key, modifier_value)
                                    value_color = self.get_value_color(modifier_key, modifier_value)

                                    # 修正：使用正确的颜色标记符格式
                                    color_code = 'G' if value_color == 'green' else 'R' if value_color == 'red' else ''
                                    if color_code:
                                        effect_text = f"{translation}§{color_code}{value_str}§!"
                                    else:
                                        effect_text = f"{translation}{value_str}"

                                    if condition:
                                        effect_text = f"（{condition}）{effect_text}"

                                    effects.append({
                                        'text': effect_text,
                                        'color': 'black'  # 文本部分保持黑色
                                    })
                                else:
                                    # 只有当修正键确实找不到本地化时才记录
                                    if missing_key and missing_key not in MODIFIER_TRANSLATIONS:
                                        missing_modifiers.add(missing_key)

                        # 提取resources效果 - 修改资源本地化查找方式
                        resources_match = re.search(r'resources\s*=\s*\{([^{}]*?(?:\{[^{}]*\}[^{}]*?)*)\}',
                                                    specimen_block)
                        if resources_match:
                            resources_block = resources_match.group(1)

                            # 提取produces块
                            produces_match = re.search(r'produces\s*=\s*\{([^{}]*)\}', resources_block)
                            if produces_match:
                                produces_content = produces_match.group(1)

                                # 提取资源类型和数值
                                resource_pattern = r'(\w+)\s*=\s*(\d+)'
                                resource_matches = re.findall(resource_pattern, produces_content)

                                for resource_key, resource_value in resource_matches:
                                    # 使用改进的资源本地化查找方式
                                    resource_translation = self.get_resource_localization(resource_key)
                                    resource_translation = self.resolve_modifier_placeholders(resource_translation)

                                    # 使用专门的资源值格式化函数，不添加符号
                                    value_str = self.format_resource_value(resource_value)
                                    # 修正：使用正确的颜色标记符格式
                                    effects.append({
                                        'text': f"每月{resource_translation}§G+{value_str}§!",
                                        'color': 'black'  # 文本部分保持黑色
                                    })

                        # 收集缺失的修正字段到全局调试信息 - 只添加真正缺失的
                        for missing_modifier in missing_modifiers:
                            # 再次验证这个修正键是否确实找不到本地化
                            translation, _ = self.find_modifier_translation(missing_modifier)
                            if not translation and missing_modifier not in MODIFIER_TRANSLATIONS:
                                self.debug_info['missing_modifiers'].add(missing_modifier)

                        specimens.append({
                            'key': specimen_key,
                            'rarity': rarity,
                            'type': specimen_type,
                            'icon': icon_value,
                            'image_path': image_path,
                            'image_name': self.extract_image_name_from_icon(icon_value) or specimen_key,
                            'desc_short': desc_short,
                            'desc_details': desc_details,
                            'effects': effects,
                            'missing_modifiers': list(missing_modifiers),
                            'original_index': len(specimens)  # 记录原始解析顺序
                        })

            except Exception as e:
                print(f"解析标本文件时出错 {file_path}: {e}")

        self.specimens = specimens
        print(f"已找到 {len(specimens)} 个标本定义")

    def generate_output(self):
        """生成输出结果 - 只生成Excel数据，不生成TXT输出"""
        excel_data = []

        # 为每个标本添加翻译后的类型
        for specimen in self.specimens:
            specimen['type_translation'] = TYPE_TRANSLATIONS.get(specimen['type'], specimen['type'])

        # 按照类型和原始解析顺序排序
        sorted_specimens = sorted(
            self.specimens,
            key=lambda x: (TYPE_ORDER.get(x['type_translation'], 99), x['original_index'])
        )

        # 生成Excel数据
        for specimen in sorted_specimens:
            # 获取标本名称翻译
            specimen_name = self.resolve_localisation(self.localisation_dict.get(specimen['key'], specimen['key']))

            # 翻译品质
            rarity_translations = {
                "common": "普通",
                "rare": "稀有",
                "epic": "史诗",
                "exceptional": "卓绝"
            }
            rarity = rarity_translations.get(specimen['rarity'], specimen['rarity'])

            # 翻译类型
            specimen_type = specimen['type_translation']

            # 获取标本描述和详情
            desc_short = specimen['desc_short'] if specimen['desc_short'] else "无描述"
            desc_details = specimen['desc_details'] if specimen['desc_details'] else "无详情"

            # 添加到Excel数据 - 使用换行符分隔效果
            excel_effects = []
            for effect in specimen['effects']:
                excel_effects.append({
                    'text': effect['text'],
                    'color': effect['color']
                })

            excel_data.append({
                "图片": specimen['image_path'] or "",  # 图片路径
                "标本名称": specimen_name,
                "标本KEY": specimen['key'],
                "标本类型": specimen_type,
                "标本品质": rarity,
                "标本描述": desc_short,
                "标本详情": desc_details,
                "标本效果": excel_effects,
                "图片名称": specimen['image_name']
            })

        return excel_data

    def generate_debug_output(self):
        """生成调试输出 - 只包含缺失的修正和缺失的图片"""
        output_lines = []

        # 添加缺失修正字段的说明 - 只输出真正缺失的
        if self.debug_info['missing_modifiers']:
            output_lines.append("以下修正找不到本地化文件：")
            for modifier in sorted(self.debug_info['missing_modifiers']):
                output_lines.append(f"  - {modifier}")
            output_lines.append("")

        # 添加缺失图片的说明
        if self.debug_info['missing_images']:
            output_lines.append("以下标本找不到对应图片：")
            # 提取标本key
            missing_specimen_keys = set()
            for info in self.debug_info['missing_images']:
                # 从信息中提取标本key
                match = re.search(r'^([^->]+) ->', info)
                if match:
                    missing_specimen_keys.add(match.group(1))

            for key in sorted(missing_specimen_keys):
                output_lines.append(f"  - {key}")
            output_lines.append("")

        # 如果没有调试信息
        if not self.debug_info['missing_modifiers'] and not self.debug_info['missing_images']:
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
        excel_file = os.path.join(OUTPUT_DIR, f"{VERSION}标本.xlsx")

        # 使用xlsxwriter创建工作簿和工作表
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet("标本数据")

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
        worksheet.set_column('B:B', 20)  # 标本名称
        worksheet.set_column('C:C', 15)  # 标本KEY
        worksheet.set_column('D:D', 10)  # 标本类型
        worksheet.set_column('E:E', 10)  # 标本品质
        worksheet.set_column('F:F', 30)  # 标本描述
        worksheet.set_column('G:G', 50)  # 标本详情
        worksheet.set_column('H:H', 40)  # 标本效果
        worksheet.set_column('I:I', 25)  # 图片名称

        # 设置行高以容纳图片和文本
        worksheet.set_default_row(80)

        # 写入表头 - 按照新顺序：图片，标本名称，标本KEY，标本类型，标本品质，标本描述，标本详情，标本效果，图片名称
        headers = ["图片", "标本名称", "标本KEY", "标本类型", "标本品质", "标本描述", "标本详情", "标本效果",
                   "图片名称"]
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
            worksheet.write(row, 1, data["标本名称"], cell_format)
            worksheet.write(row, 2, data["标本KEY"], cell_format)
            worksheet.write(row, 3, data["标本类型"], cell_format)
            worksheet.write(row, 4, data["标本品质"], cell_format)
            worksheet.write(row, 5, data["标本描述"], wrap_format)
            worksheet.write(row, 6, data["标本详情"], wrap_format)
            worksheet.write(row, 8, data["图片名称"], cell_format)

            # 处理标本效果（带颜色）
            effects = data["标本效果"]
            if effects:
                effect_text_parts = []
                for i, effect in enumerate(effects):
                    if i > 0:
                        effect_text_parts.append("\n")

                    # 处理文本中的颜色标记
                    rich_parts, _ = self.process_text_with_color_markers(effect['text'], workbook)

                    for format_obj, text in rich_parts:
                        effect_text_parts.append((format_obj, text))

                # 写入富文本
                if len(effect_text_parts) == 1 and not isinstance(effect_text_parts[0], tuple):
                    # 只有一段文本
                    worksheet.write(row, 7, effect_text_parts[0], wrap_format)
                else:
                    # 多段富文本
                    rich_text_args = []
                    for part in effect_text_parts:
                        if isinstance(part, tuple):
                            rich_text_args.extend(part)
                        else:
                            rich_text_args.append(wrap_format)
                            rich_text_args.append(part)

                    worksheet.write_rich_string(row, 7, *rich_text_args)
            else:
                worksheet.write(row, 7, "无效果", wrap_format)

        workbook.close()
        print(f"Excel文件已保存到: {excel_file}")

    def save_debug_info(self):
        """保存调试信息到TXT文件"""
        print("正在保存调试信息...")

        # 生成调试输出
        debug_lines = self.generate_debug_output()

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}标本信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 标本信息调试信息 ===\n\n")
            f.write('\n'.join(debug_lines))

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取标本信息...")

        self.parse_localisation_files()
        self.load_specimen_images()
        self.parse_specimen_files()

        print("正在生成输出...")
        excel_data = self.generate_output()

        # 不再生成详细的标本TXT文件，只保存调试信息
        # 保存调试信息到TXT文件
        self.save_debug_info()

        # 保存到Excel
        self.save_to_excel(excel_data)

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = SpecimenExtractor()
    extractor.run()