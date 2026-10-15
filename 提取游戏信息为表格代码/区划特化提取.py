import os
import re
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *


class ZoneExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_loc_dict = {}
        self.gfx_dict = {}  # 存储GFX名称到图片文件名的映射
        self.script_icon_dict = {}  # 存储脚本文件名到icon值的映射
        self.zones = []
        self.insensitive_localization_dict = {}

        # 调试信息
        self.debug_info = {
            'missing_zone_name': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
            'missing_gfx_mapping': set(),
            'missing_script_icon': set(),
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(ZONES_IMAGES_DIR, exist_ok=True)

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

    def load_gfx_mappings(self):
        """加载GFX文件中的图标映射"""
        print("正在加载GFX图标映射...")

        if not os.path.exists(ZONES_GFX_FILE_PATH):
            print(f"GFX文件不存在: {ZONES_GFX_FILE_PATH}")
            return

        try:
            with open(ZONES_GFX_FILE_PATH, 'r', encoding='utf-8') as f:
                content = f.read()

            # 提取所有spriteType块
            sprite_blocks = re.findall(r'spriteType\s*=\s*\{([^}]+)\}', content)

            for block in sprite_blocks:
                # 提取name
                name_match = re.search(r'name\s*=\s*"([^"]+)"', block)
                if not name_match:
                    continue

                name = name_match.group(1)

                # 提取texturefile或textureFile
                texture_match = re.search(r'texture[fF]ile\s*=\s*"([^"]+)"', block)
                if not texture_match:
                    continue

                texture_path = texture_match.group(1)

                # 构建完整的DDS文件路径
                texture_path = texture_path.replace('/', os.path.sep)
                absolute_dds_path = os.path.join(BASE_GAME_DIR, texture_path)

                # 添加到映射字典
                self.gfx_dict[name] = absolute_dds_path

        except Exception as e:
            print(f"加载GFX文件失败: {ZONES_GFX_FILE_PATH}, 错误: {e}")

        print(f"已加载 {len(self.gfx_dict)} 个GFX图标映射")

    def load_script_icons(self):
        """加载脚本文件中的icon值"""
        print("正在加载脚本文件中的icon值...")

        if not os.path.exists(INLINE_SCRIPTS_DIR):
            print(f"脚本目录不存在: {INLINE_SCRIPTS_DIR}")
            return

        # 遍历所有不包含modifiers字段的txt文件
        for file_name in os.listdir(INLINE_SCRIPTS_DIR):
            if not file_name.endswith('.txt') or 'modifiers' in file_name:
                continue

            file_path = os.path.join(INLINE_SCRIPTS_DIR, file_name)
            script_name = file_name[:-4]  # 去掉.txt后缀

            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                # 移除注释
                content = re.sub(r'#.*', '', content)

                # 提取icon值
                icon_match = re.search(r'icon\s*=\s*"([^"]+)"', content)
                if not icon_match:
                    icon_match = re.search(r'icon\s*=\s*(\w+)', content)

                if icon_match:
                    icon_value = icon_match.group(1)
                    self.script_icon_dict[script_name] = icon_value
                else:
                    self.debug_info['missing_script_icon'].add(f"脚本文件 {file_name} 中没有找到icon值")

            except Exception as e:
                print(f"加载脚本文件失败: {file_path}, 错误: {e}")

        print(f"已加载 {len(self.script_icon_dict)} 个脚本icon值")

    def get_image_path(self, zone_key, icon_value, script_names):
        """获取特化图片路径"""
        # 首先尝试使用特化定义中的icon值
        if icon_value and icon_value in self.gfx_dict:
            dds_path = self.gfx_dict[icon_value]

            # 生成PNG文件路径
            dds_filename = os.path.basename(dds_path)
            png_filename = os.path.splitext(dds_filename)[0] + ".png"
            png_path = os.path.join(ZONES_IMAGES_DIR, png_filename)

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
                        self.debug_info['missing_gfx_mapping'].add(
                            f"转换失败: {zone_key} -> {icon_value} -> {dds_filename}")
                else:
                    self.debug_info['missing_gfx_mapping'].add(
                        f"DDS文件不存在: {zone_key} -> {icon_value} -> {dds_path}")

            return None

        # 如果没有icon值或GFX映射中找不到，尝试使用脚本文件中的icon值
        for script_name in script_names:
            if script_name in self.script_icon_dict:
                script_icon = self.script_icon_dict[script_name]
                if script_icon in self.gfx_dict:
                    dds_path = self.gfx_dict[script_icon]

                    # 生成PNG文件路径
                    dds_filename = os.path.basename(dds_path)
                    png_filename = os.path.splitext(dds_filename)[0] + ".png"
                    png_path = os.path.join(ZONES_IMAGES_DIR, png_filename)

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
                                self.debug_info['missing_gfx_mapping'].add(
                                    f"脚本icon存在但转换失败: {script_name} -> {script_icon} -> {dds_filename}")
                        else:
                            self.debug_info['missing_gfx_mapping'].add(
                                f"脚本icon存在但DDS文件不存在: {script_name} -> {script_icon} -> {dds_path}")

        # 如果都找不到，返回None
        self.debug_info['missing_gfx_mapping'].add(
            f"无法找到图片: {zone_key} (尝试了icon: {icon_value}, 脚本: {script_names})")
        return None

    def preprocess_localization_text(self, text):
        """预处理本地化文本"""
        if not text:
            return text

        # 1. 递归处理$$替换符
        text, _ = self.process_dollar_replacement(text)

        # 2. 清除颜色标识符 §Yxxx§!
        text = re.sub(r'§[A-Z](.*?)§!', r'\1', text)

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

        if current_depth >= max_depth - 1 and re.search(r'\$[^$]+\$', text):
            return text, True

        return text, False

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符"""
        if not text or current_depth >= max_depth:
            return text

        def replace_brackets(match):
            inner_text = match.group(1)
            result = inner_text

            # 处理 [xxx.GetNamePlural] 或 [xxx.GetName]
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

            # 处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' not in inner_text:
                concept_key = inner_text[1:-1]
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            # 处理 ['xxx:yyy'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)
                if len(parts) == 2:
                    zone_key = parts[1]
                    if zone_key in self.localization_dict:
                        result = self.localization_dict[zone_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到特化本地化: {zone_key} (来自 [{inner_text}])"
                        )

            # 常规[]替换符处理
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

    def process_special_pound_replacement(self, text):
        """处理特殊的£xxx£替换符 - 固定翻译"""
        if not text:
            return text

        # 固定£替换符翻译字典
        SPECIAL_POUND_TRANSLATIONS = {
            'society': '社会学研究',
            'physics': '物理学研究',
            'engineering': '工程学研究',
            'trade_value': '贸易'
        }

        text = text.replace('£blocker£', '')

        for pound_key, translation in SPECIAL_POUND_TRANSLATIONS.items():
            text = text.replace(f'£{pound_key}£', translation)

        return text

    def process_zone_name(self, text):
        """专门处理特化名称"""
        if not text:
            return text

        # 1. 预处理：处理$$替换符和清理颜色标识符
        text = self.preprocess_localization_text(text)

        # 2. 清理空白文本
        text = self.clean_whitespace_text(text)

        # 3. 处理[]替换符
        text = self.process_bracket_replacement(text)

        # 4. 处理特殊£xxx£替换（固定翻译）
        text = self.process_special_pound_replacement(text)

        # 5. 处理常规£xxx£替换
        text = self.process_pound_replacement(text)

        # 6. 处理换行符
        text = text.replace('\\n', '\n')

        # 7. 最终清理空白
        text = self.clean_whitespace_text(text)

        return text

    def process_inline_script(self, content):
        """处理inline_script块 - 参考物种特质提取代码的实现方式"""
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

    def parse_zone_file(self, file_path):
        """解析单个特化文件"""
        file_name = os.path.basename(file_path)
        print(f"正在解析文件: {file_name}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 处理inline_script块
            content = self.process_inline_script(content)

            # 使用更强大的方法提取完整的特化定义
            zones = self.extract_complete_blocks(content)

            for zone_key, zone_content in zones:
                print(f"  解析特化: {zone_key}")
                self.parse_single_zone(zone_key, zone_content, file_name)

        except Exception as e:
            print(f"解析文件失败: {file_path}, 错误: {e}")

    def extract_complete_blocks(self, content):
        """提取完整的特化定义块，正确处理嵌套花括号"""
        blocks = []
        pos = 0

        while pos < len(content):
            # 查找任何键后跟等号和花括号的模式
            pattern = r'(\w+)\s*=\s*\{'
            match = re.search(pattern, content[pos:])

            if not match:
                break

            key = match.group(1)
            brace_start = pos + match.end() - 1
            block_content, new_pos = self.extract_balanced_braces(content, brace_start)

            if block_content:
                # 提取内容部分（去掉外层花括号）
                content_start = block_content.find('{') + 1
                content_end = block_content.rfind('}')
                block_content_clean = block_content[content_start:content_end].strip()

                blocks.append((key, block_content_clean))

            pos = new_pos

        return blocks

    def parse_single_zone(self, zone_key, zone_content, source_file):
        """解析单个特化定义"""
        # 只处理以zone_开头的键
        if not zone_key.startswith('zone_'):
            return

        zone_data = {
            'key': zone_key,
            'name': '',
            'zone_sets': [],
            'included_building_sets': [],
            'excluded_building_sets': [],
            'source': source_file,
            'icon': '',
            'script_names': [],  # 存储脚本文件名
            'image_path': ''
        }

        # 获取特化名称
        raw_name = self.localization_dict.get(zone_key, zone_key)
        zone_data['name'] = self.process_zone_name(raw_name)

        if zone_key not in self.localization_dict:
            self.debug_info['missing_zone_name'].add(zone_key)

        # 提取icon
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', zone_content)
        if not icon_match:
            icon_match = re.search(r'icon\s*=\s*(\w+)', zone_content)

        if icon_match:
            zone_data['icon'] = icon_match.group(1)

        # 提取inline_script中的script值 - 现在直接处理内容，无需单独提取
        # 查找所有inline_script块（已处理过的内容）
        script_matches = re.findall(r'script\s*=\s*([^\s\n]+)', zone_content)
        for script_value in script_matches:
            # 检查script值是否以zones/开头且不包含modifiers
            if script_value.startswith('zones/') and 'modifiers' not in script_value:
                script_name = script_value[6:]  # 去掉'zones/'前缀
                zone_data['script_names'].append(script_name)

        # 获取图片路径
        zone_data['image_path'] = self.get_image_path(zone_key, zone_data['icon'], zone_data['script_names'])

        # 提取zone_sets
        zone_sets_match = re.search(r'zone_sets\s*=\s*\{', zone_content)
        if zone_sets_match:
            brace_start = zone_sets_match.end() - 1
            zone_sets_block, _ = self.extract_balanced_braces(zone_content, brace_start)

            if zone_sets_block:
                # 提取zone_sets中的所有值
                sets_matches = re.findall(r'\b(\w+)\b', zone_sets_block)
                zone_data['zone_sets'] = [set_name for set_name in sets_matches
                                          if set_name not in ['zone_sets', '=', '{', '}']]

        # 提取included_building_sets
        included_sets_match = re.search(r'included_building_sets\s*=\s*\{', zone_content)
        if included_sets_match:
            brace_start = included_sets_match.end() - 1
            included_sets_block, _ = self.extract_balanced_braces(zone_content, brace_start)

            if included_sets_block:
                # 提取included_building_sets中的所有值
                sets_matches = re.findall(r'\b(\w+)\b', included_sets_block)
                zone_data['included_building_sets'] = [set_name for set_name in sets_matches
                                                       if set_name not in ['included_building_sets', '=', '{', '}']]

        # 提取excluded_building_sets
        excluded_sets_match = re.search(r'excluded_building_sets\s*=\s*\{', zone_content)
        if excluded_sets_match:
            brace_start = excluded_sets_match.end() - 1
            excluded_sets_block, _ = self.extract_balanced_braces(zone_content, brace_start)

            if excluded_sets_block:
                # 提取excluded_building_sets中的所有值
                sets_matches = re.findall(r'\b(\w+)\b', excluded_sets_block)
                zone_data['excluded_building_sets'] = [set_name for set_name in sets_matches
                                                       if set_name not in ['excluded_building_sets', '=', '{', '}']]

        self.zones.append(zone_data)

    def parse_all_zones(self):
        """解析所有特化文件"""
        print("正在解析特化文件...")

        if not os.path.exists(ZONES_DIR):
            print(f"特化目录不存在: {ZONES_DIR}")
            return

        for file_name in os.listdir(ZONES_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(ZONES_DIR, file_name)
                self.parse_zone_file(file_path)

        print(f"已解析 {len(self.zones)} 个特化")

    def generate_excel(self):
        """生成Excel文件"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}特化信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('特化信息')

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

        # 设置列宽和行高
        worksheet.set_column('A:A', 15)  # 图片列
        worksheet.set_column('B:B', 40)  # 特化名称
        worksheet.set_column('C:C', 40)  # 特化KEY
        worksheet.set_column('D:D', 30)  # 特化类型
        worksheet.set_column('E:E', 30)  # 允许的建筑类别
        worksheet.set_column('F:F', 30)  # 不允许的建筑类别
        worksheet.set_column('G:G', 20)  # 来源文件
        worksheet.set_default_row(80)  # 设置默认行高为80像素

        # 写入表头
        headers = ['图片', '特化名称', '特化KEY', '特化类型', '允许的建筑类别', '不允许的建筑类别', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for zone in self.zones:
            # 插入图片 - 使用xlsxwriter直接插入，无需PIL调整大小
            if zone['image_path'] and os.path.exists(zone['image_path']):
                try:
                    # 直接插入图片，使用xlsxwriter的缩放选项
                    worksheet.insert_image(f'A{row + 1}', zone['image_path'], {
                        'x_scale': 0.5,  # 根据需要调整缩放比例
                        'y_scale': 0.5,
                        'object_position': 1  # 移动但不调整大小
                    })
                except Exception as e:
                    print(f"插入图片失败: {zone['image_path']}, 错误: {e}")

            # 写入其他数据
            worksheet.write(row, 1, zone['name'], cell_format)
            worksheet.write(row, 2, zone['key'], cell_format)

            # 特化类型用逗号分隔
            zone_sets_text = ', '.join(zone['zone_sets'])
            worksheet.write(row, 3, zone_sets_text, cell_format)

            # 允许的建筑类别用逗号分隔
            included_sets_text = ', '.join(zone['included_building_sets'])
            worksheet.write(row, 4, included_sets_text, cell_format)

            # 不允许的建筑类别用逗号分隔
            excluded_sets_text = ', '.join(zone['excluded_building_sets'])
            worksheet.write(row, 5, excluded_sets_text, cell_format)

            worksheet.write(row, 6, zone['source'], cell_format)
            row += 1

        workbook.close()

        print(f"Excel文件已生成: {output_path}")

    def save_debug_info(self):
        """保存调试信息"""
        print("正在保存调试信息...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}特化信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 特化信息调试信息 ===\n\n")

            if self.debug_info['missing_zone_name']:
                f.write("=== 缺失特化名称 ===\n")
                for info in sorted(self.debug_info['missing_zone_name']):
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

            if self.debug_info['missing_gfx_mapping']:
                f.write("=== 缺失GFX映射或图片 ===\n")
                for info in sorted(self.debug_info['missing_gfx_mapping']):
                    f.write(f"{info}\n")
                f.write("\n")

            if self.debug_info['missing_script_icon']:
                f.write("=== 缺失脚本icon ===\n")
                for info in sorted(self.debug_info['missing_script_icon']):
                    f.write(f"{info}\n")
                f.write("\n")

            if not any(self.debug_info.values()):
                f.write("没有发现调试信息\n")

        print(f"调试信息已保存: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取特化信息...")

        self.load_localization_files()
        self.load_scripted_loc()
        self.load_gfx_mappings()
        self.load_script_icons()
        self.parse_all_zones()
        self.generate_excel()
        self.save_debug_info()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = ZoneExtractor()
    extractor.run()