import os
import re
import xlsxwriter
import imageio.v3 as imageio

# 导入配置文件
from A配置路径 import *

# 固定£替换符翻译字典
SPECIAL_POUND_TRANSLATIONS = {
    'society': '社会学研究',
    'physics': '物理学研究',
    'engineering': '工程学研究',
    'trade_value': '贸易'
}


class BuildingExtractor:
    def __init__(self):
        self.localization_dict = {}
        self.scripted_loc_dict = {}
        self.buildings = []
        self.insensitive_localization_dict = {}

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(BUILDINGS_IMAGES_DIR, exist_ok=True)

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
                    if job_key in self.localization_dict:
                        result = self.localization_dict[job_key]

            # 2. 优先级次高：处理 ['xxx'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' not in inner_text:
                concept_key = inner_text[1:-1]  # 去掉单引号
                if concept_key in self.localization_dict:
                    result = self.localization_dict[concept_key]

            # 3. 优先级再次：处理 ['xxx:yyy'] 格式
            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)  # 去掉单引号并按冒号分割
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localization_dict:
                        result = self.localization_dict[building_key]

            # 4. 常规[]替换符处理 - 查找scripted_loc_dict
            elif inner_text in self.scripted_loc_dict:
                loc_key = self.scripted_loc_dict[inner_text]
                if loc_key in self.localization_dict:
                    result = self.localization_dict[loc_key]

            # 递归处理替换后的文本中的[]替换符
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

    def process_building_name(self, text):
        """专门处理建筑名称"""
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

        # 6. 清除重复文本
        text = self.remove_duplicate_text(text)

        # 7. 处理换行符
        text = text.replace('\\n', '\n')

        # 8. 最终清理空白
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

        # 3. 处理[]替换符
        text = self.process_bracket_replacement(text)

        # 4. 处理常规£xxx£替换
        text = self.process_pound_replacement(text)

        # 5. 清除重复文本
        text = self.remove_duplicate_text(text)

        # 6. 处理换行符
        text = text.replace('\\n', '\n')

        # 7. 最终清理空白
        text = self.clean_whitespace_text(text)

        return text

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

    def get_category_localization(self, category_key):
        """获取建筑类别的本地化名称"""
        # 尝试直接查找
        if category_key in self.localization_dict:
            return self.localization_dict[category_key]

        # 尝试查找category_前缀
        category_loc_key = f"category_{category_key}"
        if category_loc_key in self.localization_dict:
            return self.localization_dict[category_loc_key]

        # 如果都找不到，返回原始键
        return category_key

    def get_image_path(self, building_key, building_content):
        """获取建筑图片路径，如果是DDS则自动转换为PNG"""
        icon_path = None

        # 优化正则表达式，避免匹配到ruined_icon等无关字段
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', building_content)
        if icon_match:
            icon_path = icon_match.group(1)
        else:
            # 如果没有找到带双引号的icon，尝试不带双引号的，但要排除ruined_icon等
            icon_match = re.search(r'(?<!ruined_)icon\s*=\s*(\w+)', building_content)
            if icon_match:
                icon_path = icon_match.group(1)

        if icon_path:
            # 优先使用icon值查找图片
            dds_path = os.path.join(BUILDINGS_GFX_FILE_PATH, f"{icon_path}.dds")
            png_filename = f"{icon_path}.png"
            png_path = os.path.join(BUILDINGS_IMAGES_DIR, png_filename)

            if os.path.exists(png_path):
                return png_path
            else:
                if os.path.exists(dds_path):
                    success = self.convert_dds_to_png(dds_path, png_path)
                    if success:
                        return png_path

        # 如果icon图片不存在，使用建筑key查找
        default_dds_path = os.path.join(BUILDINGS_GFX_FILE_PATH, f"{building_key}.dds")
        png_filename = f"{building_key}.png"
        png_path = os.path.join(BUILDINGS_IMAGES_DIR, png_filename)

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

    def parse_building_file(self, file_path):
        """解析单个建筑文件"""
        file_name = os.path.basename(file_path)
        print(f"正在解析文件: {file_name}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 移除注释
            content = re.sub(r'#.*', '', content)

            # 使用更强大的方法提取完整的建筑定义
            buildings = self.extract_complete_blocks(content)

            for building_key, building_content in buildings:
                print(f"  解析建筑: {building_key}")
                self.parse_single_building(building_key, building_content, file_name)

        except Exception as e:
            print(f"解析文件失败: {file_path}, 错误: {e}")

    def extract_complete_blocks(self, content):
        """提取完整的建筑定义块，正确处理嵌套花括号"""
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

    def is_building_definition(self, building_content):
        """判断是否为建筑定义"""
        # 检查是否包含建筑相关属性
        building_indicators = [
            r'category\s*=\s*\w+',
            r'building_sets\s*=\s*\{',
            r'base_buildtime\s*=\s*@?\w+',
            r'resources\s*=\s*\{',
            r'planet_limit\s*=\s*\d+',
            r'empire_limit\s*=\s*\d+',
            r'upgrades\s*=\s*\{',
            r'prerequisites\s*=\s*\{'
        ]

        for indicator in building_indicators:
            if re.search(indicator, building_content):
                return True
        return False

    def parse_single_building(self, building_key, building_content, source_file):
        """解析单个建筑定义"""
        # 检查是否为建筑定义
        if not self.is_building_definition(building_content):
            return

        building_data = {
            'key': building_key,
            'name': '',
            'category': '',
            'building_sets': [],
            'icon': '',
            'image_path': '',
            'source': source_file
        }

        # 获取建筑名称 - 使用专门的建筑名称处理
        # 首先尝试使用原始key查找本地化
        raw_name = self.localization_dict.get(building_key, building_key)

        # 如果找不到，尝试添加building_前缀
        if raw_name == building_key:
            building_prefixed_key = f"building_{building_key}"
            raw_name = self.localization_dict.get(building_prefixed_key, building_key)

        building_data['name'] = self.process_building_name(raw_name)

        # 提取icon - 优化正则表达式，避免匹配到ruined_icon
        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', building_content)
        if not icon_match:
            # 如果没有找到带双引号的icon，尝试不带双引号的，但要排除ruined_icon
            icon_match = re.search(r'(?<!ruined_)icon\s*=\s*(\w+)', building_content)

        if icon_match:
            building_data['icon'] = icon_match.group(1)
        else:
            building_data['icon'] = building_key

        # 获取图片路径 - 使用新的图片提取方法
        building_data['image_path'] = self.get_image_path(building_key, building_content)

        # 提取owner_type - 优先于category
        owner_type_match = re.search(r'owner_type\s*=\s*(\w+)', building_content)
        if owner_type_match:
            owner_type = owner_type_match.group(1)
            if owner_type == 'corporate':
                building_data['category'] = "企业分部建筑"
            elif owner_type == 'subject_holding':
                building_data['category'] = "宗主分部建筑"
            else:
                building_data['category'] = f"未知类型: {owner_type}"
        else:
            # 如果没有owner_type，则提取category
            category_match = re.search(r'category\s*=\s*(\w+)', building_content)
            if category_match:
                category_key = category_match.group(1)
                building_data['category'] = self.get_category_localization(category_key)
            else:
                building_data['category'] = "未知类别"

        # 提取building_sets
        building_sets_match = re.search(r'building_sets\s*=\s*\{', building_content)
        if building_sets_match:
            brace_start = building_sets_match.end() - 1
            building_sets_block, _ = self.extract_balanced_braces(building_content, brace_start)

            if building_sets_block:
                # 提取building_sets中的所有值
                sets_matches = re.findall(r'\b(\w+)\b', building_sets_block)
                building_data['building_sets'] = [set_name for set_name in sets_matches
                                                  if set_name not in ['building_sets', '=', '{', '}']]

        self.buildings.append(building_data)

    def parse_all_buildings(self):
        """解析所有建筑文件"""
        print("正在解析建筑文件...")

        if not os.path.exists(BUILDINGS_DIR):
            print(f"建筑目录不存在: {BUILDINGS_DIR}")
            return

        for file_name in os.listdir(BUILDINGS_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(BUILDINGS_DIR, file_name)
                self.parse_building_file(file_path)

        print(f"已解析 {len(self.buildings)} 个建筑")

    def generate_excel(self):
        """生成Excel文件"""
        print("正在生成Excel文件...")

        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}建筑信息.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('建筑信息')

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

        # 设置列宽和行高
        worksheet.set_column('A:A', 15)  # 图片列
        worksheet.set_column('B:B', 40)  # 建筑名称
        worksheet.set_column('C:C', 40)  # 建筑KEY
        worksheet.set_column('D:D', 30)  # 建筑类别
        worksheet.set_column('E:E', 30)  # 建筑允许位置
        worksheet.set_column('F:F', 20)  # 来源文件
        worksheet.set_default_row(80)  # 设置默认行高为80像素

        # 写入表头
        headers = ['图片', '建筑名称', '建筑KEY', '建筑类别', '建筑允许位置', '来源文件']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        row = 1
        for building in self.buildings:
            # 插入图片
            if building['image_path'] and os.path.exists(building['image_path']):
                try:
                    worksheet.insert_image(f'A{row + 1}', building['image_path'], {
                        'x_offset': 5,
                        'y_offset': 5,
                        'x_scale': 1,
                        'y_scale': 1,
                        'object_position': 1
                    })
                except Exception as e:
                    print(f"插入图片失败: {building['image_path']}, 错误: {e}")

            # 写入其他数据
            worksheet.write(row, 1, building['name'], cell_format)
            worksheet.write(row, 2, building['key'], cell_format)
            worksheet.write(row, 3, building['category'], cell_format)

            # 建筑允许位置用逗号分隔
            building_sets_text = ', '.join(building['building_sets'])
            worksheet.write(row, 4, building_sets_text, cell_format)

            worksheet.write(row, 5, building['source'], cell_format)
            row += 1

        workbook.close()

        print(f"Excel文件已生成: {output_path}")

    def run(self):
        """运行完整的提取流程"""
        print("开始提取建筑信息...")

        self.load_localization_files()
        self.load_scripted_loc()
        self.parse_all_buildings()
        self.generate_excel()

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = BuildingExtractor()
    extractor.run()