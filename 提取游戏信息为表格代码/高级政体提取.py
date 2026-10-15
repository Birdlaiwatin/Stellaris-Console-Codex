# -*- coding: utf-8 -*-
"""
群星（Stellaris）"进阶权威政体"数据提取脚本
================================================
从游戏文件中提取所有 advanced_authority_swap（进阶政体）的：
  · 名称 / 描述
  · 效果加成（modifier 解析为中文，带颜色）
  · 特殊机制（custom_tooltip 文本）
  · 触发条件（trigger 解析为中文）
  · 政体图标（DDS 自动转 PNG 并嵌入 Excel）

输出：导出目录下生成 {VERSION}高级政体.xlsx 与 {VERSION}高级政体信息调试.txt

架构与项目内其他提取器（如 标本提取.py）保持一致：
  类结构 + 外部配置/字典导入 + imageio 转 PNG + xlsxwriter 富文本/图片
"""

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


# 基座权威翻译字典（advanced_authority_swap 所属的顶层 authority）
BASE_AUTHORITY_TRANSLATIONS = {
    "auth_democratic": "民主",
    "auth_oligarchic": "寡头",
    "auth_dictatorial": "独裁",
    "auth_imperial": "帝国",
    "auth_corporate": "企业",
    "auth_hive": "蜂巢",
    "auth_machine": "智械",
}

# 基座权威排序映射
BASE_AUTHORITY_ORDER = {
    "民主": 1,
    "寡头": 2,
    "独裁": 3,
    "帝国": 4,
    "企业": 5,
    "蜂巢": 6,
    "智械": 7,
}

# modifier 块内需要跳过的配置指令（非真正修正）
SKIP_MODIFIER_KEYS = {
    "custom_tooltip", "show_only_custom_tooltip", "hidden_effect",
    "inherit_icon", "inherit_effects", "ruler_council_position",
    "has_heir", "election_type", "name", "description", "tags",
}

# trigger 块内需要过滤掉的有效性检查噪声
TRIGGER_NOISE_PATTERN = re.compile(r'is_scope_valid\s*=\s*\w+')


class AuthorityExtractor:
    """进阶权威政体提取器，架构对齐 标本提取.SpecimenExtractor"""

    def __init__(self):
        self.localisation_dict = {}
        self.authorities = []
        self.debug_info = {
            'missing_images': [],
            'missing_modifiers': set()
        }

        # 确保输出目录存在
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(AUTHORITIES_IMAGES_DIR, exist_ok=True)

    # ===================== 通用工具（与标本提取一致） =====================

    def convert_dds_to_png(self, dds_path, png_path):
        """将 DDS 文件转换为 PNG 格式。"""
        try:
            if not os.path.exists(dds_path):
                return False
            os.makedirs(os.path.dirname(png_path), exist_ok=True)
            img = imageio.imread(dds_path)
            imageio.imwrite(png_path, img)
            return True
        except Exception:
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
                    modifier_key.endswith('_uninhabitable')):
                return True
            return False

        use_percentage = should_format_as_percentage(key)

        # 特殊处理 _add 后缀（除非强制百分比）
        if key.endswith('_add') and key not in FORCE_PERCENTAGE_KEYS:
            use_percentage = False

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

    def clean_localisation_value(self, value):
        """清理本地化值，移除注释和多余引号"""
        if '#' in value:
            value = value.split('#')[0].strip()
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]
        return value.strip()

    def preprocess_localisation(self, text):
        """预处理本地化文本 - 清除格式标记、换行符和所有空格"""
        if not text:
            return text
        text = re.sub(r'\$t+\$', '', text)
        text = re.sub(r'\\n', '', text)
        text = re.sub(r'\s+', '', text)
        return text

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
                            parts = line.split(':', 1)
                            if len(parts) == 2:
                                key = parts[0].strip()
                                value = parts[1].strip()
                                value = self.clean_localisation_value(value)
                                value = self.preprocess_localisation(value)
                                localisation_dict[key] = value
            except Exception as e:
                print(f"解析本地化文件时出错 {file_path}: {e}")

        self.localisation_dict = localisation_dict
        print(f"已加载 {len(localisation_dict)} 个本地化条目")

    def resolve_localisation(self, value, depth=0, max_depth=10):
        """递归解析本地化值中的变量引用"""
        if depth > max_depth:
            return value
        value = self.preprocess_localisation(value)
        pattern = r'\$([^$]+)\$'
        matches = re.findall(pattern, value)
        if not matches:
            return value
        for match in matches:
            resolved_value = self.resolve_localisation(
                self.localisation_dict.get(match, match), depth + 1, max_depth)
            value = value.replace(f'${match}$', resolved_value)
        return value

    def find_modifier_translation(self, modifier_key, max_depth=8):
        """查找修正字段的本地化翻译，并递归解析占位符"""
        if modifier_key in MODIFIER_TRANSLATIONS:
            return MODIFIER_TRANSLATIONS[modifier_key], None

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

        return self.resolve_modifier_placeholders(translation, max_depth), None

    def deduplicate_repeated_text(self, text):
        """去除重复的文本，如"灵能使灵能使" -> "灵能使" """
        if not text or len(text) < 4:
            return text

        replacements = {'军用舰船舰队中': '舰队中'}
        result = text
        for old, new in replacements.items():
            result = result.replace(old, new)

        pattern = r'([\u4e00-\u9fa5]{2,}?)\1+'

        def remove_duplicate(match):
            return match.group(1)

        previous_result = result
        result = re.sub(pattern, remove_duplicate, result)
        if result != previous_result:
            return self.deduplicate_repeated_text(result)
        return result

    def resolve_modifier_placeholders(self, text, max_depth=0):
        """递归解析修正文本中的占位符"""
        if max_depth <= 0:
            return text
        text = self.preprocess_localisation(text)

        # 解析 £...£ 格式的占位符
        for match in re.findall(r'£([^£]+)£', text):
            concept_key = f"concept_{match}"
            concept_value = self.localisation_dict.get(concept_key, match)
            resolved_concept = self.resolve_modifier_placeholders(concept_value, max_depth - 1)
            text = text.replace(f'£{match}£', resolved_concept)

        # 解析 $...$ 格式的占位符
        for match in re.findall(r'\$([^$]+)\$', text):
            resolved_value = self.localisation_dict.get(match, match)
            resolved_value = self.resolve_modifier_placeholders(resolved_value, max_depth - 1)
            text = text.replace(f'${match}$', resolved_value)

        text = self.deduplicate_repeated_text(text)
        return text

    def parse_condition(self, condition_content):
        """解析条件并返回翻译后的文本"""
        condition_matches = re.findall(r'(\w+)\s*=\s*(\w+)', condition_content)
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
            if escape_next:
                escape_next = False
                continue
            if char == '\\':
                escape_next = True
                continue
            if char == '"' and not escape_next:
                in_quotes = not in_quotes
                continue
            if not in_quotes:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        return content[start_pos:i + 1]
        return None

    def process_text_with_color_markers(self, text, workbook, default_color='black'):
        """处理文本中的颜色标记符，返回富文本格式"""
        if not text:
            return [], workbook.add_format({'font_color': default_color, 'text_wrap': True})

        red_format = workbook.add_format({'font_color': 'red', 'text_wrap': True})
        green_format = workbook.add_format({'font_color': 'green', 'text_wrap': True})
        black_format = workbook.add_format({'font_color': 'black', 'text_wrap': True})

        if '§R' not in text and '§G' not in text:
            format_obj = workbook.add_format({'font_color': default_color, 'text_wrap': True})
            return [(format_obj, text)], format_obj

        rich_text_parts = []
        current_format = black_format
        current_text = ""

        i = 0
        while i < len(text):
            if text[i:i + 2] == '§R' and text.find('§!', i + 2) != -1:
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    rich_text_parts.append((red_format, text[i + 2:end_pos]))
                    i = end_pos + 2
                else:
                    i += 2
                current_format = black_format
            elif text[i:i + 2] == '§G' and text.find('§!', i + 2) != -1:
                if current_text:
                    rich_text_parts.append((current_format, current_text))
                    current_text = ""
                end_pos = text.find('§!', i + 2)
                if end_pos != -1:
                    rich_text_parts.append((green_format, text[i + 2:end_pos]))
                    i = end_pos + 2
                else:
                    i += 2
                current_format = black_format
            else:
                current_text += text[i]
                i += 1

        if current_text:
            rich_text_parts.append((current_format, current_text))

        return rich_text_parts, black_format

    # ===================== 政体专用逻辑 =====================

    def get_authority_image_path(self, auth_key):
        """获取政体图标路径：直接在 AUTHORITIES_GFX_DIR 找 {key}.dds 转 PNG"""
        if not auth_key:
            return None

        dds_path = os.path.join(AUTHORITIES_GFX_DIR, auth_key + ".dds")
        png_filename = auth_key + ".png"
        png_path = os.path.join(AUTHORITIES_IMAGES_DIR, png_filename)

        if os.path.exists(png_path):
            return png_path

        if os.path.exists(dds_path):
            if self.convert_dds_to_png(dds_path, png_path):
                return png_path
            self.debug_info['missing_images'].append(f"{auth_key} -> 转换失败: {png_filename}")
            return None

        self.debug_info['missing_images'].append(f"{auth_key} -> DDS文件不存在: {dds_path}")
        return None

    def parse_authority_file(self):
        """解析 00_authorities.txt 中所有 advanced_authority_swap 块"""
        print("正在解析政体文件...")

        if not os.path.exists(AUTHORITIES_FILE):
            print(f"错误：找不到政体文件 {AUTHORITIES_FILE}")
            self.authorities = []
            return

        with open(AUTHORITIES_FILE, 'r', encoding='utf-8-sig') as f:
            content = f.read()

        # 移除单行注释
        content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

        authorities = []

        # 找到每个顶层 authority 块（auth_democratic / auth_oligarchic ...）
        for auth_m in re.finditer(r'^(auth_\w+)\s*=\s*\{', content, re.MULTILINE):
            base_authority = auth_m.group(1)
            auth_block = self.extract_complete_block(content, auth_m.start())
            if not auth_block:
                continue

            # 在 authority 块内查找所有 advanced_authority_swap
            for sw_m in re.finditer(r'advanced_authority_swap\s*=\s*\{', auth_block):
                swap_block = self.extract_complete_block(auth_block, sw_m.start())
                if not swap_block:
                    continue

                name_m = re.search(r'name\s*=\s*"([^"]+)"', swap_block)
                if not name_m:
                    continue
                key = name_m.group(1)

                desc_m = re.search(r'description\s*=\s*"([^"]+)"', swap_block)
                desc_key = desc_m.group(1) if desc_m else ""

                # 提取 modifier 块
                mod_start = re.search(r'modifier\s*=\s*\{', swap_block)
                mod_block = ""
                if mod_start:
                    mod_block = self.extract_complete_block(swap_block, mod_start.start()) or ""

                # 提取 trigger 块
                trig_start = re.search(r'trigger\s*=\s*\{', swap_block)
                trig_block = ""
                if trig_start:
                    trig_block = self.extract_complete_block(swap_block, trig_start.start()) or ""

                # ---- 解析效果（modifier 块内的数值修正）----
                effects = []
                missing_modifiers = set()
                # 仅匹配 数值型 key=value，自然跳过 custom_tooltip 等非数值指令
                for mod_key, mod_val in re.findall(r'(\w+)\s*=\s*(-?[\d.]+)', mod_block):
                    if mod_key in SKIP_MODIFIER_KEYS:
                        continue
                    translation, missing_key = self.find_modifier_translation(mod_key)
                    if translation:
                        value_str = self.format_modifier_value(mod_key, mod_val)
                        value_color = self.get_value_color(mod_key, mod_val)
                        color_code = 'G' if value_color == 'green' else ('R' if value_color == 'red' else '')
                        if color_code:
                            effect_text = f"{translation}§{color_code}{value_str}§!"
                        else:
                            effect_text = f"{translation}{value_str}"
                        effects.append({'text': effect_text, 'color': 'black'})
                    else:
                        if missing_key and missing_key not in MODIFIER_TRANSLATIONS:
                            missing_modifiers.add(missing_key)

                # ---- 特殊机制（custom_tooltip 文本）----
                tt_m = re.search(r'custom_tooltip\s*=\s*"([^"]+)"', swap_block)
                tt_key = tt_m.group(1) if tt_m else (key + "_tt")
                tt = self.resolve_localisation(self.localisation_dict.get(tt_key, ""))

                # ---- 触发条件（trigger 块）----
                # 过滤掉 is_scope_valid 等有效性检查噪声
                trig_content = TRIGGER_NOISE_PATTERN.sub('', trig_block).strip()
                condition = self.parse_condition(trig_content) if trig_content else ""

                # ---- 名称 / 描述 ----
                name = self.resolve_localisation(self.localisation_dict.get(key, key))
                desc = self.resolve_localisation(
                    self.localisation_dict.get(desc_key, "")) if desc_key else ""

                # ---- 图标 ----
                image_path = self.get_authority_image_path(key)

                authorities.append({
                    'key': key,
                    'base_authority': base_authority,
                    'name': name,
                    'desc': desc,
                    'effects': effects,
                    'tt': tt,
                    'condition': condition,
                    'image_path': image_path,
                    'missing_modifiers': list(missing_modifiers),
                    'original_index': len(authorities)
                })

                # 收集缺失修正到全局调试信息
                for mm in missing_modifiers:
                    self.debug_info['missing_modifiers'].add(mm)

        self.authorities = authorities
        print(f"已找到 {len(authorities)} 个进阶政体")

    # ===================== 输出 =====================

    def generate_output(self):
        """生成输出结果 - 按基座权威排序，组装 Excel 行"""
        excel_data = []

        # 为每个政体添加翻译后的基座权威
        for authority in self.authorities:
            authority['base_authority_translation'] = BASE_AUTHORITY_TRANSLATIONS.get(
                authority['base_authority'], authority['base_authority'])

        # 按基座权威 + 原始解析顺序排序
        sorted_authorities = sorted(
            self.authorities,
            key=lambda x: (BASE_AUTHORITY_ORDER.get(x['base_authority_translation'], 99),
                           x['original_index'])
        )

        for authority in sorted_authorities:
            name = authority['name'] if authority['name'] else authority['key']
            desc = authority['desc'] if authority['desc'] else "无描述"
            tt = authority['tt'] if authority['tt'] else "无"
            condition = authority['condition'] if authority['condition'] else "无特殊条件"

            excel_data.append({
                "图片": authority['image_path'] or "",
                "政体名称": name,
                "政体KEY": authority['key'],
                "基座权威": authority['base_authority_translation'],
                "描述": desc,
                "效果加成": authority['effects'],
                "特殊机制": tt,
                "触发条件": condition,
                "图片名称": authority['key']
            })

        return excel_data

    def generate_debug_output(self):
        """生成调试输出 - 只包含缺失的修正和缺失的图片"""
        output_lines = []

        if self.debug_info['missing_modifiers']:
            output_lines.append("以下修正找不到本地化文件：")
            for modifier in sorted(self.debug_info['missing_modifiers']):
                output_lines.append(f"  - {modifier}")
            output_lines.append("")

        if self.debug_info['missing_images']:
            output_lines.append("以下政体找不到对应图片：")
            missing_keys = set()
            for info in self.debug_info['missing_images']:
                m = re.search(r'^([^->]+) ->', info)
                if m:
                    missing_keys.add(m.group(1).strip())
            for key in sorted(missing_keys):
                output_lines.append(f"  - {key}")
            output_lines.append("")

        if not self.debug_info['missing_modifiers'] and not self.debug_info['missing_images']:
            output_lines.append("没有发现调试信息")

        return output_lines

    def save_to_excel(self, excel_data):
        """将数据保存到 Excel 文件，并嵌入图片 - 使用 xlsxwriter"""
        if not excel_data:
            print("没有数据可导出到 Excel")
            return

        excel_file = os.path.join(OUTPUT_DIR, f"{VERSION}高级政体.xlsx")
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet("进阶政体")

        # 定义格式
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        wrap_format = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})

        # 设置列宽：A图片 B名称 CKEY D基座 E描述 F效果 G特殊机制 H触发条件 I图片名称
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 22)
        worksheet.set_column('C:C', 38)
        worksheet.set_column('D:D', 10)
        worksheet.set_column('E:E', 40)
        worksheet.set_column('F:F', 44)
        worksheet.set_column('G:G', 40)
        worksheet.set_column('H:H', 30)
        worksheet.set_column('I:I', 38)

        worksheet.set_default_row(80)

        headers = ["图片", "政体名称", "政体KEY", "基座权威", "描述", "效果加成",
                   "特殊机制", "触发条件", "图片名称"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        for row, data in enumerate(excel_data, 1):
            # 图片
            image_path = data["图片"]
            if image_path and os.path.exists(image_path):
                try:
                    worksheet.insert_image(row, 0, image_path, {
                        'x_scale': 0.8, 'y_scale': 0.8,
                        'x_offset': 5, 'y_offset': 5, 'object_position': 1
                    })
                except Exception as e:
                    print(f"插入图片失败: {image_path}, 错误: {e}")
                    worksheet.write(row, 0, "图片加载失败", cell_format)
            else:
                worksheet.write(row, 0, "无图片", cell_format)

            # 基本数据
            worksheet.write(row, 1, data["政体名称"], cell_format)
            worksheet.write(row, 2, data["政体KEY"], cell_format)
            worksheet.write(row, 3, data["基座权威"], cell_format)
            worksheet.write(row, 4, data["描述"], wrap_format)
            worksheet.write(row, 6, data["特殊机制"], wrap_format)
            worksheet.write(row, 7, data["触发条件"], wrap_format)
            worksheet.write(row, 8, data["图片名称"], cell_format)

            # 效果加成（带颜色富文本）
            effects = data["效果加成"]
            if effects:
                effect_text_parts = []
                for i, effect in enumerate(effects):
                    if i > 0:
                        effect_text_parts.append("\n")
                    rich_parts, _ = self.process_text_with_color_markers(effect['text'], workbook)
                    for format_obj, text in rich_parts:
                        effect_text_parts.append((format_obj, text))

                if len(effect_text_parts) == 1 and not isinstance(effect_text_parts[0], tuple):
                    worksheet.write(row, 5, effect_text_parts[0], wrap_format)
                else:
                    rich_text_args = []
                    for part in effect_text_parts:
                        if isinstance(part, tuple):
                            rich_text_args.append(part[0])
                            rich_text_args.append(part[1])
                        else:
                            rich_text_args.append(wrap_format)
                            rich_text_args.append(part)
                    worksheet.write_rich_string(row, 5, *rich_text_args)
            else:
                worksheet.write(row, 5, "无效果", wrap_format)

        workbook.close()
        print(f"Excel 文件已保存到: {excel_file}")

    def save_debug_info(self):
        """保存调试信息到 TXT 文件"""
        print("正在保存调试信息...")
        debug_lines = self.generate_debug_output()
        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}高级政体信息调试.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 进阶政体信息调试信息 ===\n\n")
            f.write('\n'.join(debug_lines))
        print(f"调试信息已保存: {output_path}")

    # ===================== 入口 =====================

    def run(self):
        """运行完整的提取流程"""
        print("开始提取进阶政体信息...")

        self.parse_localisation_files()
        self.parse_authority_file()

        print("正在生成输出...")
        excel_data = self.generate_output()

        self.save_debug_info()
        self.save_to_excel(excel_data)

        print("提取完成！")


# 运行提取器
if __name__ == "__main__":
    extractor = AuthorityExtractor()
    extractor.run()
