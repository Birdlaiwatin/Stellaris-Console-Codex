import os
import re
import pandas as pd
from pathlib import Path
import xlsxwriter
import imageio.v3 as imageio
import numpy as np

# 配置路径
from A配置路径 import *
from DLC翻译字典 import DLC_TRANSLATIONS
from 颜色转换字典 import SPECIAL_SUFFIXES

# 创建输出目录
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)


class StellarisPlanetParser:
    def __init__(self):
        self.localisation_dict = {}
        self.scripted_variables = {}
        self.color_dict = {}
        self.gfx_mappings = {}
        self.sprite_sheets = {}
        self.planet_types = []
        self.random_lists = []
        self.processed_references = set()

        # 调试信息
        self.debug_info = {
            'missing_localization': set(),
            'missing_gfx': set(),
        }

        # 构建大小写不敏感的本地化字典
        self.insensitive_localization_dict = {}

    def load_gfx_mappings(self):
        """加载GFX文件中的图标映射"""
        self.gfx_mappings = {}
        self.sprite_sheets = {}

        gfx_files = [
            PLANETS_GFX_FILE_PATH,
        ]

        for gfx_file in gfx_files:
            if not os.path.exists(gfx_file):
                continue

            try:
                with open(gfx_file, 'r', encoding='utf-8') as f:
                    content = f.read()

                sprite_pattern = r'spriteType\s*=\s*\{([^}]+(?:\{[^}]*\}[^}]*)*)\}'
                sprite_blocks = re.findall(sprite_pattern, content, re.DOTALL)

                for block in sprite_blocks:
                    self.parse_sprite_block(block, gfx_file)

            except Exception as e:
                print(f"解析GFX文件失败 {gfx_file}: {e}")

        print(f"已加载 {len(self.gfx_mappings)} 个图标映射")

    def parse_sprite_block(self, block_content, gfx_file):
        """解析单个spriteType块"""
        name_match = re.search(r'name\s*=\s*"([^"]+)"', block_content)
        if not name_match:
            return

        name = name_match.group(1)

        texture_match = re.search(r'texturefile\s*=\s*"([^"]+)"', block_content)
        if texture_match:
            texturefile = texture_match.group(1).replace('/', '\\')
            no_of_frames_match = re.search(r'noOfFrames\s*=\s*(\d+)', block_content)
            no_of_frames = int(no_of_frames_match.group(1)) if no_of_frames_match else 1

            self.sprite_sheets[name] = {
                'texturefile': texturefile,
                'noOfFrames': no_of_frames,
                'source_file': gfx_file
            }

        sprite_sheet_match = re.search(r'sprite_sheet_sprite_type\s*=\s*"([^"]+)"', block_content)
        if sprite_sheet_match:
            sheet_name = sprite_sheet_match.group(1)
            default_frame_match = re.search(r'default_frame\s*=\s*(\d+)', block_content)
            default_frame = int(default_frame_match.group(1)) if default_frame_match else 1

            if sheet_name in self.sprite_sheets:
                texturefile = self.sprite_sheets[sheet_name]['texturefile']
                self.gfx_mappings[name] = {
                    'texturefile': texturefile,
                    'is_sprite_sheet': True,
                    'sheet_name': sheet_name,
                    'frame': default_frame,
                    'total_frames': self.sprite_sheets[sheet_name]['noOfFrames']
                }
            return

        if texture_match:
            texturefile = texture_match.group(1).replace('/', '\\')
            self.gfx_mappings[name] = {
                'texturefile': texturefile,
                'is_sprite_sheet': False,
                'source_file': gfx_file
            }

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

    def convert_dds_to_png_with_path(self, dds_path, planet_key, icon_name):
        """将DDS文件转换为PNG格式，并返回路径"""
        if not os.path.exists(dds_path):
            return None

        try:
            png_filename = f"{planet_key}_{os.path.basename(icon_name)}.png"
            png_path = os.path.join(PLANETS_IMAGES_DIR, "temp", png_filename)
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            img = imageio.imread(dds_path)
            imageio.imwrite(png_path, img)
            return png_path
        except Exception as e:
            return None

    def handle_sprite_sheet(self, sheet_path, mapping, planet_key):
        """处理精灵表，提取指定帧"""
        if not os.path.exists(sheet_path):
            return None

        try:
            sheet_image = imageio.imread(sheet_path)
            total_frames = mapping.get('total_frames', 1)
            frame_num = mapping.get('frame', 1) - 1

            if frame_num < 0 or frame_num >= total_frames:
                frame_num = 0

            if len(sheet_image.shape) == 3:
                sheet_height, sheet_width, channels = sheet_image.shape
            else:
                sheet_height, sheet_width = sheet_image.shape
                channels = 1

            frame_width = sheet_width // total_frames
            start_x = frame_width * frame_num
            end_x = start_x + frame_width

            if end_x > sheet_width:
                end_x = sheet_width

            if len(sheet_image.shape) == 3:
                frame_image = sheet_image[:, start_x:end_x, :]
            else:
                frame_image = sheet_image[:, start_x:end_x]

            png_filename = f"{planet_key}_frame_{frame_num + 1}.png"
            png_path = os.path.join(PLANETS_IMAGES_DIR, "temp", png_filename)
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            imageio.imwrite(png_path, frame_image)
            return png_path

        except Exception as e:
            return self.convert_dds_to_png_with_path(sheet_path, planet_key, f"{planet_key}_full_sheet")

    def find_icon_file(self, icon_path, planet_key):
        """查找图标文件"""
        if not icon_path:
            return None

        if icon_path.startswith('"') and icon_path.endswith('"'):
            icon_path = icon_path[1:-1]

        if icon_path in self.gfx_mappings:
            mapping = self.gfx_mappings[icon_path]
            texturefile = mapping['texturefile']
            full_path = os.path.join(BASE_GAME_DIR, texturefile)

            if mapping.get('is_sprite_sheet', False):
                return self.handle_sprite_sheet(full_path, mapping, planet_key)
            else:
                return self.convert_dds_to_png_with_path(full_path, planet_key, icon_path)

        if icon_path.startswith("gfx/"):
            direct_path = icon_path.replace('/', '\\')
            full_path = os.path.join(BASE_GAME_DIR, direct_path)
            return self.convert_dds_to_png_with_path(full_path, planet_key, icon_path)

        possible_matches = []
        for gfx_name, mapping in self.gfx_mappings.items():
            if icon_path in gfx_name or gfx_name in icon_path:
                possible_matches.append((gfx_name, mapping))

        if possible_matches:
            best_match = min(possible_matches, key=lambda x: len(x[0]))
            texturefile = best_match[1]['texturefile']
            full_path = os.path.join(BASE_GAME_DIR, texturefile)

            if best_match[1].get('is_sprite_sheet', False):
                return self.handle_sprite_sheet(full_path, best_match[1], planet_key)
            else:
                return self.convert_dds_to_png_with_path(full_path, planet_key, best_match[0])

        return None

    def get_planet_image_path(self, planet_key, planet_content):
        """获取星球图片路径"""
        # 首先尝试提取 icon
        icon_match = re.search(r'icon\s*=\s*(?:"([^"]+)"|([^\s{]+))', planet_content)
        if icon_match:
            icon_path = icon_match.group(1) if icon_match.group(1) else icon_match.group(2)
            png_path = self.find_icon_file(icon_path, planet_key)
            if png_path:
                return png_path

        # 如果没有找到icon，尝试提取icon_large
        icon_large_match = re.search(r'icon_large\s*=\s*(?:"([^"]+)"|([^\s{]+))', planet_content)
        if icon_large_match:
            icon_large_path = icon_large_match.group(1) if icon_large_match.group(1) else icon_large_match.group(2)
            png_path = self.find_icon_file(icon_large_path, planet_key)
            if png_path:
                return png_path

        return None

    def load_localisation(self):
        """加载本地化文件"""
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
                    self.localisation_dict[current_key] = value_str

            except Exception:
                pass

        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localisation_dict.items()}
        print(f"已加载 {len(self.localisation_dict)} 个本地化条目")

    def load_scripted_variables(self):
        """加载脚本变量文件"""
        if not os.path.exists(SCRIPTED_VARIABLES_DIR):
            return

        for file_name in os.listdir(SCRIPTED_VARIABLES_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(SCRIPTED_VARIABLES_DIR, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                    matches = re.findall(pattern, content)

                    for var_name, var_value in matches:
                        self.scripted_variables[var_name] = float(var_value)

                except Exception:
                    pass

        print(f"已加载 {len(self.scripted_variables)} 个脚本变量")

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

    def get_localisation(self, key, resolve_nested=True):
        """获取本地化文本"""
        if key in self.localisation_dict:
            text = self.localisation_dict[key]
        elif key.lower() in self.localisation_dict:
            text = self.localisation_dict[key.lower()]
        else:
            return key

        if resolve_nested:
            return self.process_localisation_text(text)
        else:
            return text

    def process_localisation_text(self, text):
        """处理本地化文本中的替换符"""
        if not text:
            return text

        # 处理$xxx$替换符
        text, _ = self.process_dollar_replacement(text)

        # 处理£xxx£替换符
        text = self.process_pound_replacement(text)

        # 处理[]替换符
        text = self.process_bracket_replacement(text)

        # 清理格式字符
        text = re.sub(r'§[A-Z](.*?)§!', r'\1', text)
        text = text.replace('\\n', '\n')

        return text.strip()

    def process_dollar_replacement(self, text, max_depth=5, current_depth=0):
        """处理$xxx$替换符"""
        if current_depth >= max_depth:
            return text, True

        dollar_matches = re.findall(r'\$([^$]+)\$', text)
        if not dollar_matches:
            return text, False

        for match in dollar_matches:
            if match in self.localisation_dict:
                replacement = self.localisation_dict[match]
                replacement, exceeded = self.process_dollar_replacement(replacement, max_depth, current_depth + 1)
                text = text.replace(f'${match}$', replacement)

        if current_depth >= max_depth - 1 and re.search(r'\$[^$]+\$', text):
            return text, True

        return text, False

    def process_pound_replacement(self, text):
        """处理£xxx£替换符"""
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
                    return inner_text

        text = re.sub(r'£([^£]+)£', replace_pounds, text)
        return text

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符"""
        if not text or current_depth >= max_depth:
            return text

        def replace_brackets(match):
            inner_text = match.group(1)

            # 处理.GetName
            name_match = re.match(r'(\w+)\.GetName(Plural)?', inner_text)
            if name_match:
                prefix = name_match.group(1)
                is_plural = name_match.group(2) is not None

                job_key = f"job_{prefix}"
                if is_plural:
                    job_key_plural = f"{job_key}_plural"
                    if job_key_plural in self.localisation_dict:
                        return self.localisation_dict[job_key_plural]
                    elif job_key in self.localisation_dict:
                        return self.localisation_dict[job_key]
                else:
                    if job_key in self.localisation_dict:
                        return self.localisation_dict[job_key]

            # 处理其他情况
            if inner_text.startswith("'") and inner_text.endswith("'"):
                concept_key = inner_text[1:-1]
                if concept_key in self.localisation_dict:
                    return self.localisation_dict[concept_key]

            return match.group(0)

        text = re.sub(r'\[([^\]]+)\]', replace_brackets, text)
        return text

    def parse_planet_file(self, file_path):
        """解析星球类型文件"""
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                content = f.read()

            content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

            pos = 0
            while pos < len(content):
                # 查找星球类型定义 (pc_开头)
                planet_match = re.search(r'(pc_\w+)\s*=\s*\{', content[pos:])
                if planet_match:
                    planet_key = planet_match.group(1)
                    brace_start = pos + planet_match.end() - 1
                    planet_block, new_pos = self.extract_balanced_braces(content, brace_start)

                    if planet_block:
                        self.parse_planet_type(planet_key, planet_block, file_path)
                        pos = new_pos
                    else:
                        pos = brace_start + 1
                else:
                    # 查找random_list
                    random_match = re.search(r'random_list\s*=\s*\{', content[pos:])
                    if random_match:
                        brace_start = pos + random_match.end() - 1
                        random_block, new_pos = self.extract_balanced_braces(content, brace_start)

                        if random_block:
                            self.parse_random_list(random_block, file_path)
                            pos = new_pos
                        else:
                            pos = brace_start + 1
                    else:
                        break

        except Exception as e:
            print(f"解析文件失败 {file_path}: {e}")

    def parse_planet_type(self, planet_key, content, file_path):
        """解析单个星球类型"""
        planet_data = {
            'key': planet_key,
            'name': self.get_localisation(planet_key),
            'icon': None,
            'file_source': os.path.basename(file_path)
        }

        # 获取图片路径
        planet_data['icon'] = self.get_planet_image_path(planet_key, content)

        # 如果找不到本地化名称，记录调试信息
        if planet_data['name'] == planet_key:
            self.debug_info['missing_localization'].add(planet_key)

        self.planet_types.append(planet_data)

    def parse_random_list(self, content, file_path):
        """解析random_list"""
        random_data = {
            'name': '',
            'planets': [],
            'file_source': os.path.basename(file_path)
        }

        # 提取name
        name_match = re.search(r'name\s*=\s*"([^"]+)"', content)
        if name_match:
            random_data['name'] = name_match.group(1)

        # 提取planets列表
        planets_match = re.search(r'planets\s*=\s*\{', content)
        if planets_match:
            brace_start = planets_match.end() - 1
            planets_block, _ = self.extract_balanced_braces(content, brace_start)

            if planets_block:
                # 提取planets块中的所有星球key
                planet_keys = re.findall(r'"([^"]+)"|(\w+)', planets_block)
                for key1, key2 in planet_keys:
                    planet_key = key1 if key1 else key2
                    if planet_key and planet_key.startswith('pc_'):
                        random_data['planets'].append(planet_key)

        self.random_lists.append(random_data)

    def process_all_planets(self):
        """处理所有星球类型文件"""
        print("开始处理星球类型文件...")

        self.load_scripted_variables()
        self.load_localisation()
        self.load_gfx_mappings()

        if not os.path.exists(PLANETS_DIR):
            print(f"星球类型目录不存在: {PLANETS_DIR}")
            return

        planet_files = list(Path(PLANETS_DIR).glob("*.txt"))
        for file_path in planet_files:
            self.parse_planet_file(file_path)

        print(f"解析完成！找到 {len(self.planet_types)} 个星球类型和 {len(self.random_lists)} 个随机列表")
        self.export_results()

    def export_results(self):
        """导出结果到Excel"""
        if not self.planet_types:
            print("没有找到星球类型数据")
            return

        print("正在生成Excel文件...")
        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}星球类型.xlsx')
        workbook = xlsxwriter.Workbook(output_path)

        # 星球类型工作表
        worksheet = workbook.add_worksheet('星球类型')

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

        # 设置列宽
        worksheet.set_column('A:A', 15)  # 图片列
        worksheet.set_column('B:B', 30)  # 星球名称
        worksheet.set_column('C:C', 30)  # 星球KEY
        worksheet.set_column('D:D', 20)  # 文件来源

        worksheet.set_default_row(60)

        headers = ['图片', '星球名称', '星球KEY', '文件来源']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入星球类型数据
        row = 1
        planets_with_images = 0
        for planet in self.planet_types:
            worksheet.write(row, 0, '', cell_format)
            if planet['icon'] and os.path.exists(planet['icon']):
                try:
                    worksheet.insert_image(row, 0, planet['icon'], {
                        'x_offset': 5,
                        'y_offset': 5,
                        'x_scale': 0.6,
                        'y_scale': 0.6,
                        'object_position': 1
                    })
                    planets_with_images += 1
                except Exception as e:
                    print(f"插入图片失败 {planet['icon']}: {e}")

            worksheet.write(row, 1, planet['name'], cell_format)
            worksheet.write(row, 2, planet['key'], cell_format)
            worksheet.write(row, 3, planet['file_source'], cell_format)

            row += 1

        # 在星球类型数据后添加随机列表信息
        if self.random_lists:
            # 空一行作为分隔
            row += 1

            # 写入随机列表标题
            random_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#E2EFDA',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            random_headers = ['随机列表名称', '包含星球', '文件来源']
            for col, header in enumerate(random_headers):
                worksheet.write(row, col, header, random_header_format)

            row += 1

            # 写入随机列表数据
            for random_list in self.random_lists:
                worksheet.write(row, 0, random_list['name'], cell_format)
                worksheet.write(row, 1, ', '.join(random_list['planets']), cell_format)
                worksheet.write(row, 2, random_list['file_source'], cell_format)
                row += 1

        workbook.close()
        print(f"星球类型已导出到: {output_path}")
        print(f"成功生成 {planets_with_images} 个星球的图片")
        print(f"在表格末尾添加了 {len(self.random_lists)} 个随机列表")

        # 导出调试信息
        self.export_debug_info()

    def export_debug_info(self):
        """导出调试信息"""
        if self.debug_info['missing_localization']:
            output_path = os.path.join(OUTPUT_DIR, f'{VERSION}星球类型缺失本地化.txt')
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("=== 缺失本地化的星球类型 ===\n\n")
                for key in sorted(self.debug_info['missing_localization']):
                    f.write(f"{key}\n")
            print(f"缺失本地化信息已导出到: {output_path}")

        if self.debug_info['missing_gfx']:
            output_path = os.path.join(OUTPUT_DIR, f'{VERSION}星球类型缺失图片.txt')
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("=== 缺失图片的星球类型 ===\n\n")
                for key in sorted(self.debug_info['missing_gfx']):
                    f.write(f"{key}\n")
            print(f"缺失图片信息已导出到: {output_path}")


def main():
    parser = StellarisPlanetParser()
    parser.process_all_planets()
    print("处理完成！")


if __name__ == "__main__":
    main()