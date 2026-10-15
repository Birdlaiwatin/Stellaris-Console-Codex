import os
import re
import pandas as pd
from pathlib import Path
import math
import xlsxwriter
import imageio.v3 as imageio
import numpy as np

# 配置路径 - 使用从A配置路径导入
from A配置路径 import *
from DLC翻译字典 import DLC_TRANSLATIONS
from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
from 颜色转换字典 import SPECIAL_SUFFIXES
from 修正翻译字典 import MODIFIER_TRANSLATIONS
from 条件翻译字典 import CONDITION_TRANSLATIONS

# 固定£替换符翻译字典
SPECIAL_POUND_TRANSLATIONS = {
    'society': '社会学研究',
    'physics': '物理学研究',
    'engineering': '工程学研究',
    'trade_value': '贸易'
}

# 创建输出目录
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)


class StellarisTraitParser:
    def __init__(self):
        self.localisation_dict = {}
        self.modifier_localisation_dict = {}
        self.scripted_variables = {}
        self.scripted_loc_dict = {}
        self.color_dict = {}
        self.gfx_mappings = {}
        self.sprite_sheets = {}
        self.leader_traits = []
        self.processed_references = set()
        self.handled_conditions = set()
        self.untranslated_pound_placeholders = set()
        self.traits_without_effects = []

        # 调试信息
        self.debug_info = {
            'missing_tooltip': set(),
            'missing_scripted_variable': set(),
            'missing_bracket_replacement': set(),
            'missing_pound_replacement': set(),
        }

        # 构建大小写不敏感的本地化字典
        self.insensitive_localization_dict = {}

        # 条件翻译字典
        self.condition_translations = CONDITION_TRANSLATIONS

    def load_gfx_mappings(self):
        """加载GFX文件中的图标映射"""
        self.gfx_mappings = {}
        self.sprite_sheets = {}

        gfx_files = [
            LEADER_TRAITS_GFX_FILE_PATH,
            LEADER_TRAITS_2_GFX_FILE_PATH
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

    def convert_dds_to_png_with_path(self, dds_path, trait_key, icon_name):
        """将DDS文件转换为PNG格式，并返回路径"""
        if not os.path.exists(dds_path):
            return None

        try:
            png_filename = f"{trait_key}_{os.path.basename(icon_name)}.png"
            png_path = os.path.join(LEADER_TRAITS_IMAGES_DIR, "temp", png_filename)
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            img = imageio.imread(dds_path)
            imageio.imwrite(png_path, img)
            return png_path
        except Exception as e:
            return None

    def handle_sprite_sheet(self, sheet_path, mapping, trait_key):
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

            png_filename = f"{trait_key}_frame_{frame_num + 1}.png"
            png_path = os.path.join(LEADER_TRAITS_IMAGES_DIR, "temp", png_filename)
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            imageio.imwrite(png_path, frame_image)
            return png_path

        except Exception as e:
            return self.convert_dds_to_png_with_path(sheet_path, trait_key, f"{trait_key}_full_sheet")

    def get_trait_image_path(self, trait_key, trait_content):
        """获取特质图片路径"""
        icon_match = re.search(r'icon\s*=\s*{', trait_content)
        if not icon_match:
            return None

        brace_start = icon_match.end() - 1
        icon_block, _ = self.extract_balanced_braces(trait_content, brace_start)

        if not icon_block:
            return None

        processed_icon_block, script_count = self.process_inline_script(icon_block, "trait_icon", trait_key)

        layers = self.extract_icon_layers(processed_icon_block)

        if not layers:
            return None

        final_image = self.merge_icon_layers(layers, trait_key)

        if final_image is not None:
            png_filename = f"{trait_key}.png"
            png_path = os.path.join(LEADER_TRAITS_IMAGES_DIR, png_filename)
            os.makedirs(os.path.dirname(png_path), exist_ok=True)

            try:
                # 使用PIL库保存高质量PNG
                from PIL import Image
                pil_img = Image.fromarray(final_image)
                pil_img.save(png_path, 'PNG', optimize=True, quality=95)
                return png_path
            except Exception:
                try:
                    imageio.imwrite(png_path, final_image)
                    return png_path
                except Exception:
                    return None

        return None

    def extract_icon_layers(self, icon_block):
        """提取icon块中的所有元素，保持原始顺序"""
        layers = []
        pos = 0

        while pos < len(icon_block):
            layer_match = re.search(r'layer\s*=\s*{', icon_block[pos:])
            icon_match = re.search(r'icon\s*=\s*(?:"([^"]+)"|([^\s{]+))', icon_block[pos:])

            if layer_match and icon_match:
                if layer_match.start() < icon_match.start():
                    brace_start = pos + layer_match.end() - 1
                    layer_block, new_pos = self.extract_balanced_braces(icon_block, brace_start)

                    if layer_block:
                        layer_data = self.parse_layer_block(layer_block)
                        if layer_data:
                            layers.append(layer_data)

                    pos = new_pos
                else:
                    icon_start = pos + icon_match.start()
                    icon_end = pos + icon_match.end()

                    icon_path = icon_match.group(1) if icon_match.group(1) else icon_match.group(2)

                    color_match = None
                    color_search_start = max(0, icon_start - 100)
                    color_search_end = min(len(icon_block), icon_end + 100)
                    search_area = icon_block[color_search_start:color_search_end]

                    color_match = re.search(r'color\s*=\s*"([^"]+)"', search_area)
                    if not color_match:
                        color_match = re.search(r'color\s*=\s*([^\s{]+)', search_area)

                    layer_data = {
                        'icon_path': icon_path,
                        'color': self.get_color_value(color_match.group(1)) if color_match else None
                    }

                    layers.append(layer_data)
                    pos = icon_end
            elif layer_match:
                brace_start = pos + layer_match.end() - 1
                layer_block, new_pos = self.extract_balanced_braces(icon_block, brace_start)

                if layer_block:
                    layer_data = self.parse_layer_block(layer_block)
                    if layer_data:
                        layers.append(layer_data)

                pos = new_pos
            elif icon_match:
                icon_start = pos + icon_match.start()
                icon_end = pos + icon_match.end()

                icon_path = icon_match.group(1) if icon_match.group(1) else icon_match.group(2)

                color_match = None
                color_search_start = max(0, icon_start - 100)
                color_search_end = min(len(icon_block), icon_end + 100)
                search_area = icon_block[color_search_start:color_search_end]

                color_match = re.search(r'color\s*=\s*"([^"]+)"', search_area)
                if not color_match:
                    color_match = re.search(r'color\s*=\s*([^\s{]+)', search_area)

                layer_data = {
                    'icon_path': icon_path,
                    'color': self.get_color_value(color_match.group(1)) if color_match else None
                }

                layers.append(layer_data)
                pos = icon_end
            else:
                break

        return layers

    def parse_layer_block(self, layer_block):
        """解析单个layer块"""
        layer_data = {
            'icon_path': None,
            'color': None
        }

        icon_match = re.search(r'icon\s*=\s*"([^"]+)"', layer_block)
        if not icon_match:
            icon_match = re.search(r'icon\s*=\s*([^\s{]+)', layer_block)

        if icon_match:
            icon_path = icon_match.group(1) if icon_match.group(1) else icon_match.group(2)
            layer_data['icon_path'] = icon_path

        color_match = re.search(r'color\s*=\s*"([^"]+)"', layer_block)
        if color_match:
            color_name = color_match.group(1)
            layer_data['color'] = self.get_color_value(color_name)
        else:
            color_match = re.search(r'color\s*=\s*([^\s{]+)', layer_block)
            if color_match:
                color_name = color_match.group(1)
                layer_data['color'] = self.get_color_value(color_name)

        return layer_data if layer_data['icon_path'] else None

    def get_color_value(self, color_name):
        """从颜色字典中获取颜色值"""
        if not color_name:
            return None

        if color_name.startswith('"') and color_name.endswith('"'):
            color_name = color_name[1:-1]

        if color_name in self.color_dict:
            return self.color_dict[color_name]

        for key, value in self.color_dict.items():
            if key.endswith(f".{color_name}") or key == color_name:
                return value

        return None

    def merge_icon_layers(self, layers, trait_key):
        """合并所有图层，按照中心点叠加，保持所有内容完整"""
        if not layers:
            return None

        base_image = None

        for i, layer in enumerate(layers):
            icon_path = self.find_icon_file(layer['icon_path'], trait_key)
            if not icon_path:
                continue

            try:
                layer_image = imageio.imread(icon_path)

                if layer_image.shape[2] == 3:
                    alpha_channel = np.ones((layer_image.shape[0], layer_image.shape[1]), dtype=layer_image.dtype) * 255
                    layer_image = np.dstack((layer_image, alpha_channel))

                if layer['color'] and len(layer['color']) == 4:
                    layer_image = self.apply_color_to_image(layer_image, layer['color'])

                if base_image is None:
                    base_image = layer_image
                else:
                    base_image = self.merge_images_preserve_all(base_image, layer_image)

            except Exception:
                continue

        return base_image

    def merge_images_preserve_all(self, base, overlay):
        """合并两张图片，按照中心点对齐，保留所有内容"""
        try:
            base_h, base_w = base.shape[:2]
            overlay_h, overlay_w = overlay.shape[:2]

            canvas_h = max(base_h, overlay_h)
            canvas_w = max(base_w, overlay_w)

            canvas = np.zeros((canvas_h, canvas_w, 4), dtype=np.uint8)

            base_y = (canvas_h - base_h) // 2
            base_x = (canvas_w - base_w) // 2
            overlay_y = (canvas_h - overlay_h) // 2
            overlay_x = (canvas_w - overlay_w) // 2

            if base.shape[2] == 4:
                base_alpha = base[:, :, 3] / 255.0
                for c in range(3):
                    canvas[base_y:base_y + base_h, base_x:base_x + base_w, c] = (
                            base[:, :, c] * base_alpha +
                            canvas[base_y:base_y + base_h, base_x:base_x + base_w, c] * (1 - base_alpha)
                    ).astype(np.uint8)
                canvas[base_y:base_y + base_h, base_x:base_x + base_w, 3] = np.maximum(
                    canvas[base_y:base_y + base_h, base_x:base_x + base_w, 3],
                    (base_alpha * 255).astype(np.uint8)
                )
            else:
                canvas[base_y:base_y + base_h, base_x:base_x + base_w, :3] = base
                canvas[base_y:base_y + base_h, base_x:base_x + base_w, 3] = 255

            if overlay.shape[2] == 4:
                overlay_alpha = overlay[:, :, 3] / 255.0
                for c in range(3):
                    canvas[overlay_y:overlay_y + overlay_h, overlay_x:overlay_x + overlay_w, c] = (
                            overlay[:, :, c] * overlay_alpha +
                            canvas[overlay_y:overlay_y + overlay_h, overlay_x:overlay_x + overlay_w, c] * (
                                    1 - overlay_alpha)
                    ).astype(np.uint8)
                canvas[overlay_y:overlay_y + overlay_h, overlay_x:overlay_x + overlay_w, 3] = np.maximum(
                    canvas[overlay_y:overlay_y + overlay_h, overlay_x:overlay_x + overlay_w, 3],
                    (overlay_alpha * 255).astype(np.uint8)
                )
            else:
                canvas[overlay_y:overlay_y + overlay_h, overlay_x:overlay_x + overlay_w, :3] = overlay
                canvas[overlay_y:overlay_y + overlay_h, overlay_x:overlay_x + overlay_w, 3] = 255

            return canvas

        except Exception:
            return base

    def apply_color_to_image(self, image, color):
        """将颜色应用到图片"""
        try:
            if len(color) != 4:
                return image

            r, g, b, a = color

            if image.shape[2] == 3:
                alpha_channel = np.ones((image.shape[0], image.shape[1]), dtype=image.dtype) * 255
                image = np.dstack((image, alpha_channel))

            r_norm = r / 255.0
            g_norm = g / 255.0
            b_norm = b / 255.0
            a_norm = a / 255.0

            color_matrix = np.ones_like(image, dtype=np.float32)
            color_matrix[:, :, 0] = r_norm
            color_matrix[:, :, 1] = g_norm
            color_matrix[:, :, 2] = b_norm
            color_matrix[:, :, 3] = a_norm

            result = (image.astype(np.float32) * color_matrix).astype(np.uint8)

            return result
        except Exception:
            return image

    def find_icon_file(self, icon_path, trait_key):
        """查找图标文件"""
        if not icon_path:
            return None

        if icon_path.startswith('"') and icon_path.endswith('"'):
            icon_path = icon_path[1:-1]

        if icon_path == "GFX_trait_disabled":
            return None

        if icon_path in self.gfx_mappings:
            mapping = self.gfx_mappings[icon_path]
            texturefile = mapping['texturefile']
            full_path = os.path.join(BASE_GAME_DIR, texturefile)

            if mapping.get('is_sprite_sheet', False):
                return self.handle_sprite_sheet(full_path, mapping, trait_key)
            else:
                return self.convert_dds_to_png_with_path(full_path, trait_key, icon_path)

        if icon_path.startswith("gfx/"):
            direct_path = icon_path.replace('/', '\\')
            full_path = os.path.join(BASE_GAME_DIR, direct_path)
            return self.convert_dds_to_png_with_path(full_path, trait_key, icon_path)

        possible_matches = []
        for gfx_name, mapping in self.gfx_mappings.items():
            if icon_path in gfx_name or gfx_name in icon_path:
                possible_matches.append((gfx_name, mapping))

        if possible_matches:
            best_match = min(possible_matches, key=lambda x: len(x[0]))
            texturefile = best_match[1]['texturefile']
            full_path = os.path.join(BASE_GAME_DIR, texturefile)

            if best_match[1].get('is_sprite_sheet', False):
                return self.handle_sprite_sheet(full_path, best_match[1], trait_key)
            else:
                return self.convert_dds_to_png_with_path(full_path, trait_key, best_match[0])

        return None

    def process_inline_script(self, content, file_name, deposit_key=None):
        """处理inline_script块"""
        if not content:
            return content, 0

        processed_count = 0
        original_content = content

        simple_pattern = r'inline_script\s*=\s*(?:"([^"]+)"|([^\s{]+))'
        pos = 0

        while pos < len(content):
            match = re.search(simple_pattern, content[pos:])
            if not match:
                break

            script_start = pos + match.start()
            script_end = pos + match.end()

            script_path = match.group(1) if match.group(1) else match.group(2)

            normalized_script_path = script_path.replace('/', os.path.sep)
            script_file_path = os.path.join(INLINE_SCRIPTS_DIR, f"{normalized_script_path}.txt")

            if os.path.exists(script_file_path):
                try:
                    with open(script_file_path, 'r', encoding='utf-8') as f:
                        script_content = f.read()

                    content = content[:script_start] + script_content + content[script_end:]
                    processed_count += 1

                    pos = script_start + len(script_content)
                except Exception:
                    pos = script_end
            else:
                pos = script_end

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
                    normalized_script_path = script_path.replace('/', os.path.sep)
                    script_file_path = os.path.join(INLINE_SCRIPTS_DIR, f"{normalized_script_path}.txt")

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
                            processed_count += 1

                            pos = brace_start
                        except Exception:
                            pos = new_pos
                    else:
                        pos = new_pos
                else:
                    pos = new_pos
            else:
                pos = brace_start + 1

        return content, processed_count

    def parse_color_definitions(self, content, filename):
        """解析颜色定义"""
        content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

        lines = [line.strip() for line in content.split('\n') if line.strip()]
        content = '\n'.join(lines)

        color_pattern = r'(\w+)\s*=\s*\{\s*color\s*=\s*\{\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\}\s*\}'
        matches = re.findall(color_pattern, content)

        for color_name, r, g, b, a in matches:
            r, g, b, a = map(int, (r, g, b, a))
            self.color_dict[color_name] = (r, g, b, a)

        simple_pattern = r'(\w+)\s*\{\s*color\s*=\s*\{\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\}\s*\}'
        simple_matches = re.findall(simple_pattern, content)

        for color_name, r, g, b, a in simple_matches:
            r, g, b, a = map(int, (r, g, b, a))
            self.color_dict[color_name] = (r, g, b, a)

        if not matches and not simple_matches:
            self.parse_color_definitions_line_by_line(content, filename)

    def parse_color_definitions_line_by_line(self, content, filename):
        """逐行解析颜色定义"""
        lines = content.split('\n')
        current_color = None
        in_color_block = False
        in_color_value = False
        color_values = []

        for line_num, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            if '=' in line and '{' in line and not in_color_block:
                parts = line.split('=', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    value = parts[1].strip()

                    if value.startswith('{'):
                        current_color = key
                        in_color_block = True
                        continue

            elif '{' in line and not '=' in line and not in_color_block:
                match = re.match(r'(\w+)\s*\{', line)
                if match:
                    current_color = match.group(1)
                    in_color_block = True
                    continue

            if in_color_block and current_color:
                if 'color' in line and '{' in line:
                    in_color_value = True
                    color_values = []
                    continue

                if in_color_value:
                    numbers = re.findall(r'\d+', line)
                    color_values.extend(numbers)

                    if '}' in line:
                        if len(color_values) >= 4:
                            try:
                                r, g, b, a = map(int, color_values[:4])
                                self.color_dict[current_color] = (r, g, b, a)
                            except ValueError:
                                pass

                        in_color_value = False
                        color_values = []

                if '}' in line and not in_color_value:
                    in_color_block = False
                    current_color = None

    def load_colors(self):
        """加载颜色定义"""
        if not os.path.exists(COLORS_DIR):
            return

        color_files = list(Path(COLORS_DIR).glob("*.txt"))
        if not color_files:
            return

        primary_color_file = os.path.join(COLORS_DIR, "01_trait_colors.txt")
        if os.path.exists(primary_color_file):
            try:
                with open(primary_color_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.parse_color_definitions(content, "01_trait_colors.txt")
            except Exception:
                pass

        for file_path in color_files:
            if file_path.name == "01_trait_colors.txt":
                continue

            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.parse_color_definitions(content, file_path.name)
            except Exception:
                pass

        print(f"已加载 {len(self.color_dict)} 个颜色定义")

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

        if any(key.endswith(suffix) for suffix in SPECIAL_SUFFIXES):
            return 'red' if value_float > 0 else 'green'

        return 'green' if value_float > 0 else 'red'

    def preprocess_localization_text(self, text):
        """预处理本地化文本"""
        if not text:
            return text

        text, _ = self.process_dollar_replacement(text)
        text = re.sub(r'§[A-Z](.*?)§!', r'\1', text)

        return text

    def remove_duplicate_text(self, text):
        """清除连续重复文本，只处理文字，不处理数字"""
        if not text:
            return text

        text = re.sub(r'\s+', '', text)
        pattern = r'([\u4e00-\u9fa5a-zA-Z]{2,10})\1+'

        changed = True
        while changed:
            changed = False
            matches = re.findall(pattern, text)
            if matches:
                for match in matches:
                    text = text.replace(match + match, match)
                    changed = True

        return text

    def clean_whitespace_text(self, text):
        """清理空白文本"""
        if not text:
            return text

        text = re.sub(r'\s+', ' ', text)
        text = text.strip()

        return text

    def process_special_at_replacement(self, text):
        """处理特殊@替换符结构：@[ xxx * 数值]"""
        if not text:
            return text

        pattern = r'@\[\s*(\w+)\s*\*\s*([\d.-]+)\s*\]'

        def replace_special_at(match):
            var_name = match.group(1)
            multiplier = float(match.group(2))

            if var_name in self.scripted_variables:
                result = self.scripted_variables[var_name] * multiplier

                if abs(result - round(result)) < 1e-10:
                    result = round(result)

                if result == int(result):
                    return str(int(result))
                else:
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

        pattern = r'\$@(\w+)\|([^$]+)%?\$'

        def replace_special_dollar_at(match):
            var_name = match.group(1)
            default_value = match.group(2)
            has_percent = match.group(0).endswith('%$')

            if var_name in self.scripted_variables:
                result = self.scripted_variables[var_name]
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
                if has_percent:
                    return f"{default_value}%"
                else:
                    return default_value

        text = re.sub(pattern, replace_special_dollar_at, text)

        return text

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

            except Exception:
                pass

        self.insensitive_localization_dict = {k.lower(): v for k, v in self.localisation_dict.items()}

        for key, value in self.localisation_dict.items():
            if (key.startswith('mod_') or key.startswith('MOD_') or
                    key.startswith('MOD_COUNTRY_') or key.startswith('MOD_FEDERATION_') or
                    key.startswith('MOD_LEADER_') or key.startswith('concept_')):
                self.modifier_localisation_dict[key] = value
                self.modifier_localisation_dict[key.lower()] = value

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

        if os.path.exists(TRAITS_DIR):
            for file_name in os.listdir(TRAITS_DIR):
                if file_name.endswith('.txt'):
                    file_path = os.path.join(TRAITS_DIR, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        lines = content.split('\n')
                        brace_level = 0

                        for line in lines:
                            line = line.strip()
                            if not line or line.startswith('#'):
                                continue

                            for char in line:
                                if char == '{':
                                    brace_level += 1
                                elif char == '}':
                                    brace_level -= 1

                            if brace_level == 0:
                                pattern = r'@(\w+)\s*=\s*([\d.-]+)'
                                match = re.match(pattern, line)
                                if match:
                                    var_name, var_value = match.groups()
                                    self.scripted_variables[var_name] = float(var_value)

                    except Exception:
                        pass

        print(f"已加载 {len(self.scripted_variables)} 个脚本变量")

    def load_scripted_loc(self):
        """加载脚本本地化"""
        if not os.path.exists(SCRIPTED_LOC_DIR):
            return

        for file_name in os.listdir(SCRIPTED_LOC_DIR):
            if file_name.endswith('.txt'):
                file_path = os.path.join(SCRIPTED_LOC_DIR, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    self.extract_defined_text_blocks(content)

                except Exception:
                    pass

        print(f"已加载 {len(self.scripted_loc_dict)} 个脚本本地化条目")

    def extract_defined_text_blocks(self, content):
        """提取defined_text块，处理嵌套花括号"""
        lines = content.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith('defined_text'):
                block_lines = []
                brace_count = 0
                in_block = False

                while i < len(lines):
                    current_line = lines[i]
                    block_lines.append(current_line)

                    for char in current_line:
                        if char == '{':
                            brace_count += 1
                            in_block = True
                        elif char == '}':
                            brace_count -= 1

                    i += 1

                    if in_block and brace_count == 0:
                        break

                block_content = '\n'.join(block_lines)
                self.parse_defined_text_block(block_content)
            else:
                i += 1

    def parse_defined_text_block(self, block_content):
        """解析单个defined_text块"""
        name_match = re.search(r'name\s*=\s*(\w+)', block_content)
        if not name_match:
            return

        name = name_match.group(1)

        default_match = re.search(r'default\s*=\s*(\w+)', block_content)
        if default_match:
            self.scripted_loc_dict[name] = default_match.group(1)
        else:
            loc_key_match = re.search(r'localization_key\s*=\s*(\w+)', block_content)
            if loc_key_match:
                self.scripted_loc_dict[name] = loc_key_match.group(1)

    def extract_modifier_value(self, value_str):
        """提取修正值，处理@变量"""
        if isinstance(value_str, (int, float)):
            return value_str

        if isinstance(value_str, str) and '#' in value_str:
            value_str = value_str.split('#')[0].strip()

        value_str = self.process_special_at_replacement(value_str)
        value_str = self.process_special_dollar_at_replacement(value_str)

        if isinstance(value_str, str) and value_str.startswith('@'):
            var_name = value_str[1:]
            if var_name in self.scripted_variables:
                return self.scripted_variables[var_name]
            else:
                self.debug_info['missing_scripted_variable'].add(var_name)
                return value_str

        if isinstance(value_str, str) and '$' in value_str:
            value_str, _ = self.process_dollar_replacement(value_str)

        return value_str

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

    def process_bracket_replacement(self, text, max_depth=3, current_depth=0):
        """处理[]替换符"""
        if not text or current_depth >= max_depth:
            return text

        text = re.sub(r'\[[^\]]*\.GetIcon\]', '', text)

        def replace_brackets(match):
            inner_text = match.group(1)
            result = inner_text

            name_match = re.match(r'(\w+)\.GetName(Plural)?', inner_text)
            if name_match:
                prefix = name_match.group(1)
                is_plural = name_match.group(2) is not None

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

            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' not in inner_text:
                concept_key = inner_text[1:-1]
                if concept_key in self.localisation_dict:
                    result = self.localisation_dict[concept_key]
                else:
                    self.debug_info['missing_bracket_replacement'].add(
                        f"未找到概念本地化: {concept_key} (来自 [{inner_text}])"
                    )

            elif inner_text.startswith("'") and inner_text.endswith("'") and ':' in inner_text:
                parts = inner_text[1:-1].split(':', 1)
                if len(parts) == 2:
                    building_key = parts[1]
                    if building_key in self.localisation_dict:
                        result = self.localisation_dict[building_key]
                    else:
                        self.debug_info['missing_bracket_replacement'].add(
                            f"未找到建筑本地化: {building_key} (来自 [{inner_text}])"
                        )

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

            if concept_key in self.localisation_dict:
                return self.localisation_dict[concept_key]
            else:
                if inner_text in self.localisation_dict:
                    return self.localisation_dict[inner_text]
                else:
                    self.debug_info['missing_pound_replacement'].add(
                        f"概念本地化 '{concept_key}' 和 '{inner_text}' 均未找到")
                    return ""

        text = re.sub(r'£([^£]+)£', replace_pounds, text)

        text, _ = self.process_dollar_replacement(text)
        text = self.process_bracket_replacement(text)

        return text

    def process_special_pound_replacement(self, text):
        """处理特殊的£xxx£替换符"""
        if not text:
            return text

        text = text.replace('£blocker£', '')

        for pound_key, translation in SPECIAL_POUND_TRANSLATIONS.items():
            text = text.replace(f'£{pound_key}£', translation)

        return text

    def process_localisation_text(self, text, modifier_key=""):
        """通用文本处理方法"""
        if not text:
            return text

        text = self.preprocess_localization_text(text)
        text = self.clean_whitespace_text(text)
        text = self.process_special_dollar_at_replacement(text)
        text = self.process_bracket_replacement(text)
        text = self.process_special_pound_replacement(text)
        text = self.process_pound_replacement(text)
        text = self.remove_duplicate_text(text)
        text = text.replace('\\n', '\n')
        text = self.clean_whitespace_text(text)

        return text

    def get_localisation(self, key, resolve_nested=True):
        """获取本地化文本"""
        if key in self.localisation_dict:
            text = self.localisation_dict[key]
        elif key.lower() in self.localisation_dict:
            text = self.localisation_dict[key.lower()]
        else:
            if key.endswith(('_2', '_3')):
                base_key = key.rsplit('_', 1)[0]
                roman_numeral = 'II' if key.endswith('_2') else 'III'

                base_name = self.get_localisation(base_key, resolve_nested=False)
                if base_name != base_key:
                    if '[' in base_name and ']' in base_name:
                        base_name = self.process_bracket_replacement(base_name)
                    return f"{base_name}{roman_numeral}"

            return key

        if resolve_nested:
            return self.process_localisation_text(text)
        else:
            return text

    def get_modifier_localisation(self, modifier_key):
        """获取修正字段的中文描述"""
        builtin_modifiers = MODIFIER_TRANSLATIONS

        if modifier_key in builtin_modifiers:
            return builtin_modifiers[modifier_key]

        search_patterns = [
            [f"mod_{modifier_key}", f"mod_{modifier_key}".lower(), f"MOD_{modifier_key}".lower()],
            [f"MOD_{modifier_key}", f"MOD_{modifier_key}".lower()],
            [modifier_key],
            [f"MOD_COUNTRY_{modifier_key}", f"MOD_COUNTRY_{modifier_key}".lower()],
            [f"MOD_FEDERATION_{modifier_key}", f"MOD_FEDERATION_{modifier_key}".lower()],
            [f"MOD_LEADER_{modifier_key}", f"MOD_LEADER_{modifier_key}".lower()],
            [f"edict_{modifier_key}", f"edict_{modifier_key}".lower()],
        ]

        for pattern_group in search_patterns:
            for pattern in pattern_group:
                if pattern in self.modifier_localisation_dict:
                    return self.process_localisation_text(self.modifier_localisation_dict[pattern])

                if pattern in self.localisation_dict:
                    return self.process_localisation_text(self.localisation_dict[pattern])

        if modifier_key.endswith('_mult'):
            base_key = modifier_key[:-5]
            mod_patterns = [
                f"MOD_{base_key}",
                f"MOD_{base_key}".lower(),
                f"mod_{base_key}",
                f"mod_{base_key}".lower()
            ]

            for pattern in mod_patterns:
                if pattern in self.modifier_localisation_dict:
                    return self.process_localisation_text(self.modifier_localisation_dict[pattern])

                if pattern in self.localisation_dict:
                    return self.process_localisation_text(self.localisation_dict[pattern])

        if modifier_key.endswith('_mult'):
            base_key = modifier_key[:-5]
            if base_key in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key])
            if base_key in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key])
            if base_key.lower() in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key.lower()])
            if base_key.lower() in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key.lower()])

        elif modifier_key.endswith('_add'):
            base_key = modifier_key[:-4]
            if base_key in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key])
            if base_key in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key])
            if base_key.lower() in self.modifier_localisation_dict:
                return self.process_localisation_text(self.modifier_localisation_dict[base_key.lower()])
            if base_key.lower() in self.localisation_dict:
                return self.process_localisation_text(self.localisation_dict[base_key.lower()])

        return modifier_key

    def parse_value(self, value_str):
        """解析数值，处理@primary等占位符"""
        if isinstance(value_str, (int, float)):
            return value_str

        value_str = str(value_str).strip()

        if value_str.startswith('@'):
            if value_str in self.scripted_variables:
                return self.scripted_variables[value_str]
            else:
                if value_str == '@primary':
                    return 1.0
                elif value_str == '@secondary':
                    return 0.5
                elif value_str == '@class_negative_trait_weight':
                    return 10
                else:
                    return 1.0

        try:
            if '.' in value_str:
                return float(value_str)
            else:
                return int(value_str)
        except ValueError:
            return value_str

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

    def remove_excluded_blocks(self, content):
        """移除需要排除的块"""
        excluded_blocks = ['selectable_weight', 'ai_weight']

        for block_type in excluded_blocks:
            pattern = rf'{block_type}\s*=\s*{{'
            pos = 0
            while pos < len(content):
                match = re.search(pattern, content[pos:])
                if not match:
                    break

                brace_start = pos + match.end() - 1
                block_content, new_pos = self.extract_balanced_braces(content, brace_start)

                if block_content:
                    content = content[:brace_start] + content[new_pos:]
                    pos = brace_start
                else:
                    pos = brace_start + 1

        return content

    def parse_trait_file(self, file_path):
        """解析特质文件，通过是否含有leader_class判定是否为领袖特质"""
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                content = f.read()

            content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

            content, script_count = self.process_inline_script(content, os.path.basename(file_path))

            pos = 0
            while pos < len(content):
                match = re.search(r'(\w+)\s*=\s*\{', content[pos:])
                if not match:
                    break

                trait_key = match.group(1)
                brace_start = pos + match.end() - 1

                trait_block, new_pos = self.extract_balanced_braces(content, brace_start)
                if trait_block:
                    if 'leader_class' in trait_block or 'leader_trait_type' in trait_block:
                        self.parse_leader_trait(trait_key, trait_block)

                    pos = new_pos
                else:
                    pos = brace_start + 1

        except Exception:
            pass

    def parse_leader_trait(self, trait_key, content):
        """解析领袖特质"""
        trait_data = {
            'key': trait_key,
            'name': self.get_localisation(trait_key),
            'other_names': [],
            'effects': [],
            'extra_effects': [],
            'opposites': [],
            'leader_classes': [],
            'trait_type': '普通',
            'council_trait': '是',
            'randomized': True,
            'allowed_classes': [],
            'add_conditions': [],
            'modifier_conditions': [],
            'raw_add_conditions': [],
            'raw_modifier_conditions': [],
            'icon': None
        }

        trait_data['icon'] = self.get_trait_image_path(trait_key, content)

        base_name = trait_data['name']

        base_trait_key = trait_key
        roman_numeral = ""

        if trait_key.endswith('_2'):
            base_trait_key = trait_key[:-2]
            roman_numeral = "II"
        elif trait_key.endswith('_3'):
            base_trait_key = trait_key[:-2]
            roman_numeral = "III"

        hive_key = f"{base_trait_key}_hive"
        machine_key = f"{base_trait_key}_machine"

        if hive_key in self.localisation_dict:
            hive_name = self.get_localisation(hive_key)
            if roman_numeral:
                full_hive_name = f"{hive_name}{roman_numeral}"
            else:
                full_hive_name = hive_name

            if full_hive_name != base_name:
                trait_data['other_names'].append(f"蜂巢：{full_hive_name}")

        if machine_key in self.localisation_dict:
            machine_name = self.get_localisation(machine_key)
            if roman_numeral:
                full_machine_name = f"{machine_name}{roman_numeral}"
            else:
                full_machine_name = machine_name

            if full_machine_name != base_name:
                trait_data['other_names'].append(f"机械：{full_machine_name}")

        trait_type = '普通'
        is_negative = 'leader_trait_type = negative' in content
        is_subclass = 'leader_trait_type = subclass' in content

        rarity_match = re.search(r'RARITY\s*=\s*(\w+)', content)
        if rarity_match:
            rarity = rarity_match.group(1)
            if rarity == 'veteran':
                trait_type = '老练'
            elif rarity == 'paragon':
                trait_type = '命定'
            elif rarity == 'common':
                trait_type = '普通'

        if is_negative:
            trait_type = '负面'
        elif is_subclass:
            trait_type = '职业'

        trait_data['trait_type'] = trait_type

        council_match = re.search(r'COUNCIL\s*=\s*(\w+)', content)
        if council_match:
            council_value = council_match.group(1)
            if council_value == 'no':
                trait_data['council_trait'] = '否'

        tooltip_pattern = r'custom_tooltip(?:_with_modifiers)?\s*=\s*"?([^"\s]+)"?'
        tooltip_matches = re.findall(tooltip_pattern, content)
        for tooltip_key in tooltip_matches:
            tooltip_text = self.get_localisation(tooltip_key)
            tooltip_text = self.process_localisation_text(tooltip_text)
            if tooltip_text != tooltip_key:
                trait_data['effects'].append({
                    'description': f"【额外特质说明】{tooltip_text}",
                    'formatted_value': '',
                    'value_color': 'black',
                    'key': tooltip_key,
                    'value': ''
                })

        triggered_desc_pattern = r'triggered_desc\s*=\s*\{[^}]*text\s*=\s*"?([^"\s]+)"?[^}]*\}'
        triggered_desc_matches = re.findall(triggered_desc_pattern, content, re.DOTALL)
        for desc_key in triggered_desc_matches:
            desc_text = self.get_localisation(desc_key)
            if desc_text != desc_key:
                trait_data['extra_effects'].append(desc_text)

        opposites_match = re.search(r'opposites\s*=\s*\{([^}]+)\}', content, re.DOTALL)
        if opposites_match:
            opposites_content = opposites_match.group(1)
            opposites = re.findall(r'(\w+)', opposites_content)
            opposite_names = []
            for opp in opposites:
                localized_name = self.get_localisation(opp)
                if localized_name != opp:
                    opposite_names.append(localized_name)
                else:
                    opposite_names.append(opp)
            trait_data['opposites'] = opposite_names

        self.parse_trait_tier_opposites(trait_data, content)

        class_match = re.search(r'leader_class\s*=\s*\{([^}]+)\}', content, re.DOTALL)
        if class_match:
            classes_content = class_match.group(1)
            classes = re.findall(r'(\w+)', classes_content)
            trait_data['leader_classes'] = classes

        if 'randomized = no' in content:
            trait_data['randomized'] = False

        self.parse_leader_potential_add(trait_data, content)

        cleaned_content = self.remove_excluded_blocks(content)
        self.parse_leader_modifiers(trait_data, cleaned_content)

        has_effects = bool(trait_data['effects'] or trait_data['extra_effects'])
        if not has_effects:
            self.traits_without_effects.append(trait_key)

        self.leader_traits.append(trait_data)

    def parse_trait_tier_opposites(self, trait_data, content):
        """解析has_trait_tier1or2or3和has_trait_tier1or2条件中的互斥特质"""
        not_pattern = r'NOT\s*=\s*\{([^}]*?has_trait_tier(?:1or2or3|1or2)\s*=\s*\{[^}]*TRAIT\s*=\s*(\w+)[^}]*\}[^}]*?)\}'
        not_matches = re.findall(not_pattern, content, re.DOTALL)

        for not_content, trait in not_matches:
            trait_localized = self.get_localisation(trait)
            if trait_localized != trait:
                if trait_localized not in trait_data['opposites']:
                    trait_data['opposites'].append(trait_localized)

    def parse_leader_potential_add(self, trait_data, content):
        """解析leader_potential_add块"""
        pattern = r'leader_potential_add\s*=\s*\{'
        pos = 0

        while pos < len(content):
            match = re.search(pattern, content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            potential_block, new_pos = self.extract_balanced_braces(content, brace_start)

            if potential_block:
                if 'has_any_governor_focus = no' in potential_block:
                    governor_focus_traits = [
                        '矿业专精', '发电专精', '农学专精',
                        '工业专精', '冶金专精', '精炼专精'
                    ]

                    current_trait_name = trait_data['name']

                    for focus_trait in governor_focus_traits:
                        if focus_trait != current_trait_name:
                            if focus_trait not in trait_data['opposites']:
                                trait_data['opposites'].append(focus_trait)

                not_pattern = r'NOT\s*=\s*\{([^}]*?has_trait\s*=\s*(\w+)[^}]*?)\}'
                not_matches = re.findall(not_pattern, potential_block, re.DOTALL)

                for not_content, trait_key in not_matches:
                    trait_localized = self.get_localisation(trait_key)
                    if trait_localized != trait_key:
                        if trait_localized not in trait_data['opposites']:
                            trait_data['opposites'].append(trait_localized)

                self.parse_allowed_classes(trait_data, potential_block)
                self.parse_potential_conditions(trait_data, potential_block, 'add_conditions')
                self.parse_raw_potential_conditions(trait_data, potential_block, 'raw_add_conditions')

                pos = new_pos
            else:
                pos = brace_start + 1

    def parse_allowed_classes(self, trait_data, potential_block):
        """解析允许的职业"""
        allowed_classes = []

        if 'has_council_subclass = yes' in potential_block:
            council_classes = [
                'subclass_commander_councilor',
                'subclass_official_diplomacy_councilor',
                'subclass_official_economy_councilor',
                'subclass_scientist_councilor'
            ]
            for cls in council_classes:
                localized = self.get_localisation(cls)
                if localized != cls:
                    allowed_classes.append(localized)
                else:
                    allowed_classes.append(cls)

        elif 'has_governor_subclass = yes' in potential_block:
            governor_classes = [
                'subclass_commander_governor',
                'subclass_official_governor',
                'subclass_scientist_governor'
            ]
            for cls in governor_classes:
                localized = self.get_localisation(cls)
                if localized != cls:
                    allowed_classes.append(localized)
                else:
                    allowed_classes.append(cls)

        else:
            or_pattern = r'OR\s*=\s*\{([^}]+)\}'
            or_matches = re.findall(or_pattern, potential_block, re.DOTALL)

            if or_matches:
                for or_content in or_matches:
                    traits = re.findall(r'has_trait\s*=\s*(\w+)', or_content)
                    for trait in traits:
                        localized = self.get_localisation(trait)
                        if localized != trait:
                            allowed_classes.append(localized)
                        else:
                            allowed_classes.append(trait)
            else:
                traits = re.findall(r'has_trait\s*=\s*(\w+)', potential_block)
                for trait in traits:
                    localized = self.get_localisation(trait)
                    if localized != trait:
                        allowed_classes.append(localized)
                    else:
                        allowed_classes.append(trait)

        if allowed_classes:
            trait_data['allowed_classes'] = list(set(allowed_classes))

    def parse_potential_conditions(self, trait_data, potential_block, condition_type):
        """解析potential中的其他条件（翻译版本）"""
        conditions = self.parse_condition_block(potential_block, translate=True, context='add')

        if conditions:
            trait_data[condition_type] = conditions

    def parse_raw_potential_conditions(self, trait_data, potential_block, condition_type):
        """解析potential中的原始条件（未翻译版本）"""
        conditions = self.parse_condition_block(potential_block, translate=False, context='add')

        filtered_conditions = []
        for condition in conditions:
            if condition and condition.strip():
                filtered_conditions.append(condition)

        if filtered_conditions:
            trait_data[condition_type] = filtered_conditions

    def translate_condition(self, key, value, context='add', logic_context=None):
        """翻译条件"""
        if key in self.condition_translations:
            base_translation = self.condition_translations[key]

            if value in ['yes', 'no']:
                if context == 'add':
                    if value == 'yes':
                        return f"是{base_translation}"
                    else:
                        return f"不是{base_translation}"
                else:
                    if value == 'yes':
                        return f"{base_translation}"
                    else:
                        return f"非{base_translation}"
            else:
                value_localized = self.get_localisation(value)
                if value_localized == value:
                    value_localized = value

                if context == 'add':
                    return f"有{base_translation} {value_localized}"
                else:
                    return f"需要{base_translation} {value_localized}"

        return f"{key} = {value}"

    def parse_condition_block(self, block_content, translate=True, context='add', logic_context=None):
        """解析条件块"""
        conditions = []
        processed_content = block_content

        not_pattern = r'NOT\s*=\s*\{([^}]+)\}'
        not_matches = re.findall(not_pattern, processed_content, re.DOTALL)
        for not_content in not_matches:
            original_condition_count = self.count_conditions_in_content(not_content)
            filtered_not_content = self.filter_conditions_by_context(not_content, context)

            if filtered_not_content.strip():
                if translate:
                    not_conditions = self.parse_simple_conditions_with_count(
                        filtered_not_content, translate=True, context=context,
                        logic_context='NOT', parent_condition_count=original_condition_count
                    )
                    if not_conditions:
                        if len(not_conditions) > 1:
                            conditions.append(f"不是：（{', '.join(not_conditions)}）")
                        elif len(not_conditions) == 1:
                            condition = not_conditions[0]
                            if not condition.startswith("不是"):
                                condition = f"不是{condition}"
                            conditions.append(condition)
                else:
                    not_conditions = self.parse_simple_conditions_with_count(
                        filtered_not_content, translate=False, context=context,
                        logic_context='NOT', parent_condition_count=original_condition_count
                    )
                    if not_conditions:
                        if len(not_conditions) > 1:
                            conditions.append(f"NOT: （{'; '.join(not_conditions)}）")
                        elif len(not_conditions) == 1:
                            conditions.append(f"NOT: {not_conditions[0]}")

            processed_content = re.sub(not_pattern, '', processed_content, count=1, flags=re.DOTALL)

        original_condition_count = self.count_conditions_in_content(processed_content)
        filtered_content = self.filter_conditions_by_context(processed_content, context)
        simple_conditions = self.parse_simple_conditions_with_count(
            filtered_content, translate=translate, context=context,
            logic_context=logic_context, parent_condition_count=original_condition_count
        )
        conditions.extend(simple_conditions)

        return conditions

    def count_conditions_in_content(self, content):
        """计算内容中的条件数量"""
        if not content or not content.strip():
            return 0

        content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)
        simple_pattern = r'(\w+)\s*=\s*(\w+)'
        simple_matches = re.findall(simple_pattern, content)

        logic_patterns = [
            r'AND\s*=\s*\{[^}]+\}',
            r'OR\s*=\s*\{[^}]+\}',
            r'NOT\s*=\s*\{[^}]+\}',
            r'NOR\s*=\s*\{[^}]+\}',
            r'NAND\s*=\s*\{[^}]+\}'
        ]

        logic_count = 0
        for pattern in logic_patterns:
            logic_matches = re.findall(pattern, content, re.DOTALL)
            logic_count += len(logic_matches)

        total_count = len(simple_matches) + logic_count
        return total_count

    def parse_simple_conditions_with_count(self, content, translate=True, context='add', logic_context=None,
                                           parent_condition_count=1):
        """解析简单条件"""
        conditions = []
        filtered_content = self.filter_conditions_by_context(content, context)
        if not filtered_content.strip():
            return conditions

        generic_pattern = r'(\w+)\s*=\s*(\w+)'
        generic_matches = re.findall(generic_pattern, filtered_content)

        for key, value in generic_matches:
            if not key or not value:
                continue

            if key.startswith('has_') and key.endswith('_dlc'):
                if parent_condition_count > 1:
                    continue

            if translate:
                condition_text = self.translate_condition(key, value, context, logic_context)
                if condition_text and condition_text not in conditions:
                    conditions.append(condition_text)
            else:
                condition_text = f"{key} = {value}"
                if condition_text not in conditions:
                    conditions.append(condition_text)

        return conditions

    def filter_conditions_by_context(self, content, context='add'):
        """根据上下文过滤不需要的条件"""
        if not content:
            return content

        lines = content.split('\n')
        filtered_lines = []

        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line or line.startswith('#'):
                i += 1
                continue

            should_skip = False

            if line == 'exists = owner':
                should_skip = True
            elif re.match(r'has_trait\s*=\s*\w+', line):
                should_skip = True
            elif re.match(r'has_council_subclass\s*=\s*\w+', line) or re.match(r'has_governor_subclass\s*=\s*\w+',
                                                                               line):
                should_skip = True
            elif re.match(r'has_any_governor_focus\s*=\s*\w+', line):
                should_skip = True
            elif re.match(r'can_leader_get_council_trait_trigger\s*=\s*\w+', line):
                should_skip = True
            elif context == 'add' and re.match(r'has_\w+_dlc\s*=\s*yes', line):
                should_skip = True
            elif 'has_trait_tier' in line:
                if '{' in line:
                    brace_count = line.count('{') - line.count('}')
                    j = i
                    while j < len(lines) and brace_count > 0:
                        j += 1
                        if j < len(lines):
                            brace_count += lines[j].count('{') - lines[j].count('}')
                    should_skip = True
                    i = j
                else:
                    should_skip = True
            elif line.startswith('NOT = {') and 'has_trait_tier' in content[i:i + 10]:
                brace_count = 1
                j = i + 1
                while j < len(lines) and brace_count > 0:
                    if 'has_trait_tier' in lines[j]:
                        should_skip = True
                    if '{' in lines[j]:
                        brace_count += lines[j].count('{')
                    if '}' in lines[j]:
                        brace_count -= lines[j].count('}')
                    j += 1
                if should_skip:
                    i = j - 1

            if not should_skip:
                filtered_lines.append(line)

            i += 1

        return '\n'.join(filtered_lines)

    def parse_modifier_block(self, modifier_block):
        """解析修正块，只提取实际的修正效果"""
        effects = []

        lines = modifier_block.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()

            if not line or line.startswith('#'):
                i += 1
                continue

            if line.startswith('potential'):
                brace_count = line.count('{') - line.count('}')
                j = i
                while j < len(lines) and brace_count > 0:
                    j += 1
                    if j < len(lines):
                        brace_count += lines[j].count('{') - lines[j].count('}')
                i = j + 1
                continue

            if any(key in line for key in ['show_if_not_potential', 'not_potential_override_text_key']):
                i += 1
                continue

            if '=' in line and not line.startswith('potential'):
                parts = line.split('=', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    value = parts[1].strip().strip('"')

                    if key in ['potential', 'trigger', 'custom_tooltip', 'custom_tooltip_with_modifiers', 'factor']:
                        i += 1
                        continue

                    if key and value:
                        processed_value = self.extract_modifier_value(value)
                        formatted_value = self.format_modifier_value(key, processed_value)
                        value_color = self.get_value_color(key, processed_value)

                        chinese_desc = self.get_modifier_localisation(key)
                        effects.append({
                            'description': chinese_desc,
                            'formatted_value': formatted_value,
                            'value_color': value_color,
                            'key': key,
                            'value': processed_value
                        })

            i += 1

        return effects

    def parse_leader_modifiers(self, trait_data, content):
        """解析领袖特质的所有修正效果"""
        category_effects = {
            '通用效果': {},
            '统治者': {},
            '内阁成员': {},
            '行政官': {},
            '指挥官': {},
            '科学家': {}
        }

        processed_positions = {}

        triggered_modifier_types = [
            ('triggered_fleet_modifier', '军用舰队效果'),
            ('triggered_army_modifier', '陆军效果'),
            ('triggered_planet_modifier', '行星效果'),
            ('triggered_sector_modifier', '星域效果'),
            ('triggered_federation_modifier', '联邦效果'),
            ('triggered_galcom_modifier', '星海共同体效果'),
            ('triggered_councilor_modifier', '内阁效果'),
            ('triggered_self_modifier', '领袖自身效果'),
            ('triggered_modifier', '通用效果'),
            ('triggered_background_planet_modifier', '母星效果')
        ]

        for modifier_key, effect_type in triggered_modifier_types:
            pattern = rf'{modifier_key}\s*=\s*{{'
            pos = 0
            while pos < len(content):
                match = re.search(pattern, content[pos:])
                if not match:
                    break

                brace_start = pos + match.end() - 1
                modifier_block, new_pos = self.extract_balanced_braces(content, brace_start)

                if modifier_block and brace_start not in processed_positions:
                    conditions = self.extract_conditions_from_modifier_block(modifier_block)
                    effects = self.parse_modifier_block(modifier_block)
                    category, condition_marks = self.categorize_conditions(conditions)

                    if effects:
                        if condition_marks:
                            for effect in effects:
                                effect['description'] = f"（{condition_marks}）{effect['description']}"

                        if category not in category_effects:
                            category_effects[category] = {}

                        if effect_type not in category_effects[category]:
                            category_effects[category][effect_type] = []

                        category_effects[category][effect_type].extend(effects)
                        processed_positions[brace_start] = True

                pos = new_pos

        direct_modifier_types = [
            ('fleet_modifier', '军用舰队效果'),
            ('army_modifier', '陆军效果'),
            ('planet_modifier', '行星效果'),
            ('sector_modifier', '星域效果'),
            ('federation_modifier', '联邦效果'),
            ('galcom_modifier', '星海共同体效果'),
            ('councilor_modifier', '内阁效果'),
            ('self_modifier', '领袖自身效果')
        ]

        for modifier_key, effect_type in direct_modifier_types:
            pattern = rf'{modifier_key}\s*=\s*{{'
            pos = 0
            while pos < len(content):
                match = re.search(pattern, content[pos:])
                if not match:
                    break

                brace_start = pos + match.end() - 1
                modifier_block, new_pos = self.extract_balanced_braces(content, brace_start)

                if modifier_block and brace_start not in processed_positions:
                    effects = self.parse_modifier_block(modifier_block)
                    if effects:
                        if '通用效果' not in category_effects:
                            category_effects['通用效果'] = {}

                        if effect_type not in category_effects['通用效果']:
                            category_effects['通用效果'][effect_type] = []

                        category_effects['通用效果'][effect_type].extend(effects)
                        processed_positions[brace_start] = True

                pos = new_pos

        pattern = r'\bmodifier\s*=\s*{'
        pos = 0
        while pos < len(content):
            match = re.search(pattern, content[pos:])
            if not match:
                break

            brace_start = pos + match.end() - 1
            modifier_block, new_pos = self.extract_balanced_braces(content, brace_start)

            if modifier_block and brace_start not in processed_positions:
                effects = self.parse_modifier_block(modifier_block)
                if effects:
                    if '通用效果' not in category_effects:
                        category_effects['通用效果'] = {}

                    if '通用效果' not in category_effects['通用效果']:
                        category_effects['通用效果']['通用效果'] = []

                    category_effects['通用效果']['通用效果'].extend(effects)
                    processed_positions[brace_start] = True

            pos = new_pos

        trait_data['effects'].extend(self.merge_effects_by_category(category_effects))

        self.record_unhandled_modifier_conditions(trait_data, content)

    def extract_conditions_from_modifier_block(self, modifier_block):
        """从修正块中提取条件"""
        conditions = []
        potential_match = re.search(r'potential\s*=\s*{', modifier_block)
        if potential_match:
            brace_start = potential_match.end() - 1
            potential_block, _ = self.extract_balanced_braces(modifier_block, brace_start)

            if potential_block:
                cleaned_potential = re.sub(r'exists\s*=\s*FROM', '', potential_block)
                cleaned_potential = re.sub(r'FROM\s*=\s*\{[^}]*\}', '', cleaned_potential, flags=re.DOTALL)
                self.extract_conditions_from_text(cleaned_potential, conditions, context='modifier')

        return conditions

    def extract_conditions_from_text(self, text, conditions, context='add'):
        """从文本中提取条件"""
        added_conditions = set()
        specific_conditions = [
            ('is_ruler', '是统治者'),
            ('is_councilor', '是内阁成员'),
            ('leader_class', '领袖类型为')
        ]

        self._extract_not_conditions(text, conditions, added_conditions, context, specific_conditions)
        text_without_not = re.sub(r'NOT\s*=\s*\{[^}]+\}', '', text, flags=re.DOTALL)
        self._extract_conditions_from_content(text_without_not, conditions, added_conditions, context,
                                              specific_conditions)

        owner_pattern = r'owner\s*=\s*\{([^}]+)\}'
        owner_matches = re.findall(owner_pattern, text_without_not)
        for owner_content in owner_matches:
            self._extract_conditions_from_content(owner_content, conditions, added_conditions, context,
                                                  specific_conditions)

    def _extract_conditions_from_content(self, content, conditions, added_conditions, context, specific_conditions):
        """从内容中提取条件"""
        for condition_key, condition_text in specific_conditions:
            pattern = rf'{condition_key}\s*=\s*(\w+)'
            matches = re.findall(pattern, content)

            for match in matches:
                if condition_key == 'leader_class':
                    translated_condition = f"{condition_text}{match}"
                else:
                    if match == 'yes':
                        translated_condition = condition_text
                    else:
                        translated_condition = f"不是{condition_text.replace('是', '')}"

                if translated_condition and translated_condition not in added_conditions:
                    conditions.append(translated_condition)
                    added_conditions.add(translated_condition)

        for condition_key, condition_translation in self.condition_translations.items():
            if condition_key in ['is_ruler', 'is_councilor', 'leader_class']:
                continue

            pattern = rf'{condition_key}\s*=\s*(\w+)'
            matches = re.findall(pattern, content)

            for match in matches:
                condition_text = self.translate_condition(condition_key, match, context)

                if condition_text and condition_text not in added_conditions:
                    conditions.append(condition_text)
                    added_conditions.add(condition_text)

    def _extract_not_conditions(self, content, conditions, added_conditions, context, specific_conditions=None):
        """提取 NOT 块中的条件"""
        if specific_conditions is None:
            specific_conditions = []

        not_pattern = r'NOT\s*=\s*\{([^}]+)\}'
        not_matches = re.findall(not_pattern, content)

        for not_content in not_matches:
            for condition_key, condition_text in specific_conditions:
                pattern = rf'{condition_key}\s*=\s*(\w+)'
                matches = re.findall(pattern, not_content)

                for match in matches:
                    if condition_key == 'leader_class':
                        translated_condition = f"不是领袖类型为{match}"
                    else:
                        if match == 'yes':
                            translated_condition = f"不是{condition_text.replace('是', '')}"
                        else:
                            translated_condition = condition_text

                    if translated_condition and translated_condition not in added_conditions:
                        conditions.append(translated_condition)
                        added_conditions.add(translated_condition)

            for condition_key, condition_translation in self.condition_translations.items():
                if condition_key in ['is_ruler', 'is_councilor', 'leader_class']:
                    continue

                pattern = rf'{condition_key}\s*=\s*(\w+)'
                matches = re.findall(pattern, not_content)

                for match in matches:
                    if match == 'yes':
                        condition_text = f"不是{condition_translation}"
                    else:
                        value_localized = self.get_localisation(match)
                        if value_localized == match:
                            value_localized = match

                        condition_text = f"不是{condition_translation} {value_localized}"

                    if condition_text and condition_text not in added_conditions:
                        conditions.append(condition_text)
                        added_conditions.add(condition_text)

    def categorize_conditions(self, conditions):
        """根据条件分类"""
        has_ruler = False
        has_councilor = False
        leader_class = None
        condition_marks = []

        for condition in conditions:
            if condition == "是统治者":
                has_ruler = True
            elif condition == "是内阁成员":
                has_councilor = True
            elif condition.startswith("领袖类型为"):
                leader_class = condition.replace("领袖类型为", "").strip()
            else:
                condition_marks.append(condition)

        category = "通用效果"
        if has_ruler:
            category = "统治者"
        elif has_councilor:
            category = "内阁成员"
        elif leader_class:
            if leader_class == "official":
                category = "行政官"
            elif leader_class == "commander":
                category = "指挥官"
            elif leader_class == "scientist":
                category = "科学家"

        condition_marks_str = "；".join(condition_marks) if condition_marks else ""
        return category, condition_marks_str

    def merge_effects_by_category(self, category_effects):
        """按职业分类合并效果"""
        merged_effects = []
        category_order = ['通用效果', '统治者', '内阁成员', '行政官', '指挥官', '科学家']
        effect_type_order = ['通用效果', '母星效果', '行星效果', '星域效果', '领袖自身效果', '内阁效果',
                             '联邦效果', '星海共同体效果', '军用舰队效果', '陆军效果']

        for category in category_order:
            if category in category_effects and category_effects[category]:
                category_effect_parts = []

                for effect_type in effect_type_order:
                    if effect_type in category_effects[category] and category_effects[category][effect_type]:
                        direct_effects = []
                        triggered_effects = []

                        for effect_dict in category_effects[category][effect_type]:
                            if '（' in effect_dict['description'] and '）' in effect_dict['description']:
                                triggered_effects.append(effect_dict)
                            else:
                                direct_effects.append(effect_dict)

                        all_effects = direct_effects + triggered_effects

                        if all_effects:
                            type_effects = {
                                'category': category,
                                'effect_type': effect_type,
                                'effects': all_effects
                            }
                            category_effect_parts.append(type_effects)

                if category_effect_parts:
                    full_effect = {
                        'category': category,
                        'effect_parts': category_effect_parts
                    }
                    merged_effects.append(full_effect)

        return merged_effects

    def record_unhandled_modifier_conditions(self, trait_data, content):
        """记录未处理的修正条件"""
        potential_pattern = r'potential\s*=\s*\{([^}]+)\}'
        potential_matches = re.findall(potential_pattern, content, re.DOTALL)

        for potential_content in potential_matches:
            if 'owner = {' in potential_content:
                continue
            if 'has_trait =' in potential_content:
                continue

            conditions = self.parse_condition_block(potential_content, translate=False, context='modifier')
            filtered_conditions = []
            for condition in conditions:
                if condition.startswith('exists = owner'):
                    continue
                if re.search(r'has_\w+_dlc = yes', condition, re.IGNORECASE):
                    continue
                filtered_conditions.append(condition)

            if filtered_conditions:
                trait_data['raw_modifier_conditions'].extend(filtered_conditions)

    def process_all_traits(self):
        """处理所有特质文件"""
        print("开始处理特质文件...")

        self.load_colors()
        self.load_scripted_variables()
        self.load_localisation()
        self.load_scripted_loc()
        self.load_gfx_mappings()

        trait_files = list(Path(TRAITS_DIR).glob("*.txt"))
        for file_path in trait_files:
            self.parse_trait_file(file_path)

        print(f"解析完成！找到 {len(self.leader_traits)} 个领袖特质")
        self.export_results()

    def export_results(self):
        """导出结果到Excel"""
        if not self.leader_traits:
            print("没有找到领袖特质数据")
            return

        print("正在生成Excel文件...")
        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}领袖特质.xlsx')
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet('领袖特质')

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

        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 30)
        worksheet.set_column('G:G', 30)
        worksheet.set_column('H:H', 50)
        worksheet.set_column('I:I', 80)
        worksheet.set_column('J:J', 15)

        worksheet.set_default_row(60)

        headers = ['图片', '特质名称', '特质KEY', '特质类型', '内阁特质', '适用职业', '互斥特质', '特质额外效果',
                   '效果',
                   '可刷新']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        row = 1
        traits_with_images = 0
        for trait in self.leader_traits:
            worksheet.write(row, 0, '', cell_format)
            if trait['icon'] and os.path.exists(trait['icon']):
                try:
                    worksheet.insert_image(row, 0, trait['icon'], {
                        'x_offset': 5,
                        'y_offset': 5,
                        'x_scale': 0.6,
                        'y_scale': 0.6,
                        'object_position': 1
                    })
                    traits_with_images += 1
                except Exception:
                    pass

            worksheet.write(row, 1, trait['name'], cell_format)
            worksheet.write(row, 2, trait['key'], cell_format)
            worksheet.write(row, 3, trait['trait_type'], cell_format)
            worksheet.write(row, 4, trait['council_trait'], cell_format)
            worksheet.write(row, 5, ', '.join(trait['leader_classes']), cell_format)
            worksheet.write(row, 6, ', '.join(trait['opposites']), cell_format)
            worksheet.write(row, 7, '; '.join(trait['extra_effects']), description_format)
            worksheet.write(row, 9, '是' if trait['randomized'] else '否', cell_format)

            if trait['effects']:
                needs_rich_string = False
                for effect in trait['effects']:
                    if isinstance(effect, dict) and 'effect_parts' in effect:
                        for part in effect['effect_parts']:
                            for eff in part['effects']:
                                if eff['value_color'] != 'black' and eff['formatted_value']:
                                    needs_rich_string = True
                                    break
                    elif isinstance(effect, dict) and effect['value_color'] != 'black' and effect['formatted_value']:
                        needs_rich_string = True
                        break

                if needs_rich_string:
                    rich_text = []
                    for i, effect in enumerate(trait['effects']):
                        if i > 0:
                            rich_text.append(black_format)
                            rich_text.append('\n')

                        if isinstance(effect, dict) and 'effect_parts' in effect:
                            rich_text.append(black_format)
                            rich_text.append(f"【{effect['category']}】")

                            for j, part in enumerate(effect['effect_parts']):
                                if j > 0:
                                    rich_text.append(black_format)
                                    rich_text.append('；')

                                rich_text.append(black_format)
                                rich_text.append(f"{part['effect_type']}：")

                                for k, eff in enumerate(part['effects']):
                                    if k > 0:
                                        rich_text.append(black_format)
                                        rich_text.append('，')

                                    desc_text = eff['description'] if eff['description'] else " "
                                    rich_text.append(black_format)
                                    rich_text.append(desc_text)

                                    value_text = eff['formatted_value'] if eff['formatted_value'] else ""
                                    if value_text:
                                        if eff['value_color'] == 'green':
                                            rich_text.append(green_format)
                                        elif eff['value_color'] == 'red':
                                            rich_text.append(red_format)
                                        else:
                                            rich_text.append(black_format)
                                        rich_text.append(value_text)
                        elif isinstance(effect, dict):
                            desc_text = effect['description'] if effect['description'] else " "
                            value_text = effect['formatted_value'] if effect['formatted_value'] else ""

                            rich_text.append(black_format)
                            rich_text.append(desc_text)

                            if value_text:
                                if effect['value_color'] == 'green':
                                    rich_text.append(green_format)
                                elif effect['value_color'] == 'red':
                                    rich_text.append(red_format)
                                else:
                                    rich_text.append(black_format)
                                rich_text.append(value_text)
                        else:
                            rich_text.append(black_format)
                            rich_text.append(str(effect))

                    if rich_text and len(rich_text) >= 2:
                        try:
                            worksheet.write_rich_string(row, 8, *rich_text, description_format)
                        except Exception:
                            text_parts = []
                            for effect in trait['effects']:
                                if isinstance(effect, dict) and 'effect_parts' in effect:
                                    category_text = f"【{effect['category']}】"
                                    effect_parts = []
                                    for part in effect['effect_parts']:
                                        effect_type_text = f"{part['effect_type']}："
                                        effect_details = []
                                        for eff in part['effects']:
                                            effect_str = f"{eff['description']}{eff['formatted_value']}"
                                            effect_details.append(effect_str)
                                        effect_parts.append(f"{effect_type_text}{'，'.join(effect_details)}")
                                    text_parts.append(f"{category_text}{'；'.join(effect_parts)}")
                                elif isinstance(effect, dict):
                                    text_parts.append(f"{effect['description']}{effect['formatted_value']}")
                                else:
                                    text_parts.append(str(effect))
                            worksheet.write(row, 8, '\n'.join(text_parts), description_format)
                    else:
                        text_parts = []
                        for effect in trait['effects']:
                            if isinstance(effect, dict) and 'effect_parts' in effect:
                                category_text = f"【{effect['category']}】"
                                effect_parts = []
                                for part in effect['effect_parts']:
                                    effect_type_text = f"{part['effect_type']}："
                                    effect_details = []
                                    for eff in part['effects']:
                                        effect_str = f"{eff['description']}{eff['formatted_value']}"
                                        effect_details.append(effect_str)
                                    effect_parts.append(f"{effect_type_text}{'，'.join(effect_details)}")
                                text_parts.append(f"{category_text}{'；'.join(effect_parts)}")
                            elif isinstance(effect, dict):
                                text_parts.append(f"{effect['description']}{effect['formatted_value']}")
                            else:
                                text_parts.append(str(effect))
                        worksheet.write(row, 8, '\n'.join(text_parts), description_format)
                else:
                    text_parts = []
                    for effect in trait['effects']:
                        if isinstance(effect, dict) and 'effect_parts' in effect:
                            category_text = f"【{effect['category']}】"
                            effect_parts = []
                            for part in effect['effect_parts']:
                                effect_type_text = f"{part['effect_type']}："
                                effect_details = []
                                for eff in part['effects']:
                                    effect_str = f"{eff['description']}{eff['formatted_value']}"
                                    effect_details.append(effect_str)
                                effect_parts.append(f"{effect_type_text}{'，'.join(effect_details)}")
                            text_parts.append(f"{category_text}{'；'.join(effect_parts)}")
                        elif isinstance(effect, dict):
                            text_parts.append(f"{effect['description']}{effect['formatted_value']}")
                        else:
                            text_parts.append(str(effect))
                    worksheet.write(row, 8, '\n'.join(text_parts), description_format)
            else:
                worksheet.write(row, 8, "无效果", description_format)

            row += 1

        workbook.close()
        print(f"领袖特质已导出到: {output_path}")
        print(f"成功生成 {traits_with_images} 个特质的图片")

        self.export_unhandled_conditions()
        self.export_untranslated_modifiers()

    def export_unhandled_conditions(self):
        """导出未处理的限制条件"""
        output_path = os.path.join(OUTPUT_DIR, f'{VERSION}领袖限制条件.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("=== 领袖限制条件 ===\n\n")
            f.write("注意：此文件仅包含未在条件翻译字典中定义的限制条件。\n")
            f.write("已在字典中定义的条件（如has_ascension_perk等）不会在此列出。\n\n")

            untranslated_conditions = set()

            for trait in self.leader_traits:
                has_untranslated = False

                # 检查添加条件
                untranslated_add_conditions = []
                for condition in trait['raw_add_conditions']:
                    # 提取条件键
                    condition_key = self._extract_condition_key(condition)
                    if condition_key and condition_key not in self.condition_translations:
                        untranslated_add_conditions.append(condition)
                        untranslated_conditions.add(condition_key)

                # 检查修正条件
                untranslated_modifier_conditions = []
                for condition in trait['raw_modifier_conditions']:
                    # 提取条件键
                    condition_key = self._extract_condition_key(condition)
                    if condition_key and condition_key not in self.condition_translations:
                        untranslated_modifier_conditions.append(condition)
                        untranslated_conditions.add(condition_key)

                # 只输出包含未翻译条件的特质
                if untranslated_add_conditions or untranslated_modifier_conditions:
                    f.write(f"领袖特质key: {trait['key']}\n")
                    if untranslated_add_conditions:
                        f.write(f"未翻译的允许添加条件: {'; '.join(untranslated_add_conditions)}\n")
                    if untranslated_modifier_conditions:
                        f.write(f"未翻译的修正生效条件: {'; '.join(untranslated_modifier_conditions)}\n")
                    f.write("\n" + "-" * 30 + "\n\n")

            if untranslated_conditions:
                f.write("\n=== 未翻译的条件键 ===\n\n")
                f.write("以下条件键需要在条件翻译字典中添加翻译：\n\n")
                for condition_key in sorted(untranslated_conditions):
                    f.write(f"{condition_key}\n")

    def _extract_condition_key(self, condition_text):
        """从条件文本中提取条件键"""
        # 处理简单的 key = value 格式
        match = re.match(r'(\w+)\s*=\s*(\w+)', condition_text.strip())
        if match:
            return match.group(1)

        # 处理逻辑条件（NOT、OR等）
        logic_patterns = [
            r'NOT:\s*(\w+)\s*=\s*\w+',
            r'OR:\s*(\w+)\s*=\s*\w+',
            r'AND:\s*(\w+)\s*=\s*\w+'
        ]

        for pattern in logic_patterns:
            match = re.search(pattern, condition_text)
            if match:
                return match.group(1)

        return None

    def export_untranslated_modifiers(self):
        """导出未转换中文的修正条目到单独文件"""
        untranslated_modifiers = set()
        builtin_modifiers = MODIFIER_TRANSLATIONS

        for trait in self.leader_traits:
            for effect in trait['effects']:
                if isinstance(effect, dict) and 'effect_parts' in effect:
                    for part in effect['effect_parts']:
                        for eff in part['effects']:
                            if eff['key'] in builtin_modifiers:
                                continue
                            if eff['key'] not in self.modifier_localisation_dict:
                                prefixes = ['mod_', 'MOD_', 'MOD_COUNTRY_', 'MOD_FEDERATION_', 'MOD_LEADER_']
                                found = False
                                for prefix in prefixes:
                                    prefixed_key = f"{prefix}{eff['key']}"
                                    if prefixed_key in self.modifier_localisation_dict:
                                        found = True
                                        break
                                    if prefixed_key.lower() in self.modifier_localisation_dict:
                                        found = True
                                        break
                                if not found:
                                    untranslated_modifiers.add(eff['key'])
                elif isinstance(effect, dict):
                    if effect['key'] in builtin_modifiers:
                        continue
                    if effect['key'] not in self.modifier_localisation_dict:
                        prefixes = ['mod_', 'MOD_', 'MOD_COUNTRY_', 'MOD_FEDERATION_', 'MOD_LEADER_']
                        found = False
                        for prefix in prefixes:
                            prefixed_key = f"{prefix}{effect['key']}"
                            if prefixed_key in self.modifier_localisation_dict:
                                found = True
                                break
                            if prefixed_key.lower() in self.modifier_localisation_dict:
                                found = True
                                break
                        if not found:
                            untranslated_modifiers.add(effect['key'])

        if untranslated_modifiers:
            output_path = os.path.join(OUTPUT_DIR, f'{VERSION}未转换修正条目.txt')
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("=== 未转换中文的修正条目 ===\n\n")
                f.write("以下修正条目在本地化文件中找不到对应的中文翻译：\n\n")
                for modifier in sorted(untranslated_modifiers):
                    f.write(f"{modifier}\n")
                print(f"未转换修正条目已导出到: {output_path}")


def main():
    parser = StellarisTraitParser()
    parser.process_all_traits()
    print("处理完成！")


if __name__ == "__main__":
    main()