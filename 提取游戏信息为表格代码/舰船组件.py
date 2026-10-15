import os
import re
import csv
from pathlib import Path
import xlsxwriter

from A配置路径 import *

try:
    from 百分比转换字典 import NON_PERCENTAGE_KEYS, FORCE_PERCENTAGE_KEYS
except ImportError:
    NON_PERCENTAGE_KEYS = []
    FORCE_PERCENTAGE_KEYS = []

try:
    from 颜色转换字典 import SPECIAL_SUFFIXES
except ImportError:
    SPECIAL_SUFFIXES = []

try:
    from 修正翻译字典 import MODIFIER_TRANSLATIONS
except ImportError:
    MODIFIER_TRANSLATIONS = {}


COMPONENT_TYPE_LABELS = {
    "weapon_component_template": "武器",
    "utility_component_template": "通用/防御",
    "strike_craft_component_template": "舰载机",
    "component_template": "组件",
}

MUTATION_TYPE_LABELS = {
    "offensive": "攻击",
    "defensive": "防御",
    "aux": "辅助",
}

# 资源花费翻译字典（cost/upkeep 块中的资源键 -> 中文）
RESOURCE_CN = {
    "energy": "能量币",
    "minerals": "矿物",
    "food": "食物",
    "consumer_goods": "消费品",
    "alloys": "合金",
    "unity": "凝聚力",
    "influence": "影响力",
    "physics_research": "物理学研究",
    "society_research": "社会学研究",
    "engineering_research": "工程学研究",
    "volatile_motes": "易爆微粒",
    "exotic_gases": "异星天然气",
    "rare_crystals": "稀有水晶",
    "sr_living_metal": "活体金属",
    "sr_zro": "泽珞",
    "sr_dark_matter": "暗物质",
    "nanites": "纳米机器人",
    "minor_artifacts": "稀有文物",
}

# 组件槽位翻译字典（size -> 简称）
SLOT_CN = {
    "small": "S",
    "medium": "M",
    "large": "L",
    "extra_large": "X",
    "guided": "G",
    "point_defense": "P",
    "point_defence": "P",
    "hangar": "H",
    "titanic": "T",
    "aux": "A",
    "planet_killer": "W",
    "torpedo": "G",
}

WEAPON_CSV_CANDIDATES = [
    os.path.join(COMPONENT_TEMPLATES_DIR, "weapon_components.csv"),
    os.path.join(COMPONENT_TEMPLATES_DIR, "00_weapon_components.csv"),
    os.path.join(BASE_GAME_DIR, "common", "weapon_components.csv"),
]
INLINE_SCRIPTS_SHIP_DIR = os.path.join(INLINE_SCRIPTS_DIR, "ship_components")
MUTATIONS_DIR = os.path.join(BASE_GAME_DIR, "common", "mutations")


class ShipComponentExtractor:
    def __init__(self):
        self.localisation_dict = {}
        self.scripted_variables = {}
        self.weapon_csv_data = {}
        self.inline_script_cache = {}
        self.components = []
        self.mutations = []

        self.debug_info = {
            'missing_translations': set(),
            'missing_modifier_translations': set(),
            'missing_csv_keys': set(),
            'missing_mutation_translations': set(),
        }

        os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ================= 本地化 =================

    def clean_localisation_value(self, value):
        if '#' in value:
            value = value.split('#')[0].strip()
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]
        return value.strip()

    def preprocess_localisation(self, text):
        if not text:
            return text
        text = re.sub(r'\$t+\$', '', text)
        text = re.sub(r'\\n', '', text)
        return re.sub(r'\s+', '', text)

    def parse_localisation_files(self):
        print("正在解析本地化文件...")
        d = {}
        for file_path in Path(LOCALISATION_DIR).rglob("*.yml"):
            try:
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    for line in f:
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            parts = line.split(':', 1)
                            if len(parts) == 2:
                                k = parts[0].strip()
                                v = self.preprocess_localisation(
                                    self.clean_localisation_value(parts[1].strip()))
                                d[k] = v
            except Exception as e:
                print(f"解析本地化文件出错 {file_path}: {e}")
        self.localisation_dict = d
        print(f"已加载 {len(d)} 个本地化条目")

    def resolve_localisation(self, value, depth=0, max_depth=10):
        if not value or depth > max_depth:
            return value
        value = self.preprocess_localisation(value)
        for m in re.findall(r'\$([^$]+)\$', value):
            r = self.resolve_localisation(self.localisation_dict.get(m, m), depth + 1, max_depth)
            value = value.replace(f'${m}$', r)
        return value

    def find_component_translation(self, key):
        for c in (key, f"NAME_{key}", f"{key}_NAME", key.lower(), f"NAME_{key.lower()}"):
            if c in self.localisation_dict:
                return self.resolve_localisation(self.localisation_dict[c]), True
        return key, False

    def find_mutation_translation(self, key):
        """变异的本地化 key 通常就是 key 本身，或者带 MUTATION_ / _MUTATION 前后缀"""
        for c in (key, key.upper(), f"MUTATION_{key}", f"MUTATION_{key.upper()}",
                  f"{key}_MUTATION", f"mutation_{key}"):
            if c in self.localisation_dict:
                return self.resolve_localisation(self.localisation_dict[c]), True
        return key, False

    def find_modifier_translation(self, mk):
        if mk in MODIFIER_TRANSLATIONS:
            return MODIFIER_TRANSLATIONS[mk], True
        for k in (f"mod_{mk}", f"MOD_{mk.upper()}", f"mod_country_{mk}",
                  f"mod_planet_{mk}", f"mod_leader_{mk}", f"mod_federation_{mk}",
                  mk, mk.lower()):
            if k in self.localisation_dict:
                return self.resolve_localisation(self.localisation_dict[k]), True
        return mk, False

    # ================= 修正 =================

    def format_modifier_value(self, key, value):
        try:
            v = float(value)
        except (ValueError, TypeError):
            return str(value)

        def pct(k):
            if k in NON_PERCENTAGE_KEYS:
                return False
            if k in FORCE_PERCENTAGE_KEYS:
                return True
            return k.endswith(('_gain', '_perc', '_mult', '_mult_add', '_happiness',
                               '_speed', '_uninhabitable', '_reduction', '_habitability'))

        use_pct = pct(key)
        if key.endswith('_add') and key not in FORCE_PERCENTAGE_KEYS:
            use_pct = False
        if use_pct:
            x = v * 100
            s = f"{int(x)}%" if abs(x - round(x)) < 0.001 else f"{x:.2f}".rstrip('0').rstrip('.') + "%"
        else:
            s = f"{int(v)}" if v == int(v) else f"{v}"
        return f"{'+' if v >= 0 else ''}{s}"

    def get_value_color(self, key, value):
        try:
            v = float(value)
        except (ValueError, TypeError):
            return 'black'
        if any(key.endswith(s) for s in SPECIAL_SUFFIXES):
            return 'red' if v > 0 else 'green'
        return 'green' if v > 0 else 'red'

    def parse_modifier_block(self, inner):
        effects, missing = [], set()
        if not inner:
            return effects, missing
        for line in inner.split('\n'):
            line = line.strip()
            if not line or line.startswith('#') or '{' in line or '}' in line:
                continue
            m = re.match(r'(\w+)\s*=\s*(-?[\d.]+)', line)
            if not m:
                continue
            k, v = m.group(1), m.group(2)
            tr, found = self.find_modifier_translation(k)
            if not found:
                missing.add(k)
            fv = self.format_modifier_value(k, v)
            cc = {'green': 'G', 'red': 'R'}.get(self.get_value_color(k, v), '')
            effects.append(f"{tr}§{cc}{fv}§!" if cc else f"{tr}{fv}")
        return effects, missing

    # ================= 变量 / inline_script / CSV =================

    def load_scripted_variables(self):
        if not os.path.exists(SCRIPTED_VARIABLES_DIR):
            return
        for fn in os.listdir(SCRIPTED_VARIABLES_DIR):
            if not fn.endswith('.txt'):
                continue
            try:
                with open(os.path.join(SCRIPTED_VARIABLES_DIR, fn), 'r', encoding='utf-8-sig') as f:
                    for line in f:
                        m = re.match(r'\s*@(\w+)\s*=\s*(-?[\d.]+)', line)
                        if m:
                            self.scripted_variables[m.group(1)] = float(m.group(2))
            except Exception:
                pass

    def parse_at_variables_in_content(self, content):
        for m in re.finditer(r'^\s*@(\w+)\s*=\s*(-?[\d.]+)', content, re.MULTILINE):
            self.scripted_variables[m.group(1)] = float(m.group(2))

    def resolve_at_variables(self, text):
        def r1(m):
            var, num = m.group(1), float(m.group(2))
            if var in self.scripted_variables:
                r = self.scripted_variables[var] * num
                return str(int(r)) if r == int(r) else str(r)
            return m.group(0)
        text = re.sub(r'@\[\s*(\w+)\s*\*\s*([\d.-]+)\s*\]', r1, text)

        def r2(m):
            var = m.group(1)
            if var in self.scripted_variables:
                v = self.scripted_variables[var]
                return str(int(v)) if v == int(v) else str(v)
            return m.group(0)
        return re.sub(r'@(\w+)', r2, text)

    def load_inline_scripts(self):
        if not os.path.exists(INLINE_SCRIPTS_SHIP_DIR):
            return
        for fn in os.listdir(INLINE_SCRIPTS_SHIP_DIR):
            if fn.endswith('.txt'):
                try:
                    with open(os.path.join(INLINE_SCRIPTS_SHIP_DIR, fn),
                              'r', encoding='utf-8-sig') as f:
                        self.inline_script_cache[os.path.splitext(fn)[0]] = f.read()
                except Exception:
                    pass
        print(f"已加载 {len(self.inline_script_cache)} 个 inline_script")

    def expand_inline_script(self, block):
        pattern = r'inline_script\s*=\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
        for _ in range(5):
            m = re.search(pattern, block)
            if not m:
                break
            inner = m.group(1)
            sm = re.search(r'script\s*=\s*(\S+)', inner)
            if not sm:
                block = block[:m.start()] + block[m.end():]
                continue
            name = sm.group(1).split('/')[-1].strip()
            if name not in self.inline_script_cache:
                block = block[:m.start()] + block[m.end():]
                continue
            content = self.inline_script_cache[name]
            params = {k: v for k, v in re.findall(r'(\w+)\s*=\s*([^\s{}]+)', inner)
                      if k != 'script'}
            for k, v in params.items():
                content = content.replace(f'${k}$', v).replace(f'@{k}@', v)
            block = block[:m.start()] + content + block[m.end():]
        return block

    def load_weapon_csv(self):
        for path in WEAPON_CSV_CANDIDATES:
            if not os.path.exists(path):
                continue
            try:
                with open(path, 'r', encoding='utf-8-sig', newline='') as f:
                    for row in csv.DictReader(f):
                        key = row.get('key') or row.get('Key')
                        if key:
                            self.weapon_csv_data[key] = {
                                k: v for k, v in row.items() if k and v != ''
                            }
                print(f"已加载 {len(self.weapon_csv_data)} 条武器数值 ({path})")
                return
            except Exception as e:
                print(f"加载 CSV 失败 {path}: {e}")
        print("未找到 weapon_components.csv（普通武器数值列可能为空）")

    def csv_field(self, row, *names):
        if not row:
            return ""
        for n in names:
            if n and n in row and row[n]:
                return row[n]
        return ""

    # ================= 通用解析 =================

    def extract_complete_block(self, content, start_pos):
        bc, in_q, esc = 0, False, False
        for i in range(start_pos, len(content)):
            if content[i] == '{':
                start_pos = i
                break
        for i in range(start_pos, len(content)):
            c = content[i]
            if esc:
                esc = False
                continue
            if c == '\\':
                esc = True
                continue
            if c == '"':
                in_q = not in_q
                continue
            if not in_q:
                if c == '{':
                    bc += 1
                elif c == '}':
                    bc -= 1
                    if bc == 0:
                        return content[start_pos:i + 1]
        return None

    def get_num_field(self, block, field, default=""):
        m = re.search(rf'(?<![A-Za-z0-9_]){re.escape(field)}\s*=\s*(-?[\d.]+)', block)
        return m.group(1) if m else default

    def get_str_field(self, block, field, default=""):
        m = re.search(rf'(?<![A-Za-z0-9_]){re.escape(field)}\s*=\s*"([^"]*)"', block)
        if m:
            return m.group(1)
        m = re.search(rf'(?<![A-Za-z0-9_]){re.escape(field)}\s*=\s*([^\s{{}}#]+)', block)
        return m.group(1) if m else default

    def get_block(self, block, name):
        m = re.search(rf'(?<![A-Za-z0-9_]){re.escape(name)}\s*=\s*\{{', block)
        if not m:
            return None
        full = self.extract_complete_block(block, m.start())
        if not full:
            return None
        return full[full.find('{') + 1:-1]

    def parse_kv_pairs(self, text):
        if not text:
            return {}
        return {k: v for k, v in re.findall(r'(\w+)\s*=\s*(-?[\d.]+)', text)}

    def format_resources(self, d):
        if not d:
            return ""
        out = []
        for k, v in d.items():
            try:
                fv = float(v)
                out.append(f"{RESOURCE_CN.get(k, k)}:{int(fv) if fv == int(fv) else fv}")
            except (ValueError, TypeError):
                out.append(f"{RESOURCE_CN.get(k, k)}:{v}")
        return ", ".join(out)

    # ================= 组件解析 =================

    def parse_component_files(self):
        print("正在解析舰船组件文件...")
        components = []
        pattern = (r'(weapon_component_template|utility_component_template|'
                   r'strike_craft_component_template|component_template)\s*=\s*\{')

        for file_path in Path(COMPONENT_TEMPLATES_DIR).rglob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    content = f.read()
                self.parse_at_variables_in_content(content)
                content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                for match in re.finditer(pattern, content):
                    template_type = match.group(1)
                    block = self.extract_complete_block(content, match.start())
                    if not block:
                        continue
                    block = self.expand_inline_script(block)
                    block = self.resolve_at_variables(block)

                    key = self.get_str_field(block, 'key')
                    if not key:
                        continue

                    csv_row = self.weapon_csv_data.get(key, {})
                    size = self.get_str_field(block, 'size')
                    weapon_type = self.get_str_field(block, 'type')
                    power = self.get_num_field(block, 'power')

                    res_block = self.get_block(block, 'resources')
                    cost_dict, upkeep_dict = {}, {}
                    if res_block:
                        cost_dict = self.parse_kv_pairs(self.get_block(res_block, 'cost'))
                        upkeep_dict = self.parse_kv_pairs(self.get_block(res_block, 'upkeep'))

                    def F(*names):
                        v = self.csv_field(csv_row, *names)
                        if v:
                            return v
                        for n in names:
                            v2 = self.get_num_field(block, n)
                            if v2:
                                return v2
                        return ""

                    dmg_block = self.get_block(block, 'damage')
                    dmg = self.parse_kv_pairs(dmg_block) if dmg_block else {}
                    damage_min = dmg.get('min', '') or F('damage_min', 'min_damage')
                    damage_max = dmg.get('max', '') or F('damage_max', 'max_damage')

                    wu_block = self.get_block(block, 'windup')
                    wu = self.parse_kv_pairs(wu_block) if wu_block else {}
                    windup_min = wu.get('min', '') or F('windup_min')
                    windup_max = wu.get('max', '') or F('windup_max')

                    hull_damage = F('hull_damage')
                    shield_damage = F('shield_damage')
                    shield_penetration = F('shield_penetration')
                    armor_damage = F('armor_damage')
                    armor_penetration = F('armor_penetration')
                    size_damage = F('size_damage_factor', 'size_damage')
                    total_fire_time = F('total_fire_time', 'cooldown')
                    display_fire_time = F('display_fire_time')
                    fire_arc = F('fire_arc', 'firing_arc')
                    min_range = F('min_range')
                    max_range = F('max_range', 'range', 'engagement_range')
                    if not min_range and max_range:
                        min_range = "0"
                    accuracy = F('accuracy')
                    tracking = F('tracking')
                    missile_speed = F('missile_speed')
                    missile_evasion = F('missile_evasion')
                    missile_shield = F('missile_shield')
                    missile_armor = F('missile_armor')
                    missile_hull = F('missile_health', 'missile_hull')
                    missile_retarget_range = F('missile_retarget_range')
                    military_power = F('military_power', 'military_power_multiplier')

                    dpd = ""
                    try:
                        dmin = float(damage_min) if damage_min else 0
                        dmax = float(damage_max) if damage_max else 0
                        cd = float(total_fire_time) if total_fire_time else 0
                        if cd > 0:
                            dpd = f"{(dmin + dmax) / 2 / cd:.2f}"
                    except (ValueError, TypeError):
                        pass

                    modifier_effects = []
                    mod_inner = self.get_block(block, 'modifier')
                    if mod_inner:
                        modifier_effects, miss = self.parse_modifier_block(mod_inner)
                        self.debug_info['missing_modifier_translations'].update(miss)

                    if template_type == 'strike_craft_component_template':
                        for fn, label in [('count', '编队数量'), ('regeneration_per_day', '每日补充'),
                                          ('launch_time', '起飞时间'), ('health', '生命'),
                                          ('armor', '装甲'), ('shield', '护盾'),
                                          ('evasion', '闪避'), ('speed', '速度'),
                                          ('rotation_speed', '转向速度'), ('acceleration', '加速度')]:
                            v = self.get_num_field(block, fn)
                            if v:
                                modifier_effects.append(f"{label}={v}")

                    name, found = self.find_component_translation(key)
                    if not found:
                        self.debug_info['missing_translations'].add(key)
                    if template_type == 'weapon_component_template' and not csv_row:
                        self.debug_info['missing_csv_keys'].add(key)

                    components.append({
                        'name': name, 'key': key, 'size': size,
                        'component_type_label': COMPONENT_TYPE_LABELS.get(template_type, template_type),
                        'weapon_type': weapon_type,
                        'cost': self.format_resources(cost_dict),
                        'upkeep': self.format_resources(upkeep_dict),
                        'power': power,
                        'damage_min': damage_min, 'damage_max': damage_max,
                        'hull_damage': hull_damage, 'shield_damage': shield_damage,
                        'shield_penetration': shield_penetration,
                        'armor_damage': armor_damage, 'armor_penetration': armor_penetration,
                        'windup_min': windup_min, 'windup_max': windup_max,
                        'total_fire_time': total_fire_time,
                        'display_fire_time': display_fire_time,
                        'dpd': dpd, 'fire_arc': fire_arc,
                        'min_range': min_range, 'max_range': max_range,
                        'accuracy': accuracy, 'tracking': tracking,
                        'size_damage': size_damage, 'military_power': military_power,
                        'missile_speed': missile_speed, 'missile_evasion': missile_evasion,
                        'missile_shield': missile_shield, 'missile_armor': missile_armor,
                        'missile_hull': missile_hull,
                        'missile_retarget_range': missile_retarget_range,
                        'modifier_effects': modifier_effects,
                        'original_index': len(components),
                    })

            except Exception as e:
                print(f"解析组件文件出错 {file_path}: {e}")

        self.components = components
        print(f"已找到 {len(components)} 个舰船组件定义")

    # ================= 变异解析 =================

    def parse_mutation_files(self):
        """解析 common/mutations/*.txt（mutations.txt、unique_mutations.txt 等）"""
        print("正在解析变异文件...")
        if not os.path.exists(MUTATIONS_DIR):
            print(f"变异目录不存在: {MUTATIONS_DIR}")
            return

        mutations = []
        for file_path in Path(MUTATIONS_DIR).rglob("*.txt"):
            try:
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    content = f.read()

                # 记录文件级 @变量
                self.parse_at_variables_in_content(content)

                # 去掉注释和 README 的注释块
                content = re.sub(r'#.*$', '', content, flags=re.MULTILINE)

                # mutation_key = { ... }
                for m in re.finditer(r'^(\w+)\s*=\s*\{', content, re.MULTILINE):
                    mut_key = m.group(1)
                    # 跳过 README 示例（一般不会出现，这里防御性检查）
                    if mut_key in ('mutation_key',):
                        continue

                    block = self.extract_complete_block(content, m.start())
                    if not block:
                        continue

                    mut_type = self.get_str_field(block, 'type')
                    icon = self.get_str_field(block, 'icon')
                    upgrades_to = self.get_str_field(block, 'upgrades_to')
                    is_unique = self.get_str_field(block, 'is_unique', 'no')
                    ship_limit = self.get_num_field(block, 'ship_limit')
                    disable_slot = self.get_str_field(block, 'component_slot_size_to_disable')
                    custom_tooltip = self.get_str_field(block, 'custom_tooltip')

                    # ship_components 子块
                    sc_block = self.get_block(block, 'ship_components')
                    size_map = {}
                    if sc_block:
                        for s in ('small', 'medium', 'large', 'extra_large'):
                            v = self.get_str_field(sc_block, s)
                            if v:
                                size_map[s] = v

                    name, found = self.find_mutation_translation(mut_key)
                    if not found:
                        self.debug_info['missing_mutation_translations'].add(mut_key)

                    # 处理 custom_tooltip 的本地化
                    tooltip_text = ""
                    if custom_tooltip:
                        tt = self.localisation_dict.get(custom_tooltip, custom_tooltip)
                        tooltip_text = self.resolve_localisation(tt)

                    mutations.append({
                        'key': mut_key,
                        'name': name,
                        'type': mut_type,
                        'type_label': MUTATION_TYPE_LABELS.get(mut_type, mut_type),
                        'icon': icon,
                        'small': size_map.get('small', ''),
                        'medium': size_map.get('medium', ''),
                        'large': size_map.get('large', ''),
                        'extra_large': size_map.get('extra_large', ''),
                        'upgrades_to': upgrades_to,
                        'is_unique': is_unique,
                        'ship_limit': ship_limit,
                        'disable_slot': disable_slot,
                        'tooltip': tooltip_text,
                        'source': os.path.basename(file_path),
                        'original_index': len(mutations),
                    })

            except Exception as e:
                print(f"解析变异文件出错 {file_path}: {e}")

        self.mutations = mutations
        print(f"已找到 {len(mutations)} 个变异定义")

    # ================= Excel 输出 =================

    def process_text_with_color_markers(self, text, workbook):
        red = workbook.add_format({'font_color': 'red', 'text_wrap': True, 'valign': 'vcenter'})
        green = workbook.add_format({'font_color': 'green', 'text_wrap': True, 'valign': 'vcenter'})
        black = workbook.add_format({'font_color': 'black', 'text_wrap': True, 'valign': 'vcenter'})

        if not text or ('§R' not in text and '§G' not in text):
            return [(black, text)]

        parts, i, cur, fmt = [], 0, "", black
        while i < len(text):
            if text[i:i + 2] in ('§R', '§G') and text.find('§!', i + 2) != -1:
                if cur:
                    parts.append((fmt, cur)); cur = ""
                end = text.find('§!', i + 2)
                cf = red if text[i:i + 2] == '§R' else green
                parts.append((cf, text[i + 2:end]))
                i = end + 2
                fmt = black
            else:
                cur += text[i]; i += 1
        if cur:
            parts.append((fmt, cur))
        return parts

    def _write_component_sheet(self, workbook, ws):
        hf = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
                                  'align': 'center', 'valign': 'vcenter'})
        cf = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        wf = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})

        headers = ["名称", "代码", "槽位", "花费", "能源",
                   "最小伤害", "最大伤害", "船体伤害", "护盾伤害", "护盾穿透",
                   "装甲伤害", "装甲穿透", "最小发射准备时间", "最大发射准备时间",
                   "冷却", "显示冷却", "DPD", "射击角度", "最小攻击范围", "最大攻击范围",
                   "命中", "追踪", "尺寸增伤", "战斗力计算",
                   "导弹飞行速度", "导弹闪避概率", "导弹护盾", "导弹装甲",
                   "导弹生命", "导弹重定向范围", "额外效果"]
        widths = [28, 28, 10, 22, 8, 10, 10, 10, 10, 10, 10, 10, 14, 14,
                  10, 10, 10, 10, 12, 12, 8, 8, 10, 12, 12, 12, 10, 10, 10, 14, 45]
        for i, w in enumerate(widths):
            ws.set_column(i, i, w)
        for c, h in enumerate(headers):
            ws.write(0, c, h, hf)

        for r, comp in enumerate(self.components, 1):
            vals = [comp['name'], comp['key'],
                    SLOT_CN.get(comp['size'].lower(), comp['size']) if comp['size'] else "",
                    comp['cost'], comp['power'],
                    comp['damage_min'], comp['damage_max'], comp['hull_damage'],
                    comp['shield_damage'], comp['shield_penetration'],
                    comp['armor_damage'], comp['armor_penetration'],
                    comp['windup_min'], comp['windup_max'],
                    comp['total_fire_time'], comp['display_fire_time'], comp['dpd'],
                    comp['fire_arc'], comp['min_range'], comp['max_range'],
                    comp['accuracy'], comp['tracking'], comp['size_damage'],
                    comp['military_power'],
                    comp['missile_speed'], comp['missile_evasion'],
                    comp['missile_shield'], comp['missile_armor'],
                    comp['missile_hull'], comp['missile_retarget_range']]
            for c, v in enumerate(vals):
                ws.write(r, c, v if v != "" else "", cf)

            meffs = comp.get('modifier_effects', [])
            if meffs:
                sep = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})
                args = []
                for i, e in enumerate(meffs):
                    if i > 0:
                        args.append(sep); args.append('，')
                    for fmt, t in self.process_text_with_color_markers(e, workbook):
                        args.append(fmt); args.append(t)
                try:
                    ws.write_rich_string(r, 30, *args)
                except Exception:
                    plain = '，'.join(meffs).replace('§R', '').replace('§G', '').replace('§!', '')
                    ws.write(r, 30, plain, wf)
            else:
                ws.write(r, 30, "", wf)

    def _write_mutation_sheet(self, workbook, ws):
        hf = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1,
                                  'align': 'center', 'valign': 'vcenter'})
        cf = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        wf = workbook.add_format({'border': 1, 'valign': 'vcenter', 'text_wrap': True})

        headers = ["变异名称", "变异KEY", "类型", "图标",
                   "小型组件", "中型组件", "大型组件", "超大型组件",
                   "升级到", "唯一", "数量限制", "禁用槽位", "自定义提示", "来源"]
        widths = [26, 32, 8, 34, 30, 30, 30, 32,
                  20, 6, 8, 12, 40, 26]
        for i, w in enumerate(widths):
            ws.set_column(i, i, w)
        for c, h in enumerate(headers):
            ws.write(0, c, h, hf)

        for r, mut in enumerate(self.mutations, 1):
            vals = [mut['name'], mut['key'], mut['type_label'], mut['icon'],
                    mut['small'], mut['medium'], mut['large'], mut['extra_large'],
                    mut['upgrades_to'], mut['is_unique'], mut['ship_limit'],
                    mut['disable_slot'], mut['tooltip'], mut['source']]
            for c, v in enumerate(vals):
                ws.write(r, c, v if v != "" else "", cf if c != 12 else wf)

    def save_to_excel(self, data):
        if not data and not self.mutations:
            print("没有数据可导出")
            return
        path = os.path.join(OUTPUT_DIR, f"{VERSION}舰船组件.xlsx")
        wb = xlsxwriter.Workbook(path)

        ws1 = wb.add_worksheet("舰船组件")
        self._write_component_sheet(wb, ws1)

        ws2 = wb.add_worksheet("变异")
        self._write_mutation_sheet(wb, ws2)

        wb.close()
        print(f"Excel 已保存: {path}")

    def save_debug_info(self):
        lines = []
        if self.debug_info['missing_translations']:
            lines.append("以下组件找不到本地化名称：")
            lines += [f"  - {k}" for k in sorted(self.debug_info['missing_translations'])]
            lines.append("")
        if self.debug_info['missing_modifier_translations']:
            lines.append("以下修正找不到本地化名称：")
            lines += [f"  - {k}" for k in sorted(self.debug_info['missing_modifier_translations'])]
            lines.append("")
        if self.debug_info['missing_csv_keys']:
            lines.append("以下武器在 weapon_components.csv 中找不到（数值可能不全）：")
            lines += [f"  - {k}" for k in sorted(self.debug_info['missing_csv_keys'])]
            lines.append("")
        if self.debug_info['missing_mutation_translations']:
            lines.append("以下变异找不到本地化名称：")
            lines += [f"  - {k}" for k in sorted(self.debug_info['missing_mutation_translations'])]
            lines.append("")
        if not lines:
            lines.append("没有发现调试信息")

        path = os.path.join(OUTPUT_DIR, f'{VERSION}舰船组件调试.txt')
        with open(path, 'w', encoding='utf-8') as f:
            f.write("=== 舰船组件调试信息 ===\n\n" + '\n'.join(lines))
        print(f"调试信息已保存: {path}")

    def run(self):
        print("开始提取舰船组件信息...")
        self.load_scripted_variables()
        self.load_inline_scripts()
        self.load_weapon_csv()
        self.parse_localisation_files()
        self.parse_component_files()
        self.parse_mutation_files()
        self.save_to_excel(self.components)
        self.save_debug_info()
        print("提取完成！")


if __name__ == "__main__":
    ShipComponentExtractor().run()