"""恒星基地 / 星环基地 / 方舟 / 驿站 模块与建筑提取脚本。

数据来源：
  common/starbase_modules/*.txt
  common/starbase_buildings/*.txt

输出四件套：TXT / JSON / CSV / XLSX（按类型分工作表）。
"""

import csv
import json
import os
import re
from pathlib import Path

import xlsxwriter

from A配置路径 import (
    BASE_GAME_DIR,
    LOCALISATION_DIR,
    OUTPUT_DIR,
    SCRIPTED_VARIABLES_DIR,
    VERSION,
)

try:
    from 修正翻译字典 import MODIFIER_TRANSLATIONS
except ImportError:
    MODIFIER_TRANSLATIONS = {}

try:
    from 百分比转换字典 import FORCE_PERCENTAGE_KEYS, NON_PERCENTAGE_KEYS
except ImportError:
    NON_PERCENTAGE_KEYS = []
    FORCE_PERCENTAGE_KEYS = []

STARBASE_MODULE_DIR = os.path.join(BASE_GAME_DIR, "common", "starbase_modules")
STARBASE_BUILDING_DIR = os.path.join(BASE_GAME_DIR, "common", "starbase_buildings")
INLINE_SCRIPTS_DIR = os.path.join(BASE_GAME_DIR, "common", "inline_scripts")

RESOURCE_CN = {
    "energy": "能量币",
    "minerals": "矿物",
    "food": "食物",
    "consumer_goods": "消费品",
    "alloys": "合金",
    "trade": "贸易",
    "unity": "凝聚力",
    "influence": "影响力",
    "physics_research": "物理学研究",
    "society_research": "社会学研究",
    "engineering_research": "工程学研究",
    "volatile_motes": "易爆微粒",
    "exotic_gases": "异星天然气",
    "rare_crystals": "稀有水晶",
    "minor_artifacts": "稀有文物",
    "astral_threads": "星界丝缕",
    "sr_living_metal": "活体金属",
    "sr_zro": "泽珞",
    "sr_dark_matter": "暗物质",
    "nanites": "纳米机器人",
}

CATEGORY_CN = {
    "sm_naval_support_cat": "海军支援",
    "sm_economy_cat": "经济",
    "sm_utility_cat": "通用",
    "sm_defense_cat": "防御",
    "sm_colony_cat": "殖民",
    "sm_enclave_cat": "飞地",
    "sm_aura_cat": "光环",
    "sm_energy_cat": "能量武器",
    "sm_strike_craft_cat": "舰载机",
    "sm_spinal_cat": "脊柱挂载",
}

PLATFORM_TYPE_CN = {
    "starbase": "恒星基地",
    "orbital_ring": "星环基地",
    "arkship": "方舟",
    "waystation": "驿站",
}

MODIFIER_PRIORITY = {
    "planet_modifier": 100,
    "country_modifier": 95,
    "station_modifier": 90,
    "system_modifier": 85,
    "orbit_modifier": 84,
    "ship_modifier": 83,
    "defense_platform_modifier": 82,
    "waystation_network_country_modifier": 80,
    "waystation_network_system_modifier": 79,
    "triggered_planet_modifier": 75,
    "triggered_country_modifier": 74,
    "triggered_station_modifier": 73,
    "triggered_system_modifier": 72,
    "triggered_waystation_network_country_modifier": 71,
    "triggered_waystation_network_system_modifier": 70,
}

BLOCK_HEAD = re.compile(r'([A-Za-z0-9_]+)[ \t\r\n]*=[ \t\r\n]*\{')
SCALAR_PAIR = re.compile(r'([A-Za-z0-9_]+)[ \t\r\n]*=[ \t\r\n]*(-?[\d.]+)')
INLINE_HEAD = re.compile(r'(?<![A-Za-z0-9_])inline_script[ \t\r\n]*=[ \t\r\n]*\{')

# 这些键是脚本结构键，不能当成修正值输出。
NON_MODIFIER_KEYS = {
    "custom_tooltip", "show_only_custom_tooltip", "custom_tooltip_with_modifiers",
    "potential", "possible", "mult", "factor", "add", "base", "value", "trigger",
    "limit", "weight", "count", "script", "type", "amount", "MODULE", "ZONE",
}


class ClausewitzParser:
    """极简 Clausewitz 脚本解析器（字符串 / 转义 / 花括号配对安全）。"""

    def __init__(self):
        self.scripted_variables = {}
        self.inline_cache = {}
        self.localisation = {}

    # ---------- 文本级工具 ----------

    @staticmethod
    def strip_comments(text):
        out = []
        i = 0
        n = len(text)
        in_string = False
        while i < n:
            c = text[i]
            if c == '\\' and i + 1 < n:
                out.append(text[i:i + 2])
                i += 2
                continue
            if c == '"':
                in_string = not in_string
                out.append(c)
                i += 1
                continue
            if c == '#' and not in_string:
                while i < n and text[i] != '\n':
                    i += 1
                out.append('\n')
                continue
            out.append(c)
            i += 1
        return ''.join(out)

    @staticmethod
    def skip_string(text, i):
        """i 位于引号起始处，返回字符串结束后的下标。"""
        n = len(text)
        i += 1
        while i < n:
            if text[i] == '\\':
                i += 2
                continue
            if text[i] == '"':
                return i + 1
            i += 1
        return n

    @staticmethod
    def find_block_span(text, start):
        """从 start 起找第一个 '{'，返回 (起始下标, 结束下标)。"""
        brace_start = text.find('{', start)
        if brace_start < 0:
            return None
        depth = 0
        in_string = False
        i = brace_start
        n = len(text)
        while i < n:
            c = text[i]
            if c == '\\' and i + 1 < n:
                i += 2
                continue
            if c == '"':
                in_string = not in_string
                i += 1
                continue
            if not in_string:
                if c == '{':
                    depth += 1
                elif c == '}':
                    depth -= 1
                    if depth == 0:
                        return brace_start, i + 1
            i += 1
        return None

    @staticmethod
    def find_block_at(text, start):
        span = ClausewitzParser.find_block_span(text, start)
        return text[span[0]:span[1]] if span else None

    @staticmethod
    def inner(block):
        if not block:
            return ''
        first = block.find('{')
        last = block.rfind('}')
        return block[first + 1:last].strip() if first >= 0 and last > first else ''

    def extract_blocks(self, text):
        """提取顶层 `key = { ... }` 块，返回 [(key, full, inner)]。"""
        text = self.strip_comments(text)
        results = []
        i = 0
        n = len(text)
        while i < n:
            c = text[i]
            if c == '"':
                i = self.skip_string(text, i)
                continue
            m = BLOCK_HEAD.match(text, i)
            if m:
                span = self.find_block_span(text, m.start())
                if span:
                    begin, end = span
                    full = text[begin:end]
                    results.append((m.group(1), full, self.inner(full)))
                    i = end
                    continue
            i += 1
        return results

    def direct_block(self, text, name):
        """返回体部中名为 name 的直接子块内容。

        注意：不能只看第一个块 —— 例如驿站模块的 starbase_limit 排在 potential 之前。
        """
        for key, _full, inner in self.extract_blocks(text):
            if key == name:
                return inner
        return ''

    def scalar_pairs(self, text):
        """提取 `key = 数值`，自动跳过字符串与嵌套块（如 mult = value:xxx|...）。"""
        text = self.strip_comments(text)
        pairs = []
        i = 0
        n = len(text)
        while i < n:
            c = text[i]
            if c == '"':
                i = self.skip_string(text, i)
                continue
            m = SCALAR_PAIR.match(text, i)
            if m:
                pairs.append((m.group(1), m.group(2)))
                i = m.end()
                continue
            b = BLOCK_HEAD.match(text, i)
            if b:
                span = self.find_block_span(text, b.start())
                if span:
                    i = span[1]
                    continue
            i += 1
        return pairs

    # ---------- 字段级工具 ----------

    def str_field(self, text, key):
        m = re.search(rf'(?<![A-Za-z0-9_]){re.escape(key)}\s*=\s*"([^"]*)"', text)
        if m:
            return m.group(1)
        m = re.search(rf'(?<![A-Za-z0-9_]){re.escape(key)}\s*=\s*([^\s{{}}#]+)', text)
        return m.group(1) if m else ''

    def str_list(self, text, key):
        out = []
        for m in re.finditer(
            rf'(?<![A-Za-z0-9_]){re.escape(key)}\s*=\s*(?:"([^"]*)"|([^\s{{}}#]+))', text
        ):
            out.append(m.group(1) if m.group(1) is not None else m.group(2))
        return out

    @staticmethod
    def has_key(text, key):
        return re.search(rf'(?<![A-Za-z0-9_]){re.escape(key)}\s*=', text) is not None

    # ---------- 变量 / inline_script ----------

    def resolve_at(self, text):
        def repl(match):
            value = self.scripted_variables.get(match.group(1))
            if value is None:
                return match.group(0)
            return str(int(value)) if float(value).is_integer() else f"{value:g}"
        return re.sub(r'@([A-Za-z0-9_]+)', repl, text)

    def load_scripted_variables(self):
        if not os.path.isdir(SCRIPTED_VARIABLES_DIR):
            return
        for fn in sorted(os.listdir(SCRIPTED_VARIABLES_DIR)):
            if not fn.endswith('.txt'):
                continue
            try:
                with open(os.path.join(SCRIPTED_VARIABLES_DIR, fn), 'r', encoding='utf-8-sig') as f:
                    text = self.strip_comments(f.read())
            except Exception:
                continue
            for m in re.finditer(r'@([A-Za-z0-9_]+)[ \t\r\n]*=[ \t\r\n]*(-?[\d.]+)', text):
                try:
                    self.scripted_variables[m.group(1)] = float(m.group(2))
                except ValueError:
                    pass

    def load_inline_scripts(self):
        if not os.path.isdir(INLINE_SCRIPTS_DIR):
            return
        for root, _, files in os.walk(INLINE_SCRIPTS_DIR):
            for fn in files:
                if not fn.endswith('.txt'):
                    continue
                path = os.path.join(root, fn)
                rel = os.path.relpath(path, INLINE_SCRIPTS_DIR).replace('\\', '/')
                key = rel[:-4] if rel.endswith('.txt') else rel
                try:
                    with open(path, 'r', encoding='utf-8-sig') as f:
                        self.inline_cache[key] = self.strip_comments(f.read())
                except Exception:
                    pass

    def expand_inline_scripts(self, text, depth=0):
        """展开 inline_script = { script = ... }，用括号配对定位，避免嵌套截断。"""
        if depth > 8:
            return text
        out = []
        i = 0
        n = len(text)
        changed = False
        while i < n:
            c = text[i]
            if c == '"':
                j = self.skip_string(text, i)
                out.append(text[i:j])
                i = j
                continue
            m = INLINE_HEAD.match(text, i)
            if m:
                span = self.find_block_span(text, m.start())
                full = text[span[0]:span[1]] if span else None
                if full:
                    body = self.inner(full)
                    sm = re.search(r'(?<![A-Za-z0-9_])script[ \t\r\n]*=[ \t\r\n]*"?([^\s{}"]+)"?', body)
                    content = None
                    if sm:
                        key = sm.group(1)
                        content = self.inline_cache.get(key) or self.inline_cache.get(os.path.basename(key))
                    if content is None:
                        out.append(full)
                    else:
                        params = {}
                        for pm in re.finditer(
                            r'([A-Za-z0-9_]+)[ \t\r\n]*=[ \t\r\n]*(?:"([^"]*)"|([^\s{}]+))', body
                        ):
                            k = pm.group(1)
                            if k == 'script':
                                continue
                            params[k] = pm.group(2) if pm.group(2) is not None else pm.group(3)
                        expanded = content
                        for k, v in params.items():
                            expanded = expanded.replace(f'${k}$', v).replace(f'@{k}@', v)
                        out.append(expanded)
                        changed = True
                    i = span[1]
                    continue
            out.append(c)
            i += 1
        result = ''.join(out)
        return self.expand_inline_scripts(result, depth + 1) if changed else result

    # ---------- 本地化 ----------

    @staticmethod
    def strip_inline_comment(value):
        out = []
        in_string = False
        i = 0
        n = len(value)
        while i < n:
            c = value[i]
            if c == '\\' and i + 1 < n:
                out.append(value[i:i + 2])
                i += 2
                continue
            if c == '"':
                in_string = not in_string
                out.append(c)
                i += 1
                continue
            if c == '#' and not in_string:
                break
            out.append(c)
            i += 1
        return ''.join(out).strip()

    def load_localisation(self):
        if not os.path.isdir(LOCALISATION_DIR):
            return
        for root, _, files in os.walk(LOCALISATION_DIR):
            for fn in files:
                if not fn.endswith('.yml'):
                    continue
                try:
                    with open(os.path.join(root, fn), 'r', encoding='utf-8-sig') as f:
                        for line in f:
                            line = line.strip()
                            if not line or line.startswith('#') or ':' not in line:
                                continue
                            key, value = line.split(':', 1)
                            key = key.strip()
                            value = value.strip()
                            if not key:
                                continue
                            # YAML 语言声明行（l_simp_chinese: 之类）
                            if re.fullmatch(r'l_[A-Za-z0-9_]+', key) and not value:
                                continue
                            # 版本号前缀：KEY:0 "文本" / KEY:1 "文本"
                            vm = re.match(r'^\d+[ \t]*:[ \t]*(.*)$', value)
                            if vm:
                                value = vm.group(1).strip()
                            value = self.strip_inline_comment(value)
                            if len(value) >= 2 and value.startswith('"') and value.endswith('"'):
                                value = value[1:-1]
                            value = value.replace('\\n', ' ').replace('\\t', ' ').strip()
                            if value:
                                self.localisation[key] = value
                except Exception:
                    pass

    def resolve_localisation(self, text, depth=0):
        if not text or depth > 10:
            return text
        for match in re.findall(r'\$([^$]+)\$', text):
            # $t$ / $n$ 是文本占位符，不是本地化键。
            if match in ('t', 'T', 'n', 'N'):
                replacement = ' '
            else:
                replacement = self.localisation.get(match)
                if replacement is None:
                    continue
                replacement = self.resolve_localisation(replacement, depth + 1)
            text = text.replace(f'${match}$', replacement)
        # 形如 $@waystation_modifier_negative_mult_25_percent|+0%$ 的脚本变量引用
        def at_ref(match):
            inner = match.group(1)
            name = inner.split('|')[0].lstrip('@')
            value = self.scripted_variables.get(name)
            if value is None:
                return match.group(0)
            if name.endswith(('_mult', '_mult_add')) or 'percent' in name or 'pct' in name:
                return f'{value * 100:+g}%'
            return f'{value:+g}' if value >= 0 else f'{value:g}'
        text = re.sub(r'\$(@[^$]*)\$', at_ref, text)
        text = text.replace('\\n', ' ').replace('\\t', ' ')
        return re.sub(r'\s+', ' ', text).strip()

    def loc(self, key):
        if not key:
            return ''
        # 恒星基地模块/建筑的名称统一使用 sm_ 前缀（如 sm_waystation_energy_module）。
        candidates = (
            f'sm_{key}', key, f'NAME_{key}', f'{key}_NAME', key.upper(), f'NAME_{key.upper()}',
        )
        for candidate in candidates:
            if candidate in self.localisation:
                return self.resolve_localisation(self.localisation[candidate])
        return key


class StarbaseModuleExtractor:
    def __init__(self):
        self.parser = ClausewitzParser()
        self.records = []
        self.failures = []
        self.missing_modifiers = set()

    # ---------- 分类 ----------

    DEFINITION_MARKERS = ('icon', 'category', 'construction_days', 'section')

    def looks_like_definition(self, body):
        """真正的定义块必须带资源块或定义字段；用于排除 triggered_* 之类的片段。"""
        if re.search(r'(?<![A-Za-z0-9_])resources[ \t\r\n]*=[ \t\r\n]*\{', body):
            return True
        if self.parser.str_field(body, 'starbase_type'):
            return True
        return any(self.parser.has_key(body, k) for k in self.DEFINITION_MARKERS)

    @staticmethod
    def platform_from_source(filename):
        """脚本目录按平台分文件，文件名是最可靠的平台判据。"""
        name = filename.lower()
        if 'waystation' in name:
            return 'waystation'
        if 'arkship' in name:
            return 'arkship'
        if 'orbital_ring' in name:
            return 'orbital_ring'
        return None

    @staticmethod
    def flag_yes(text, token):
        """判定 `token = yes`；`token = no` 不算命中（例如 reloading_bay 的 is_waystation_starbase = no）。"""
        m = re.search(rf'(?<![A-Za-z0-9_]){token}\s*=\s*(yes|no)', text)
        return bool(m and m.group(1) == 'yes')

    def detect_platform(self, key, declared_type, potential, body, source_platform=None):
        if source_platform:
            return source_platform
        if declared_type:
            return declared_type
        if key.startswith('waystation_'):
            return 'waystation'
        if key.startswith('arkship_'):
            return 'arkship'
        normal = self.flag_yes(potential, 'is_normal_starbase')
        if not normal and self.flag_yes(potential, 'is_waystation_starbase'):
            return 'waystation'
        if not normal and self.flag_yes(potential, 'is_arkship_starbase'):
            return 'arkship'
        if not normal and self.flag_yes(potential, 'is_orbital_ring'):
            return 'orbital_ring'
        return 'starbase'

    def cross_platform_notes(self, platform, potential, possible):
        """记录“跨平台可建”的情况（例如兽栏既能建在恒星基地也能建在驿站）。"""
        notes = []
        if platform == 'starbase':
            if self.flag_yes(potential, 'is_waystation_starbase'):
                notes.append('同样可建于驿站基地')
            if self.flag_yes(potential, 'is_orbital_ring') or self.flag_yes(possible, 'is_orbital_ring'):
                notes.append('同样可建于星环基地')
        elif self.flag_yes(potential, 'is_normal_starbase'):
            notes.append('同样可建于恒星基地')
        return notes

    def classify_module(self, key, platform, potential, body):
        if platform == 'waystation':
            return '驿站模块'
        if platform == 'arkship':
            if self.parser.str_field(body, 'category') == 'sm_spinal_cat' or self.parser.has_key(body, 'fits_on_slots'):
                return '方舟脊柱模块'
            return '方舟模块'
        if platform == 'orbital_ring':
            return '星环模块'
        return '恒星基地模块'

    def classify_building(self, key, platform, potential, body):
        if platform == 'arkship':
            return '方舟建筑'
        if platform == 'waystation':
            return '驿站建筑'
        if platform == 'orbital_ring':
            return '星环建筑'
        return '恒星基地建筑'

    # ---------- 修正 / 资源 ----------

    def format_modifier(self, key, value):
        try:
            v = float(value)
        except ValueError:
            return f'{key}={value}'
        translation = MODIFIER_TRANSLATIONS.get(key)
        if not translation:
            # 官方本地化使用大写键（MOD_SHIP_HULL_MULT）或小写 mod_ 前缀。
            for candidate in (f'mod_{key}', f'MOD_{key.upper()}', f'MOD_{key}', f'mod_{key.upper()}'):
                raw = self.parser.localisation.get(candidate)
                if raw:
                    translation = self.parser.resolve_localisation(raw)
                    break
        if not translation:
            translation = key
            self.missing_modifiers.add(key)
        use_pct = False
        if key in FORCE_PERCENTAGE_KEYS:
            use_pct = True
        elif key not in NON_PERCENTAGE_KEYS and key.endswith(('_mult', '_perc', '_add_perc', '_reduction')):
            use_pct = True
        if use_pct:
            pv = v * 100
            shown = f"{int(pv)}%" if abs(pv - round(pv)) < 0.001 else f"{pv:g}%"
        else:
            shown = f"{int(v)}" if v == int(v) else f"{v:g}"
        return f"{translation}{'+' if v >= 0 else ''}{shown}"

    def parse_modifiers(self, body):
        entries = []
        seen = set()
        for kind, _full, inner in self.parser.extract_blocks(body):
            if kind not in MODIFIER_PRIORITY:
                continue
            inner = self.parser.resolve_at(inner)
            tooltip = ''
            raw_tooltip = self.parser.str_field(inner, 'custom_tooltip')
            if raw_tooltip:
                tooltip = self.parser.loc(raw_tooltip)
            show_only = self.parser.str_field(inner, 'show_only_custom_tooltip') == 'yes'
            for k, v in self.parser.scalar_pairs(inner):
                if k in NON_MODIFIER_KEYS:
                    continue
                item = (kind, k, v, tooltip)
                if item in seen:
                    continue
                seen.add(item)
                effect = self.format_modifier(k, v)
                prefix = ''
                if tooltip:
                    prefix = tooltip if show_only else f'【{tooltip}】'
                entries.append((MODIFIER_PRIORITY[kind], f'{prefix}{effect}'))
        entries.sort(key=lambda x: x[0])
        return [text for _, text in entries]

    def format_resource_list(self, entries):
        grouped = []
        for category, key, value in entries:
            if key == 'colony':
                grouped.append(f'殖民地资源（{value}）')
                continue
            try:
                v = float(value)
                shown = f"{int(v)}" if v.is_integer() else f"{v:g}"
            except (TypeError, ValueError):
                shown = str(value)
            text = f"{RESOURCE_CN.get(key, key)}{shown}"
            if category == 'triggered':
                text += '（条件）'
            grouped.append(text)
        return '、'.join(grouped)

    def parse_resources(self, body):
        cost, upkeep, produced = [], [], []
        collected = []
        for kind, _full, inner in self.parser.extract_blocks(body):
            if kind == 'resources':
                for sub_kind, _sf, sub in self.parser.extract_blocks(inner):
                    sub = self.parser.resolve_at(sub)
                    if sub_kind == 'cost':
                        for k, v in self.parser.scalar_pairs(sub):
                            cost.append(('normal', k, v))
                    elif sub_kind == 'upkeep':
                        tag = 'triggered' if re.search(r'(?<![A-Za-z0-9_])trigger\s*=', sub) else 'normal'
                        for k, v in self.parser.scalar_pairs(sub):
                            upkeep.append((tag, k, v))
                    elif sub_kind == 'produces':
                        tag = 'triggered' if re.search(r'(?<![A-Za-z0-9_])trigger\s*=', sub) else 'normal'
                        for k, v in self.parser.scalar_pairs(sub):
                            produced.append((tag, k, v))
            elif kind == 'collected_system_resources':
                for res, _sf, sub in self.parser.extract_blocks(inner):
                    number = re.search(r'(-?[\d.]+)', self.parser.resolve_at(sub))
                    collected.append(('normal', res, number.group(1) if number else '按条件'))
            elif kind == 'collected_colony_resources':
                ctype = self.parser.str_field(inner, 'collection_type')
                collected.append(('normal', 'colony', '偷取' if ctype == 'steal' else (ctype or '采集')))
        return {
            'cost': self.format_resource_list(cost),
            'upkeep': self.format_resource_list(upkeep),
            'produces': self.format_resource_list(produced),
            'collected': self.format_resource_list(collected),
        }

    # ---------- 主流程 ----------

    def parse_file(self, file_path, kind):
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                raw = f.read()
        except Exception as e:
            print(f'读取文件失败 {file_path}: {e}')
            return
        text = self.parser.strip_comments(raw)
        text = self.parser.expand_inline_scripts(text)
        text = self.parser.resolve_at(text)

        source_platform = self.platform_from_source(os.path.basename(file_path))
        for key, _full, body in self.parser.extract_blocks(text):
            if not body or not self.looks_like_definition(body):
                continue
            try:
                record = self.build_record(key, body, file_path, kind, source_platform)
            except Exception as e:  # 单条失败不阻断全局
                self.failures.append((os.path.basename(file_path), key, str(e)))
                continue
            if record:
                self.records.append(record)

    def build_record(self, key, body, file_path, kind, source_platform=None):
        declared_type = self.parser.str_field(body, 'starbase_type')
        potential = self.parser.direct_block(body, 'potential')
        possible = self.parser.direct_block(body, 'possible')
        platform = self.detect_platform(key, declared_type, potential, body, source_platform)

        category = self.parser.str_field(body, 'category')
        section = self.parser.str_field(body, 'section')
        tech_candidates = self.parser.str_list(potential, 'has_technology')
        tech_candidates += self.parser.str_list(body, 'show_in_tech')
        tech_candidates = list(dict.fromkeys(t for t in tech_candidates if t))
        tech_name = '、'.join(dict.fromkeys(self.parser.loc(t) for t in tech_candidates))

        if kind == 'module':
            record_type = self.classify_module(key, platform, potential, body)
        else:
            record_type = self.classify_building(key, platform, potential, body)

        resources = self.parse_resources(body)
        notes = []
        tooltip_key = self.parser.str_field(body, 'custom_tooltip_with_modifiers') or \
            self.parser.str_field(body, 'custom_tooltip')
        if tooltip_key:
            notes.append(self.parser.loc(tooltip_key))
        if self.parser.str_field(body, 'remote_network_resource_collection') == 'yes':
            notes.append('启用驿站网络远程采集')
        notes.extend(self.cross_platform_notes(platform, potential, possible))

        starbase_limit = self.parser.str_field(self.parser.direct_block(body, 'starbase_limit'), 'base')
        empire_limit_block = self.parser.direct_block(body, 'empire_limit')
        empire_limit = self.parser.str_field(empire_limit_block, 'base')
        limit_text = '不限' if starbase_limit == '-1' else starbase_limit

        if kind == 'module':
            generator = (f'effect starbase = {{ set_starbase_module = '
                         f'{{ slot = <槽位序号> module = {key} }} }}')
        else:
            generator = (f'effect starbase = {{ set_starbase_building = '
                         f'{{ slot = <槽位序号> building = {key} }} }}')

        return {
            'record_type': record_type,
            'base_type': platform,
            'key': key,
            'name': self.parser.loc(key),
            'category': CATEGORY_CN.get(category, category),
            'section': section,
            'icon': self.parser.str_field(body, 'icon'),
            'tech_key': tech_candidates[0] if tech_candidates else '',
            'tech_name': tech_name,
            'potential': potential,
            'possible': possible,
            'cost': resources['cost'],
            'upkeep': resources['upkeep'],
            'produces': resources['produces'],
            'collected': resources['collected'],
            'effects': self.parse_modifiers(body),
            'equipped': '、'.join(self.parser.str_list(body, 'equipped_component')),
            'component_set': '、'.join(
                self.parser.str_list(body, 'component_set')
                + self.parser.str_list(body, 'triggered_component_set')
            ),
            'starbase_limit': limit_text,
            'empire_limit': empire_limit,
            'initial': '是' if self.parser.str_field(body, 'initial') == 'yes' else '否',
            'generator': generator,
            'notes': '；'.join(notes),
            'source': os.path.basename(file_path),
        }

    def deduplicate(self):
        seen = set()
        unique = []
        for rec in self.records:
            dedupe = (rec['record_type'], rec['key'])
            if dedupe in seen:
                continue
            seen.add(dedupe)
            unique.append(rec)
        self.records = unique

    def run(self):
        print('加载脚本变量...')
        self.parser.load_scripted_variables()
        print('加载内联脚本...')
        self.parser.load_inline_scripts()
        print('加载本地化...')
        self.parser.load_localisation()
        print('解析恒星基地、星环、方舟与驿站脚本...')
        for directory, kind in ((STARBASE_MODULE_DIR, 'module'), (STARBASE_BUILDING_DIR, 'building')):
            for root, _, files in os.walk(directory):
                for fn in sorted(files):
                    if fn.endswith('.txt'):
                        self.parse_file(os.path.join(root, fn), kind)
        self.deduplicate()
        if self.failures:
            print(f'解析失败 {len(self.failures)} 条：')
            for src, key, err in self.failures[:20]:
                print(f'  {src} :: {key} :: {err}')

    # ---------- 输出 ----------

    @staticmethod
    def ensure_writable_output_dir():
        candidates = [OUTPUT_DIR, os.path.dirname(os.path.dirname(os.path.abspath(__file__)))]
        for candidate in candidates:
            if not candidate or not os.path.isdir(candidate):
                continue
            probe = os.path.join(candidate, '__dsh_write_probe.txt')
            try:
                with open(probe, 'w', encoding='utf-8') as f:
                    f.write('ok')
                os.unlink(probe)
                return candidate
            except Exception:
                continue
        return candidates[-1]

    @staticmethod
    def row_values(rec):
        return [
            rec['record_type'], rec['name'], rec['key'],
            PLATFORM_TYPE_CN.get(rec['base_type'], rec['base_type']),
            rec['category'], rec['section'] or '-',
            rec['tech_name'] or '-', rec['tech_key'] or '-',
            rec['cost'] or '-', rec['upkeep'] or '-',
            rec['produces'] or '-', rec['collected'] or '-',
            '；'.join(rec['effects']) if rec['effects'] else '-',
            rec['equipped'] or '-', rec['component_set'] or '-',
            rec['starbase_limit'] or '-', rec['initial'],
            rec['generator'],
            rec['notes'] or '-', rec['source'],
        ]

    HEADERS = [
        '类型', '名称', '代码', '适用平台', '分类', '槽位/区块', '解锁科技', '科技代码',
        '花费', '维护', '产出', '采集', '效果', '装备组件', '组件集',
        '数量上限', '初始', '生成代码', '备注', '来源文件',
    ]
    SHEET_ORDER = [
        '恒星基地模块', '星环模块', '方舟模块', '方舟脊柱模块', '驿站模块',
        '恒星基地建筑', '星环建筑', '方舟建筑', '驿站建筑',
    ]

    def save_txt(self, path):
        counts = {}
        for rec in self.records:
            counts[rec['record_type']] = counts.get(rec['record_type'], 0) + 1
        lines = ['恒星基地/星环/方舟/驿站提取结果']
        lines.append('统计：' + '，'.join(f'{k} {v}' for k, v in sorted(counts.items())))
        lines.append('')
        for rec in self.records:
            values = self.row_values(rec)
            for header, value in zip(self.HEADERS, values):
                lines.append(f'{header}：{value}')
            lines.append('')
        with open(path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    def save_json(self, path):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.records, f, ensure_ascii=False, indent=2)

    def save_csv(self, path):
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(self.HEADERS)
            for rec in self.records:
                writer.writerow(self.row_values(rec))

    def save_excel(self, path):
        wb = xlsxwriter.Workbook(path)
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
                                    'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        wrap = wb.add_format({'border': 1, 'valign': 'top', 'text_wrap': True})
        center = wb.add_format({'border': 1, 'valign': 'vcenter', 'align': 'center'})
        widths = [16, 22, 34, 12, 12, 24, 20, 30, 24, 24, 30, 34, 60, 30, 32, 10, 8, 62, 46, 34]
        for title in ['全部提取'] + self.SHEET_ORDER:
            rows = self.records if title == '全部提取' else \
                [r for r in self.records if r['record_type'] == title]
            if not rows:
                continue  # 空分类不建表（例如驿站没有独立建筑）
            ws = wb.add_worksheet(title)
            ws.freeze_panes(1, 0)
            for i, w in enumerate(widths):
                ws.set_column(i, i, w)
            for c, h in enumerate(self.HEADERS):
                ws.write(0, c, h, header_fmt)
            for r_i, rec in enumerate(rows, 1):
                for c, value in enumerate(self.row_values(rec)):
                    fmt = center if c in (0, 3, 4, 15, 16) else wrap
                    ws.write(r_i, c, value, fmt)
            ws.autofilter(0, 0, max(len(rows), 1), len(self.HEADERS) - 1)
        wb.close()


def main():
    output_dir = StarbaseModuleExtractor.ensure_writable_output_dir()
    extractor = StarbaseModuleExtractor()
    extractor.run()
    base = os.path.join(output_dir, f'{VERSION}恒星基地模块建筑')
    extractor.save_txt(base + '.txt')
    extractor.save_json(base + '.json')
    extractor.save_csv(base + '.csv')
    extractor.save_excel(base + '.xlsx')
    print(f'已提取 {len(extractor.records)} 条记录')
    counts = {}
    for rec in extractor.records:
        counts[rec['record_type']] = counts.get(rec['record_type'], 0) + 1
    for k, v in sorted(counts.items()):
        print(f'  {k}: {v}')
    if extractor.missing_modifiers:
        debug_path = base + '_缺失修正翻译.txt'
        with open(debug_path, 'w', encoding='utf-8') as f:
            f.write('以下修正没有中文翻译（可补充到 代码项目/修正翻译字典.py 的 MODIFIER_TRANSLATIONS）：\n')
            for k in sorted(extractor.missing_modifiers):
                f.write(f'  "{k}": "",\n')
        print(f'未翻译修正 {len(extractor.missing_modifiers)} 个 -> {os.path.basename(debug_path)}')
    print(f'输出目录: {output_dir}')


if __name__ == '__main__':
    main()
