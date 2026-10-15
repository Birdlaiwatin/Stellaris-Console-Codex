# -*- coding: utf-8 -*-
"""
星海共同体议案提取（群星 4.5.2）

数据来源：
    common\\resolution_categories\\*.txt   议案类别定义（组别 group、类别、议案顺序）
    common\\resolutions\\*.txt             议案定义（modifier 修正、target、effect）
    common\\static_modifiers\\*.txt        议案修正引用的静态修正
    localisation\\simp_chinese\\*.yml      组别/类别/议案名称与效果文本

输出：
    <OUTPUT_DIR>\\<VERSION>星海共同体议案.xlsx

列结构（与旧版 4.3.4 表格保持一致）：
    组别 | 类别 | 议案名称 | 立即通过 | 立即取消 | 效果 | 修正 | 目标需求 | 议案key

其中"效果"列优先使用人工维护的 议案效果人工文本.py，缺失时回落到游戏本地化文本。
"""

import os
import re
import sys
import glob
import json
import argparse
import collections

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from A配置路径 import BASE_GAME_DIR, OUTPUT_DIR, LOCALISATION_DIR, VERSION

try:
    from 议案效果人工文本 import MANUAL_EFFECTS
except ImportError:
    MANUAL_EFFECTS = {}

# ---------------------------------------------------------------- 路径配置
RESOLUTION_DIR = os.path.join(BASE_GAME_DIR, r'common\resolutions')
CATEGORY_DIR = os.path.join(BASE_GAME_DIR, r'common\resolution_categories')
STATIC_MODIFIER_DIR = os.path.join(BASE_GAME_DIR, r'common\static_modifiers')
SCRIPTED_LOC_DIR = os.path.join(BASE_GAME_DIR, r'common\scripted_loc')

# 组别展示顺序（沿用旧表顺序）。未列出的新组别排在最后。
GROUP_ORDER = (
    'commerce_industry',
    'politics_culture',
    'environment_technology',
    'defense_war',
    'suzerains_sovereignty',
    'galactic_focus',
    'galactic_empire',
    'custodianship',
)

# 组别底色（沿用旧表配色）
GROUP_FILLS = {
    '贸易与产业': 'FFF3CA',
    '政治与文化': 'E7D5FF',
    '环境与科技': 'C3FCFF',
    '防御与战争': 'FDD4D4',
    '宗主与主权': 'E3CDFF',
    '星系优先事项': 'DFF8C1',
    '星海帝国': 'FFB7C0',
    '监管人任期': 'FFF6C3',
}
DEFAULT_FILL = 'F2F2F2'

# 不参与展示的内部定义（该块仅用于代码支持）
SKIP_KEYS = {'any'}

# 修正键中作为修饰而非数值的字段
MODIFIER_META_KEYS = {
    'custom_tooltip', 'show_only_custom_tooltip', 'icon', 'icon_frame',
    'hide_from_country_list', 'apply_modifier_to_other_planets', 'important',
    '__items__', 'desc',
}

# 按"数值+百分比"显示的修正后缀；其余按原值显示
PERCENT_SUFFIXES = ('_mult', '_mult_add', '_perc', '_gain', '_happiness', '_speed',
                    '_habitability', '_reduction', '_uninhabitable', '_tolerance',
                    '_political_power', '_difficulty_add')

# 游戏脚本使用、但本地化文件未提供显示名的修正键 -> 中文显示名
MODIFIER_NAME_OVERRIDES = {
    'monthly_loyalty': '每月忠诚度变化',
    'all_technology_research_speed': '所有科技研究速度',
    'gdf_ship_alloys_cost_mult': '泛星系防卫军舰船合金花费',
    'country_community_gdf_naval_capacity': '共同体舰队容量',
    'senate_vote_council_member_speed_mult': '理事会成员议案表决速度',
    'senate_recess_council_member_speed_mult': '理事会成员议案休会速度',
}

# 即使不带常见百分比后缀也应按百分比显示的修正键
ALWAYS_PERCENT_KEYS = {
    'pop_environment_tolerance',
    'country_trade_fee',
    'diplomacy_upkeep_mult',
    'country_naval_cap_mult',
}

# 图标/概念类占位符（富文本字段）处理
RE_COLOR = re.compile(r'§.')
RE_ICON = re.compile(r'£[^£]*£')
RE_DOLLAR = re.compile(r'\$([^$]+)\$')
RE_BRACKET = re.compile(r'\[([^\]]+)\]')
RE_LOC_LINE = re.compile(r'\s*([A-Za-z0-9_.\-]+)\s*:\s*(?:\d+\s*)?"(.*)"\s*$')


# ================================================================ 基础工具
def strip_comment(line):
    """去掉行尾 # 注释（不处理引号内的 #）。"""
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif ch == '#' and not in_str:
            break
        else:
            out.append(ch)
    return ''.join(out)


TOKEN_RE = re.compile(r'"(?:[^"\\]|\\.)*"|\{|\}|=|[^\s{}=]+')


def tokenize(text):
    toks = []
    for raw in text.split('\n'):
        line = strip_comment(raw)
        if line.strip():
            toks.extend(m.group(0) for m in TOKEN_RE.finditer(line))
    return toks


def _unquote(tok):
    if len(tok) >= 2 and tok[0] == '"' and tok[-1] == '"':
        return tok[1:-1]
    return tok


class _Parser:
    """极简 Clausewitz 解析器：块 -> dict，纯列表块 -> list。"""

    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else None

    def next(self):
        v = self.t[self.i]
        self.i += 1
        return v

    def at_end(self):
        return self.i >= len(self.t)

    def parse_block(self):
        d, lst = {}, []
        while not self.at_end():
            tok = self.peek()
            if tok == '}':
                self.next()
                break
            if tok == '{':
                self.next()
                lst.append(self.parse_block())
                continue
            if self.peek(1) == '=':
                key = _unquote(self.next())
                self.next()  # '='
                val = self.parse_value()
                if key in d:
                    prev = d[key]
                    d[key] = prev + [val] if isinstance(prev, list) else [prev, val]
                else:
                    d[key] = val
            else:
                lst.append(_unquote(self.next()))
        if not d:
            return lst
        if lst:
            d.setdefault('__items__', []).extend(lst)
        return d

    def parse_value(self):
        tok = self.next()
        if tok == '{':
            return self.parse_block()
        return _unquote(tok)


def parse_script(text):
    p = _Parser(tokenize(text))
    if p.at_end():
        return {}
    if p.peek() == '{':
        p.next()
    return p.parse_block()


def read_script(path):
    with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
        return parse_script(f.read())


def _iter_blocks(parsed, key):
    """遍历顶层 key 下的所有块；同名重复键会被解析成列表，这里逐个展开。"""
    if not isinstance(parsed, dict):
        return
    block = parsed.get(key)
    if isinstance(block, list):
        for item in block:
            if isinstance(item, dict):
                yield key, item
    elif isinstance(block, dict):
        yield key, block


def _first(value):
    """同名重复键可能解析成列表；取第一个值（name 等字段实际只应有一个）。"""
    if isinstance(value, list):
        return value[0] if value else None
    return value


# ================================================================ 本地化
class Localization:
    def __init__(self, loc_dir):
        self.raw = {}
        self.scripted_loc = {}
        self._cache = {}
        self._load_yml(loc_dir)
        self._load_scripted_loc()

    def _load_yml(self, loc_dir):
        for path in glob.glob(os.path.join(loc_dir, '**', '*.yml'), recursive=True):
            try:
                with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
                    for line in f:
                        m = RE_LOC_LINE.match(line.rstrip('\n'))
                        if m:
                            # 同名键以先出现的为准（后加载的 DLC 覆盖文件不反向覆盖）
                            self.raw.setdefault(m.group(1), m.group(2))
            except OSError:
                continue

    def _load_scripted_loc(self):
        """defined_text：name -> default 本地化键，用于解析 [GetWorker] 之类替换符。"""
        if not os.path.isdir(SCRIPTED_LOC_DIR):
            return
        for path in glob.glob(os.path.join(SCRIPTED_LOC_DIR, '*.txt')):
            parsed = read_script(path)
            for _key, block in _iter_blocks(parsed, 'defined_text'):
                if not isinstance(block, dict):
                    continue
                name = _first(block.get('name'))
                default = _first(block.get('default'))
                if isinstance(name, str) and isinstance(default, str):
                    self.scripted_loc.setdefault(name, default)

    def get(self, key):
        return self.raw.get(key)

    def expand(self, text, depth=0):
        """递归展开 $key$ 替换符。"""
        if not isinstance(text, str) or depth > 8 or '$' not in text:
            return text
        if text in self._cache:
            return self._cache[text]

        def repl(m):
            inner = m.group(1)
            val = self.raw.get(inner)
            if val is None:
                return m.group(0)
            return self.expand(val, depth + 1)

        out = RE_DOLLAR.sub(repl, text)
        self._cache[text] = out
        return out

    # [xxx.GetNamePlural] 之类的岗位/概念引用：概念键 -> 岗位键前缀
    JOB_KEY_PREFIXES = {
        'bureaucrat': 'job_bureaucrat',
        'foundry': 'job_foundry',
        'metallurgist': 'job_metallurgist',
        'trader': 'job_trader',
        'worker': 'pop_cat_worker',
    }

    def _bracket_text(self, inner):
        """解析单个 [xxx] 替换符；解析不到返回 None。"""
        if inner.startswith("'") and inner.endswith("'"):
            inner = inner[1:-1]
        base = inner.split('.')[0]
        suffix = inner.split('.')[1] if '.' in inner else ''

        if suffix == 'GetIcon':
            return ''

        # 1) 直接命中本地化键
        for cand in (base, self.scripted_loc.get(base)):
            if cand and cand in self.raw:
                return self.expand(self.raw[cand])

        # 2) [bureaucrat.GetNamePlural] -> job_bureaucrat_plural / pop_cat_bureaucrat
        if suffix.startswith('GetName'):
            plural = suffix != 'GetName'
            prefix = self.JOB_KEY_PREFIXES.get(base.lower(), 'job_' + base)
            candidates = []
            if plural:
                candidates += [prefix + '_plural', prefix + 's']
            candidates.append(prefix)
            if plural:
                candidates.append(prefix.rstrip('s') + 's')
            for cand in candidates:
                if cand in self.raw:
                    return self.expand(self.raw[cand])

        # 3) ['concept_xxx'] / ['concept_xxx:yyy']
        if ':' in inner:
            tail = inner.split(':')[-1]
            if tail in self.raw:
                return self.expand(self.raw[tail])
        return None

    def resolve_brackets(self, text, depth=0):
        """把 [GetWorker] 之类的脚本本地化替换符换成本地化文本。"""
        if not isinstance(text, str) or '[' not in text or depth > 6:
            return text

        def repl(m):
            result = self._bracket_text(m.group(1).strip())
            if result is None:
                return ''
            return self.resolve_brackets(result, depth + 1)

        return RE_BRACKET.sub(repl, text)

    def clean(self, text, keep_icons=False):
        """生成单元格文本：去色标/图标/未解析替换符，规整空白。

        注意顺序：必须先展开 $key$ 与 [脚本本地化]，再清除 £图标£/遗留占位符。
        原因是 $trade_with_icon$ 这类键本身就是本地化键，其内容形如 "£trade£ $trade$"，
        先删图标会把键名一起吃掉。
        """
        if text is None:
            return ''
        text = text.replace('\\n', '\n')
        text = RE_COLOR.sub('', text)
        text = self.expand(text)
        text = self.resolve_brackets(text)
        text = RE_COLOR.sub('', text)
        if not keep_icons:
            text = RE_ICON.sub('', text)
            text = re.sub(r'£[^£\s]*', '', text)
            text = RE_DOLLAR.sub('', text)
        text = RE_BRACKET.sub('', text)
        text = text.replace('\xa0', ' ')
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r' *\n *', '\n', text)
        text = re.sub(r'\n{2,}', '\n', text)
        return text.strip()

    def plain(self, key):
        """按键取本地化文本并清洗；$key$ 引用会先展开再清洗。"""
        raw = self.raw.get(key)
        if raw is None:
            return None
        return self.clean(raw)


# ================================================================ 数据加载
def load_categories():
    """按文件顺序返回 OrderedDict：类别key -> {group, resolution_types, ...}"""
    cats = collections.OrderedDict()
    for path in sorted(glob.glob(os.path.join(CATEGORY_DIR, '*.txt'))):
        for key, val in read_script(path).items():
            if isinstance(val, dict) and 'resolution_types' in val:
                cats[key] = val
    return cats


def load_resolutions():
    res = {}
    for path in sorted(glob.glob(os.path.join(RESOLUTION_DIR, '*.txt'))):
        for key, val in read_script(path).items():
            if isinstance(val, dict) and 'resources' in val:
                res[key] = val
    return res


def load_static_modifiers():
    sm = {}
    for path in sorted(glob.glob(os.path.join(STATIC_MODIFIER_DIR, '*.txt'))):
        for key, val in read_script(path).items():
            if isinstance(val, dict):
                sm[key] = val
    return sm


def first_custom_tooltip(block):
    """按出现顺序取块内第一个 custom_tooltip 字符串。"""
    if isinstance(block, list):
        for item in block:
            found = first_custom_tooltip(item)
            if found:
                return found
        return None
    if not isinstance(block, dict):
        return None
    for key, val in block.items():
        if key == 'custom_tooltip' and isinstance(val, str):
            return val
        if isinstance(val, (dict, list)):
            found = first_custom_tooltip(val)
            if found:
                return found
    return None


# 仅用于议案流程记账的临时修正，不属于议案对成员国生效的效果
BOOKKEEPING_MODIFIERS = {
    'resolution_passed_diplomatic_weight',
    'resolution_failed_diplomatic_weight',
}


def collect_added_modifiers(block, static_modifiers, depth=0):
    """收集块内 add_modifier 引用的静态修正。

    谴责类议案不写 modifier 块，而是在 effect 里 add_modifier = { modifier = xxx }，
    这里递归找出这些引用，使"修正"列不至于为空；议案流程记账用的临时修正会被忽略。
    """
    out = collections.OrderedDict()
    if depth > 6:
        return out
    if isinstance(block, list):
        for item in block:
            for key, val in collect_added_modifiers(item, static_modifiers, depth + 1).items():
                out[key] = val
        return out
    if not isinstance(block, dict):
        return out
    for key, val in block.items():
        if key == 'add_modifier':
            entries = val if isinstance(val, list) else [val]
            for entry in entries:
                name = entry.get('modifier') if isinstance(entry, dict) else entry
                if isinstance(name, str) and name in static_modifiers \
                        and name not in BOOKKEEPING_MODIFIERS:
                    for mod_key, mod_val in flatten_modifiers(static_modifiers[name],
                                                              static_modifiers).items():
                        out[mod_key] = mod_val
        elif isinstance(val, (dict, list)):
            for mod_key, mod_val in collect_added_modifiers(val, static_modifiers,
                                                            depth + 1).items():
                out[mod_key] = mod_val
    return out


def flatten_modifiers(block, static_modifiers, depth=0):
    """把 modifier 块摊平成 {修正键: 数值}。

    块可能是单个 dict，也可能是同名重复键解析出的 dict 列表（先合并再解析）。
    块内的取值有两种含义：
      * 键本身就是修正键（如 country_naval_cap_mult = -0.1）-> 直接取数值；
      * 键是静态修正名（如 resolution_sanctions_economic = 2）-> 展开该静态修正，
        并把其中的数值乘以这里的系数（制裁议案用 1/2/3 表示强度等级）。
    """
    out = collections.OrderedDict()
    if depth > 6 or block is None:
        return out
    blocks = block if isinstance(block, list) else [block]
    for item in blocks:
        if not isinstance(item, dict):
            continue
        for key, val in item.items():
            if key in MODIFIER_META_KEYS or isinstance(val, (dict, list)):
                continue
            if key in static_modifiers:
                try:
                    mult = float(val)
                except (TypeError, ValueError):
                    mult = 1.0
                for sub_key, sub_val in flatten_modifiers(static_modifiers[key],
                                                          static_modifiers, depth + 1).items():
                    out[sub_key] = out.get(sub_key, 0.0) + sub_val * mult
                continue
            try:
                out[key] = out.get(key, 0.0) + float(val)
            except (TypeError, ValueError):
                continue
    return out


# ================================================================ 生成器
class ResolutionExtractor:
    def __init__(self):
        self.loc = Localization(LOCALISATION_DIR)
        self.categories = load_categories()
        self.resolutions = load_resolutions()
        self.static_modifiers = load_static_modifiers()
        self.group_order = collections.OrderedDict(
            (g, i) for i, g in enumerate(GROUP_ORDER))
        self.debug = collections.defaultdict(set)

    # -------------------------------------------------- 本地化名称
    def group_name(self, group_key):
        """组别名：group 值本身即本地化键（commerce_industry -> 贸易与产业）。"""
        name = self.loc.plain(group_key)
        if name is None:
            self.debug['missing_group_name'].add(group_key)
            return group_key
        return name

    def category_name(self, cat_key):
        name = self.loc.plain(cat_key)
        if name is None:
            self.debug['missing_category_name'].add(cat_key)
            return cat_key
        return name

    def resolution_name(self, key):
        name = self.loc.plain(key)
        if name is None:
            self.debug['missing_resolution_name'].add(key)
            return key
        return name

    # -------------------------------------------------- 效果文本
    def effect_text(self, key):
        """优先级：人工文本 > _effect_success > _tooltip > effect 内 custom_tooltip。"""
        if key in MANUAL_EFFECTS and MANUAL_EFFECTS[key]:
            return MANUAL_EFFECTS[key]
        for candidate in (key + '_effect_success', key + '_tooltip'):
            if candidate in self.loc.raw:
                return self.loc.clean(self.loc.raw[candidate])
        ct = first_custom_tooltip(self.resolutions.get(key, {}).get('effect'))
        if ct and ct in self.loc.raw:
            return self.loc.clean(self.loc.raw[ct])
        return ''

    # -------------------------------------------------- 修正文本
    def modifier_name(self, mod_key):
        """修正显示名：本地化 MOD_大写 / mod_小写 > 内置补充 > 原键。"""
        for candidate in ('MOD_' + mod_key.upper(), 'mod_' + mod_key.lower(), mod_key):
            if candidate in self.loc.raw:
                name = self.loc.plain(candidate)
                # 图标被清除后残留的空格（"岗位产出的 贸易"）以及多余助词统一清理
                name = re.sub(r'\s+(?=[\u4e00-\u9fff])', '', name)
                name = re.sub(r'的(维护费|维护|upkeep)$', r'\1', name).strip()
                return name
        if mod_key in MODIFIER_NAME_OVERRIDES:
            return MODIFIER_NAME_OVERRIDES[mod_key]
        self.debug['missing_modifier_name'].add(mod_key)
        return mod_key

    def modifier_text(self, mod_key, value):
        pct = (mod_key in ALWAYS_PERCENT_KEYS
               or (mod_key.endswith(PERCENT_SUFFIXES) and not mod_key.endswith('_add')))
        text = f'{value * 100:g}%' if pct else f'{value:g}'
        sign = '+' if value >= 0 else ''
        return f'{self.modifier_name(mod_key)}{sign}{text}'

    def modifier_summary(self, key):
        """汇总议案生效的修正文本。

        同一个 modifier 块内的修正拼成一行；triggered_modifier 的每个分支各自成行
        （例如发展援助分"理事会成员"与"其他成员"两档，共同体先锋分四档）。
        """
        data = self.resolutions.get(key, {})
        lines = []

        base = flatten_modifiers(data.get('modifier'), self.static_modifiers)
        if base:
            lines.append(' '.join(self.modifier_text(k, v) for k, v in base.items()))

        triggered = data.get('triggered_modifier')
        entries = triggered if isinstance(triggered, list) else ([triggered] if triggered else [])
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            branch = flatten_modifiers(entry.get('modifier'), self.static_modifiers)
            if branch:
                lines.append(' '.join(self.modifier_text(k, v) for k, v in branch.items()))

        if not lines:
            # 谴责类议案把修正写在 effect 的 add_modifier 里
            added = collect_added_modifiers(data.get('effect'), self.static_modifiers)
            if added:
                lines.append(' '.join(self.modifier_text(k, v) for k, v in added.items()))

        return '\n'.join(lines)

    # -------------------------------------------------- 控制台指令
    @staticmethod
    def pass_command(key, targeted):
        if targeted:
            return ('effect pass_targeted_resolution={\n'
                    f' resolution={key}\n'
                    ' target=this\n'
                    '}')
        return f'effect pass_resolution_no_cooldown={key}'

    @staticmethod
    def cancel_command(key):
        return f'effect cancel_resolution={key}'

    # -------------------------------------------------- 组装行
    def build_rows(self):
        rows = []
        ordered = sorted(self.categories.items(),
                         key=lambda kv: self.group_order.get(kv[1].get('group'), 999))
        for cat_key, cat in ordered:
            group_key = cat.get('group', '')
            group = self.group_name(group_key)
            category = self.category_name(cat_key)
            for key in cat['resolution_types']:
                if key in SKIP_KEYS:
                    continue
                data = self.resolutions.get(key)
                if data is None:
                    self.debug['missing_resolution_definition'].add(key)
                    data = {}
                targeted = str(data.get('target', 'no')).strip() == 'yes'
                rows.append({
                    'group_key': group_key,
                    'group': group,
                    'category': category,
                    'name': self.resolution_name(key),
                    'pass': self.pass_command(key, targeted),
                    'cancel': self.cancel_command(key),
                    'effect': self.effect_text(key),
                    'modifier': self.modifier_summary(key),
                    'target': '是' if targeted else '否',
                    'key': key,
                })
        return rows


# ================================================================ 写 Excel
def write_workbook(rows, output_path, tmpdir=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '议案'

    thin = Side(style='thin', color='FFB0B0B0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    title = f'{VERSION}星海共同体议案'
    ws.cell(1, 1, title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    tcell = ws.cell(1, 1)
    tcell.font = Font(name='黑体', size=26, bold=True, color='FF000000')
    tcell.fill = PatternFill('solid', fgColor='FFD9E1F4')
    tcell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 33

    headers = ['组别', '类别', '议案名称', '立即通过', '立即取消', '效果', '修正', '目标需求', '议案key']
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(2, col, text)
        cell.font = Font(name='黑体', size=12, bold=True, color='FFFFFFFF')
        cell.fill = PatternFill('solid', fgColor='FF3656A9')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[2].height = 15.75

    def fill_for(group):
        return PatternFill('solid', fgColor='FF' + GROUP_FILLS.get(group, DEFAULT_FILL))

    bold_font = Font(name='黑体', size=10, color='FF08090C')
    mono_font = Font(name='Calibri', size=10, color='FF08090C')
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    wrap_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    row = 3
    group_spans = []      # (起始行, 结束行, 组别名)
    category_spans = []   # (起始行, 结束行)
    for item in rows:
        fill = fill_for(item['group'])
        values = [item['group'], item['category'], item['name'], item['pass'], item['cancel'],
                  item['effect'], item['modifier'], item['target'], item['key']]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(row, col, val if val else None)
            cell.fill = fill
            cell.border = border
            if col in (1, 2, 8):
                cell.font = bold_font
                cell.alignment = center
            elif col in (4, 5, 9):
                cell.font = mono_font
                cell.alignment = left
            else:
                cell.font = bold_font
                cell.alignment = wrap_left

        # 记录合并区间
        if group_spans and group_spans[-1][2] == item['group']:
            group_spans[-1] = (group_spans[-1][0], row, item['group'])
        else:
            group_spans.append((row, row, item['group']))
        if category_spans and category_spans[-1][3] == (item['group'], item['category']):
            category_spans[-1] = (category_spans[-1][0], row, item['category'],
                                  (item['group'], item['category']))
        else:
            category_spans.append((row, row, item['category'], (item['group'], item['category'])))

        # 行高按需要换行的内容粗略估算
        est = 1
        for txt in (item['effect'], item['modifier'], item['name']):
            if txt:
                lines = txt.count('\n') + 1
                est = max(est, lines, int(len(txt) / 42) + 1)
        ws.row_dimensions[row].height = max(25.5, min(est * 16.5 + 10, 409.5))
        row += 1

    # 合并组别与类别单元格
    for start, end, _ in group_spans:
        if end > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
    for start, end, _, _ in category_spans:
        if end > start:
            ws.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)

    # 组别字号加大
    for start, _, _ in group_spans:
        ws.cell(start, 1).font = Font(name='黑体', size=20, bold=True, color='FF000000')

    widths = {'A': 14, 'B': 22, 'C': 30, 'D': 46, 'E': 46, 'F': 52, 'G': 46, 'H': 10, 'I': 40}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'C3'

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    # 先写临时文件再原子替换，避免直接覆盖时因目标文件被占用而失败。
    # 受限环境下临时目录可能不可写，可用 --tmpdir 指定到别处。
    tmp_dir = tmpdir or os.path.dirname(output_path)
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_path = os.path.join(tmp_dir, os.path.basename(output_path) + '.tmp')
    if os.path.exists(tmp_path):
        os.remove(tmp_path)
    wb.save(tmp_path)
    try:
        os.replace(tmp_path, output_path)
    except OSError:
        import shutil
        shutil.copyfile(tmp_path, output_path)
        os.remove(tmp_path)
    return output_path


# ================================================================ 主流程
def main():
    parser = argparse.ArgumentParser(description='提取群星星海共同体议案并生成 Excel 表格')
    parser.add_argument('-o', '--outdir', default=OUTPUT_DIR,
                        help=f'输出目录，默认使用 A配置路径.OUTPUT_DIR（{OUTPUT_DIR}）')
    parser.add_argument('--tmpdir', default=None,
                        help='临时文件目录（默认与输出同目录）；受限环境下可指定到可写目录')
    args = parser.parse_args()

    extractor = ResolutionExtractor()
    rows = extractor.build_rows()
    output_path = os.path.join(args.outdir, f'{VERSION}星海共同体议案.xlsx')
    write_workbook(rows, output_path, tmpdir=args.tmpdir)

    print(f'共提取议案 {len(rows)} 条')
    print(f'输出文件: {output_path}')
    if extractor.debug:
        print('--- 警告 ---')
        for kind, items in sorted(extractor.debug.items()):
            sample = list(items)[:10]
            print(f'{kind} ({len(items)}): {sample}')
    return output_path


if __name__ == '__main__':
    main()
