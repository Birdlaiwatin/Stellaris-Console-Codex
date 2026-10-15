# 规则外不转换百分比字典
NON_PERCENTAGE_KEYS = [
    'on_fauna_capture_scientist_exp_gain'
]
# 强制不按百分比显示的舰船修正
NON_PERCENTAGE_KEYS += [
    "ship_weapon_slots_add",
    "ship_weapon_slots",
    "ship_utility_slots_add",
    "ship_utility_slots",
    "ship_aux_slots_add",
    "ship_aux_slots",
    "ship_max_hull_add",
    "ship_hull_add",
    "ship_armor_add",
    "ship_shield_add",
    "ship_max_shield_add",
    "ship_hull_regen_add",
    "ship_shield_regen_add",
    "ship_armor_regen_add",
    "ship_damage_add",
    "ship_weapon_damage_add",
]
# 规则外转换百分比字典
FORCE_PERCENTAGE_KEYS = [
    'ship_capture_chance_add',
    'pop_government_ethic_attraction',
    'ship_armor_hardening_add',
    'ship_shield_hardening_add',
    'ship_damage_reduction_add',
    'ship_weapon_damage',
    'army_morale',
    'create_debris_chance',
    'planet_orbital_bombardment_damage',
    'pop_environment_tolerance',
    'army_health',
    'faction_approval',
    'shroud_delve_cooldown',
    'habitability_ceil_add',
    'ship_armor_penetration_add',
    'ship_shield_penetration_add',
    'country_trade_fee',
    'habitability_floor_add',
    'habitability_ceil_add',
]

