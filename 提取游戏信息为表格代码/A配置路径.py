"""群星文件路径配置"""
#只需修改以下三行路径即可，其余路径无需更改。
#群星根目录文件夹
BASE_GAME_DIR = r"D:\steam\steamapps\common\Stellaris"
#输出目录文件夹
OUTPUT_DIR = r"D:\0QX\导出目录"
#对应文件图片文件夹，需要自己手动查找将dds图片转换成png图片
BASE_IMAGE_DIR = r"D:\0QX\图片文件"

# 版本号信息 - 用于生成不同版本的文件名，使用时输入当前版本号，进行版本差异对比时只保留前一个版本的版本号。
VERSION = "4.5.2"




####################以下路径如无改动无需修改#####################

##################################
###本地化文件以及其他脚本替换符等配置###
##################################

#本地化文件路径
LOCALISATION_DIR = BASE_GAME_DIR + r"\localisation\simp_chinese"
#@xxx变量替换文件路径（部分变量存在于其他文件）
SCRIPTED_VARIABLES_DIR = BASE_GAME_DIR + r"\common\scripted_variables"
#[替换符]路径
SCRIPTED_LOC_DIR = BASE_GAME_DIR + r"\common\scripted_loc"
#内联脚本路径
INLINE_SCRIPTS_DIR = BASE_GAME_DIR + r"\common\inline_scripts"
#颜色路径（用于领袖特质图片提取）
COLORS_DIR = BASE_GAME_DIR + r"\common\named_colors"

##################################
##########各种游戏文件配置##########
##################################

#起源与国策路径
CIVIC_DIR = BASE_GAME_DIR + r"\common\governments\civics"
#内阁岗位路径
COUNCILORS_DIR = BASE_GAME_DIR + r"\common\governments\councilors"
#高级政体路径
AUTHORITIES_FILE = BASE_GAME_DIR + r"\common\governments\authorities\00_authorities.txt"
#传统路径
TRADITIONS_DIR = BASE_GAME_DIR + r"\common\traditions"
#飞升路径
ASCENSION_PERKS_DIR = BASE_GAME_DIR + r"\common\ascension_perks"
#特质路径（包含领袖特质和物种特质）
TRAITS_DIR = BASE_GAME_DIR + r"\common\traits"
#区划路径
DISTRICTS_DIR = BASE_GAME_DIR + r"\common\districts"
#区域特化路径
ZONES_DIR = BASE_GAME_DIR + r"\common\zones"
#建筑路径
BUILDINGS_DIR = BASE_GAME_DIR + r"\common\buildings"
#行星类型路径
PLANETS_DIR = BASE_GAME_DIR + r"\common\planet_classes"
#星球模型提取
PLANETS_MODELS_DIR = BASE_GAME_DIR + r"\gfx\models\planets"
#科技路径
TECHNOLOGY_DIR = BASE_GAME_DIR + r"\common\technology"
#异常路径
ANOMALIES_DIR = BASE_GAME_DIR + r"\common\anomalies"
#巨构路径（没做）

#事件路径（没做）

#考古遗址路径（半成品）
ARCHAEOLOGICAL_SITES_DIR = BASE_GAME_DIR + r"\common\archaeological_site_types"
#星界裂隙路径（没做）

#标本路径
SPECIMENS_DIR = BASE_GAME_DIR + r"\common\specimens"
#遗珍路径（没做）
RELICS_DIR = BASE_GAME_DIR + r"\common\relics"
#全局修正路径
STATIC_MODIFIER_DIR = BASE_GAME_DIR + r"\common\static_modifiers"
#行星矿藏路径
DEPOSITS_DIR = BASE_GAME_DIR + r"\common\deposits"
#预设星系路径
SOLAR_SYSTEM_INITIALIZERS_DIR = BASE_GAME_DIR + r"\common\solar_system_initializers"
#全局舰船路径
DESIGNS_DIR = BASE_GAME_DIR + r"\common\global_ship_designs"
#舰船类型路径
SHIPSIZE_DIR = BASE_GAME_DIR + r"\common\ship_sizes"
#陆军类型路径
ARMIES_DIR = BASE_GAME_DIR + r"\common\armies"
#环境物体路径
OBJECT_DIR = BASE_GAME_DIR + r"\common\ambient_objects"
#舰船组件路径
COMPONENT_TEMPLATES_DIR = BASE_GAME_DIR + r"\common\component_templates"


##################################
##########图片索引文件配置##########
##################################

#飞升图片索引目录
PERK_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\ascension_perks.gfx"
#传统图片索引目录
TRADITIONS_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\traditions.gfx"
#领袖特质图片索引目录
LEADER_TRAITS_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\icons\traits\leader_traits.gfx"
LEADER_TRAITS_2_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\icons\traits\traits_shared.gfx"
#区划特化图片索引目录
ZONES_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\icons.gfx"
#标本图片索引目录
SPECIMENS_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\specimens.gfx"
#遗珍图片索引目录
RELICS_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\relics_view.gfx"
#行星图片索引目录
PLANETS_GFX_FILE_PATH = BASE_GAME_DIR + r"\interface\planet.gfx"

##################################
##########图片文件默认位置##########
##################################

CIVIC_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\governments\civics"
#高级政体图标目录
AUTHORITIES_GFX_DIR = BASE_GAME_DIR + r"\gfx\interface\icons\governments\authorities"
ORIGIN_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\origins"
TRAITS_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\traits"
TECHNO_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\technologies"
STATIC_MODIFIER_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\technologies"
DEPOSITS_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\deposits"
BUILDINGS_GFX_FILE_PATH = BASE_GAME_DIR + r"\gfx\interface\icons\buildings"
#房间背景路径
ROOM_DIR = BASE_GAME_DIR + r"\gfx\portraits\city_sets"

##################################
##########图片输出文件位置##########
##################################

PERK_IMAGE_DIR = BASE_IMAGE_DIR + r"\飞升"
CIVIC_IMAGES_DIR = BASE_IMAGE_DIR + r"\国策"
ORIGIN_IMAGES_DIR = BASE_IMAGE_DIR + r"\起源"
TRADITIONS_IMAGES_DIR = BASE_IMAGE_DIR + r"\传统"
TRAITS_IMAGES_DIR = BASE_IMAGE_DIR + r"\物种特质"
TECHNO_IMAGES_DIR = BASE_IMAGE_DIR + r"\科技"
STATIC_MODIFIER_IMAGES_DIR = BASE_IMAGE_DIR + r"\全局修正"
DEPOSITS_IMAGES_DIR = BASE_IMAGE_DIR + r"\地块"
LEADER_TRAITS_IMAGES_DIR = BASE_IMAGE_DIR + r"\领袖特质"
ZONES_IMAGES_DIR = BASE_IMAGE_DIR + r"\区划特化"
BUILDINGS_IMAGES_DIR = BASE_IMAGE_DIR + r"\建筑"
SPECIMENS_IMAGES_DIR = BASE_IMAGE_DIR + r"\标本"
#高级政体图片输出目录
AUTHORITIES_IMAGES_DIR = BASE_IMAGE_DIR + r"\高级政体"
RELICS_IMAGES_DIR = BASE_IMAGE_DIR + r"\遗珍"
PLANETS_IMAGES_DIR = BASE_IMAGE_DIR + r"\行星"
ROOM_IMAGES_DIR = BASE_IMAGE_DIR + r"\房间"