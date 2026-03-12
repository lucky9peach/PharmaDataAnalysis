"""
Core Configuration
Global Pharma Market Analytics Pipeline
Version: 1.0
"""

# ==========================================================
# 1️⃣ 标准字段定义（Raw → Core 统一命名）
# ==========================================================

COLUMN_MAPPING = {
    "国家": "country",
    "年份": "year",
    "通用名单": "molecule",
    "中文剂型": "dosage_form",
    "集团/企业": "mah",
    "规格": "strength_raw",
    "销售额": "sales_value",
    "最小单包装销售数量": "units_small",
    "大包装销售数量": "units_large",
    "公斤": "api_kg",
    # --- 新增的扩展维度 ---
    "渠道": "channel",
    "商品名": "brand_name",
    "生产企业": "manufacturer",
    "盐/离子": "salt_ion",
    "通用名复": "generic_name_complex"
}


# ==========================================================
# 2️⃣ 数据类型定义（防止不同脚本类型混乱）
# ==========================================================

COLUMN_TYPES = {
    "country": "string",
    "year": "int",
    "molecule": "string",
    "dosage_form": "string",
    "mah": "string",
    "strength_raw": "string",
    "sales_value": "float",
    "units_small": "float",
    "units_large": "float",
    "api_kg": "float",
    # --- 新增的扩展维度 ---
    "channel": "string",
    "brand_name": "string",
    "manufacturer": "string",
    "salt_ion": "string",
    "generic_name_complex": "string"
}


# ==========================================================
# 3️⃣ Core 层派生字段（只做行级计算，不做聚合）
# ==========================================================

DERIVED_COLUMNS = {

    "pack_size": {
        "formula": "units_small / units_large",
        "description": "每盒多少片"
    },

    "price_per_unit": {
        "formula": "sales_value / units_small",
        "description": "每片价格"
    },

    "estimated_ex_factory_price": {
        "formula": "price_per_unit / 3",
        "description": "预估出厂价（简单假设）"
    },

    "api_per_unit": {
        "formula": "api_kg / units_small",
        "description": "单位API消耗"
    }
}


# ==========================================================
# 4️⃣ 目标市场定义
# ==========================================================

EU_TARGET_MARKETS = [
    "UNITED KINGDOM", "GERMANY", "FRANCE", "ITALY", "SPAIN",
    "NETHERLANDS", "BELGIUM", "SWEDEN", "SWITZERLAND", "AUSTRIA",
    "POLAND", "PORTUGAL", "GREECE", "CZECH REPUBLIC", "HUNGARY",
    "ROMANIA", "IRELAND", "DENMARK", "FINLAND", "NORWAY"
]

EU_BIG5 = [
    "UNITED KINGDOM", "GERMANY", "FRANCE", "ITALY", "SPAIN"
]

US_MARKET = ["UNITED STATES", "US", "USA", "美国"]

# ==========================================================
# 4️⃣b 原研药知识库 (Originator Config) — 关键词匹配公司名
# ==========================================================
ORIGINATOR_CONFIG = {
    "APIXABAN":     ["BRISTOL", "PFIZER", "SQUIBB"],
    "CITALOPRAM":   ["LUNDBECK"],
    "RIVAROXABAN":  ["BAYER"],
    "IRBESARTAN":   ["SANOFI", "BRISTOL", "SQUIBB"],
    "RANOLAZINE":   ["GILEAD"],
    "CLOPIDOGREL":  ["SANOFI", "BRISTOL", "SQUIBB"],
    "CELECOXIB":    ["PFIZER"],
    "LAMOTRIGINE":  ["GSK"],
    "TOPIRAMATE":   ["JANSSEN"],
    "DULOXETINE":   ["LILLY"],
    "BISOPROLOL":   ["MERCK"],
    "TICAGRELOR":   ["ASTRAZENECA"],
    "MEMANTINE":    ["MERZ"],
    "TADALAFIL":    ["LILLY"],

    # --- 以下为自动扫描发现的新药，请手动补充原研企业 ---
    "ACETOHEXAMIDE":   [],  # TODO: 请填写原研企业关键词
    "ACETONIDE":       [],  # TODO: 请填写原研企业关键词
    "AMIODARONE":      [],  # TODO: 请填写原研企业关键词
    "AMITRIPTYLINE":   [],  # TODO: 请填写原研企业关键词
    "ANASTROZOLE":     [],  # TODO: 请填写原研企业关键词
    "ATROPINE":        [],  # TODO: 请填写原研企业关键词
    "BACLOFEN":        [],  # TODO: 请填写原研企业关键词
    "BRIMONIDINE":     [],  # TODO: 请填写原研企业关键词
    "BROMPHENIRAMINE": [],  # TODO: 请填写原研企业关键词
    "BUPROPION":       [],  # TODO: 请填写原研企业关键词
    "CALCITONIN":      [],  # TODO: 请填写原研企业关键词
    "CARBAMAZEPINE":   [],  # TODO: 请填写原研企业关键词
    "CHLORDIAZEPOXIDE": [],  # TODO: 请填写原研企业关键词
    "CHLORPROPAMIDE":  [],  # TODO: 请填写原研企业关键词
    "CHOLINE":         [],  # TODO: 请填写原研企业关键词
    "CIMETIDINE":      [],  # TODO: 请填写原研企业关键词
    "CLOBAZAM":        [],  # TODO: 请填写原研企业关键词
    "CLOFIBRATE":      [],  # TODO: 请填写原研企业关键词
    "CODEINE":         [],  # TODO: 请填写原研企业关键词
    "DESIPRAMINE":     [],  # TODO: 请填写原研企业关键词
    "DEXAMETHASONE":   [],  # TODO: 请填写原研企业关键词
    "DICLOFENAC":      [],  # TODO: 请填写原研企业关键词
    "DIPHENHYDRAMINE": [],  # TODO: 请填写原研企业关键词
    "DIPHENOXYLATE":   [],  # TODO: 请填写原研企业关键词
    "DROXIDOPA":       [],  # TODO: 请填写原研企业关键词
    "ESTRADIOL":       [],  # TODO: 请填写原研企业关键词
    "FENOFIBRATE":     [],  # TODO: 请填写原研企业关键词
    "FENOPROFEN":      [],  # TODO: 请填写原研企业关键词
    "FLUOCINOLONE":    [],  # TODO: 请填写原研企业关键词
    "FLUOXETINE":      [],  # TODO: 请填写原研企业关键词
    "FLUOXYMESTERONE": [],  # TODO: 请填写原研企业关键词
    "FLUPHENAZINE":    [],  # TODO: 请填写原研企业关键词
    "FLURAZEPAM":      [],  # TODO: 请填写原研企业关键词
    "FLUVOXAMINE":     [],  # TODO: 请填写原研企业关键词
    "FOSINOPRIL":      [],  # TODO: 请填写原研企业关键词
    "GLUTETHIMIDE":    [],  # TODO: 请填写原研企业关键词
    "HYDRALAZINE":     [],  # TODO: 请填写原研企业关键词
    "HYDROCHLOROTHIAZIDE": [],  # TODO: 请填写原研企业关键词
    "HYDROCODONE":     [],  # TODO: 请填写原研企业关键词
    "HYDROCORTISONE":  [],  # TODO: 请填写原研企业关键词
    "HYDROFLUMETHIAZIDE": [],  # TODO: 请填写原研企业关键词
    "HYDROXYZINE":     [],  # TODO: 请填写原研企业关键词
    "IMIPRAMINE":      [],  # TODO: 请填写原研企业关键词
    "LETROZOLE":       [],  # TODO: 请填写原研企业关键词
    "LORAZEPAM":       [],  # TODO: 请填写原研企业关键词
    "MEDROXYPROGESTERONE": [],  # TODO: 请填写原研企业关键词
    "MEGESTROL":       [],  # TODO: 请填写原研企业关键词
    "MEPROBAMATE":     [],  # TODO: 请填写原研企业关键词
    "METFORMIN":       [],  # TODO: 请填写原研企业关键词
    "METHOCARBAMOL":   [],  # TODO: 请填写原研企业关键词
    "METHYCLOTHIAZIDE": [],  # TODO: 请填写原研企业关键词
    "METHYLTESTOSTERONE": [],  # TODO: 请填写原研企业关键词
    "METOCLOPRAMIDE":  [],  # TODO: 请填写原研企业关键词
    "MINOXIDIL":       [],  # TODO: 请填写原研企业关键词
    "MIRTAZAPINE":     [],  # TODO: 请填写原研企业关键词
    "OSELTAMIVIR":     [],  # TODO: 请填写原研企业关键词
    "OXANDROLONE":     [],  # TODO: 请填写原研企业关键词
    "OXYBUTYNIN":      [],  # TODO: 请填写原研企业关键词
    "PAROXETINE":      [],  # TODO: 请填写原研企业关键词
    "PENTOXIFYLLINE":  [],  # TODO: 请填写原研企业关键词
    "PHENDIMETRAZINE": [],  # TODO: 请填写原研企业关键词
    "PHENTERMINE":     [],  # TODO: 请填写原研企业关键词
    "PRAZEPAM":        [],  # TODO: 请填写原研企业关键词
    "PREDNISONE":      [],  # TODO: 请填写原研企业关键词
    "PROMETHAZINE":    [],  # TODO: 请填写原研企业关键词
    "PROPAFENONE":     [],  # TODO: 请填写原研企业关键词
    "PROPRANOLOL":     [],  # TODO: 请填写原研企业关键词
    "PSEUDOEPHEDRINE": [],  # TODO: 请填写原研企业关键词
    "QUINIDINE":       [],  # TODO: 请填写原研企业关键词
    "RESERPINE":       [],  # TODO: 请填写原研企业关键词
    "SALMON":          [],  # TODO: 请填写原研企业关键词
    "SOTALOL":         [],  # TODO: 请填写原研企业关键词
    "SPIRONOLACTONE":  [],  # TODO: 请填写原研企业关键词
    "SULFAMETHOXAZOLE": [],  # TODO: 请填写原研企业关键词
    "TEMAZEPAM":       [],  # TODO: 请填写原研企业关键词
    "TESTOSTERONE":    [],  # TODO: 请填写原研企业关键词
    "TIMOLOL":         [],  # TODO: 请填写原研企业关键词
    "TOLAZAMIDE":      [],  # TODO: 请填写原研企业关键词
    "TRAZODONE":       [],  # TODO: 请填写原研企业关键词
    "TRIMETHOPRIM":    [],  # TODO: 请填写原研企业关键词
    "TRIMIPRAMINE":    [],  # TODO: 请填写原研企业关键词
    "TRIPROLIDINE":    [],  # TODO: 请填写原研企业关键词
    "TROSPIUM":        [],  # TODO: 请填写原研企业关键词
    "WARFARIN":        [],  # TODO: 请填写原研企业关键词
    "ZALEPLON":        [],  # TODO: 请填写原研企业关键词
    "ZONISAMIDE":      [],  # TODO: 请填写原研企业关键词
}


# ==========================================================
# 5️⃣ 市场分析配置（只控制聚合维度）
# ==========================================================

MARKET_CONFIG = {

    "EU": {
        "countries": EU_TARGET_MARKETS,
        "group_by": ["country", "mah"],
        "metrics": ["sales_value"],
    },

    "EU_BIG5": {
        "countries": EU_BIG5,
        "group_by": ["country", "mah"],
        "metrics": ["sales_value"],
    },

    "US": {
        "countries": US_MARKET,
        "group_by": ["country", "mah", "dosage_form"],
        "metrics": ["sales_value", "api_kg", "price_per_unit"]
    }
}


# ==========================================================
# 6️⃣ 剂型代码映射
# ==========================================================

DOSAGE_CODE_MAP = {
    # 原始三位编码
    'AAA': '普通片剂', 'AAB': '口崩片', 'AAE': '颊含片', 'AAF': '舌下片',
    'AAG': '咀嚼片', 'AAH': '泡腾片', 'AAJ': '多层片', 'AAK': '可溶片',
    'AAY': '其他片剂', 'AAZ': '片剂组合包', 'ABA': '包衣片', 'ABB': '明胶包衣片',
    'ABC': '薄膜衣片', 'ABD': '肠溶片', 'ABG': '包衣咀嚼片', 'ABY': '其他包衣片',
    'ABZ': '包衣片组合包', 'BAA': '缓释片', 'BAB': '缓释口崩片', 'BAE': '缓释颊含片',
    'BAJ': '缓释多层片', 'BAY': '其他缓释片', 'BBA': '缓释包衣片', 'BBC': '缓释薄膜衣片',
    'BBD': '缓释肠溶片', 'BBN': '缓释膜包衣片', 'BBY': '其他缓释包衣片', 'BBZ': '缓释包衣片组合',
    'ACA': '普通胶囊', 'ACD': '肠溶胶囊', 'ACF': '咬碎胶囊', 'ACG': '咀嚼胶囊',
    'ACS': '扁胶囊', 'ACY': '其他胶囊', 'ACZ': '胶囊组合包', 'BCA': '缓释胶囊',
    'BCN': '缓释膜胶囊', 'BCY': '其他缓释胶囊',
    
    # 追加中文剂型模糊映射与部分外文简写兜底
    '薄膜衣片': '薄膜衣片', 'FILM-COATED TABLET': '薄膜衣片', 'FC TABLET': '薄膜衣片',
    '肠溶片': '肠溶片', 'ENTERIC-COATED TABLET': '肠溶片', 'EC TABLET': '肠溶片',
    '缓释片': '缓释片', 'SUSTAINED-RELEASE TABLET': '缓释片', 'SR TABLET': '缓释片', 'ER TABLET': '缓释片',
    '口崩片': '口崩片', 'ORALLY DISINTEGRATING TABLET': '口崩片', 'ODT': '口崩片',
    '咀嚼片': '咀嚼片', 'CHEWABLE TABLET': '咀嚼片',
    '泡腾片': '泡腾片', 'EFFERVESCENT TABLET': '泡腾片',
    '普通片剂': '普通片剂', '片剂': '普通片剂', 'TABLET': '普通片剂', 'TAB': '普通片剂', 'TABLETS': '普通片剂',
    '片': '普通片剂',
    
    '肠溶胶囊': '肠溶胶囊', 'ENTERIC-COATED CAPSULE': '肠溶胶囊',
    '缓释胶囊': '缓释胶囊', 'SUSTAINED-RELEASE CAPSULE': '缓释胶囊',
    '软胶囊': '软胶囊', 'SOFT CAPSULE': '软胶囊',
    '普通胶囊': '普通胶囊', '胶囊': '普通胶囊', 'CAPSULE': '普通胶囊', 'CAP': '普通胶囊', 'CAPS': '普通胶囊'
}


# ==========================================================
# 7️⃣ 全局常量（未来扩展）
# ==========================================================

DEFAULT_CURRENCY = "EUR"

SUPPORTED_MARKETS = ["EU", "EU_BIG5", "US"]

PROJECT_VERSION = "1.0.0"