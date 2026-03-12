import os
import pandas as pd
from core_config import COLUMN_MAPPING, COLUMN_TYPES, DOSAGE_CODE_MAP

def clean_and_cache_data(input_excel_path, output_parquet_dir="Cache", log_callback=print):
    """
    Step B: Clean 模块接口
    
    读取 A 模块生成的 Excel，提取中文剂型，使用 core_config 中的 COLUMN_MAPPING 
    进行表头标准化，使用 COLUMN_TYPES 强转数据类型，最后将清洗好的数据保存为 Parquet 文件。
    
    参数:
    - input_excel_path: Step A 聚合生成的 Excel 文件路径
    - output_parquet_dir: 生成的 Parquet 缓存文件保存目录
    - log_callback: 日志回调函数（接入到主程序的 UI 控制台）
    
    返回:
    - (bool, str) -> (是否成功, 生成的 Parquet 文件绝对路径)
    """
    if not os.path.exists(input_excel_path):
        log_callback(f"[-] 输入文件不存在: {input_excel_path}")
        return False, ""
        
    log_callback(f"[*] 解析并清洗数据: {os.path.basename(input_excel_path)}")
    try:
        if input_excel_path.lower().endswith('.csv'):
            df = pd.read_csv(input_excel_path, encoding='utf-8-sig', low_memory=False)
        else:
            df = pd.read_excel(input_excel_path)
    except Exception as e:
        log_callback(f"[-] 读取原始文件失败: {e}")
        return False, ""
        
    if df.empty or '检索药名' not in df.columns:
        log_callback("[-] 表格为空或缺失 '检索药名' 列，无法判定缓存映射。")
        return False, ""

    log_callback(f"[*] 完整数据加载: {len(df)} 行，开始执行增量缓存校验...")
    
    # === NEW: Cache Bypass Logic ===
    BASE_DIR = os.path.dirname(os.path.abspath(str(input_excel_path)))
    # input_excel_path 通常在 "Cache" 目录下，比如 "C/Cache/step1_latest.csv"
    # 我们需要往上一级找 TSM_Downloads
    if "Cache" in str(BASE_DIR):
        BASE_DIR = os.path.dirname(str(BASE_DIR))
        
    api_views_dir = os.path.join(str(BASE_DIR), "Cache", "API_Views")
    raw_files_dir = os.path.join(str(BASE_DIR), "TSM_Downloads", "raw_files")
    
    apis = df['检索药名'].dropna().unique().tolist()
    apis_to_drop = []
    
    for api in apis:
        api_str = str(api).strip().upper()
        if not api_str: continue
        
        safe_filename = api_str.replace("/", "_").replace("\\", "_")
        parquet_path = os.path.join(api_views_dir, f"core_cache_{safe_filename}.parquet")
        
        # 1. 检查是否存在最终分片缓存
        if os.path.exists(parquet_path):
            cache_mtime = os.path.getmtime(parquet_path)
            
            # 2. 获取针对该药所有的原始下载文件的时间戳
            raw_is_newer = False
            if os.path.exists(raw_files_dir):
                for f in os.listdir(raw_files_dir):
                    if f.endswith('.xlsx') and f"_{api_str}_" in f.upper():
                        raw_path = os.path.join(raw_files_dir, f)
                        if os.path.exists(raw_path) and os.path.getmtime(raw_path) > cache_mtime:
                            raw_is_newer = True
                            break
                            
            # 3. 如果原始文件没有更新过缓存，说明该缓存100%是对齐且生效的
            if not raw_is_newer:
                apis_to_drop.append(api)
                log_callback(f"    [!] 旁路跳过: {api_str} (因后续核心缓存未过期，无需重复清洗)")

    if apis_to_drop:
        # 将不需要重新洗的 API 行移除
        df = df[~df['检索药名'].isin(apis_to_drop)]
        log_callback(f"[*] 因存在新鲜缓存，剥离了 {len(apis_to_drop)} 个无需重算的药物，剩余 {len(df)} 行记录等待清洗。")
        
    if df.empty:
        log_callback("[+] 当前队列中所有药物均具备最新分析缓存，无需发生任何重复清洗与运算！")
        return True, ""
        
    log_callback(f"[*] 开始进行字典映射与结构标准化清洗...")
    
    # 1. 探测与提取中文剂型
    if "中文剂型" not in df.columns:
        log_callback("[*] 正在从数据特征中解析 [中文剂型] 字段...")
        def map_form_val(x):
            x_str = str(x).upper()
            for k, v in DOSAGE_CODE_MAP.items():
                if k in x_str:
                    return v
            return ''
        
        form_assigned = False
        for col in df.columns:
            if 'NFC' in str(col).upper() or '剂型' in str(col):
                found_forms = df[col].apply(map_form_val)
                if (found_forms != '').any():
                    df['中文剂型'] = found_forms
                    form_assigned = True
                    log_callback(f"    [+] 成功解析 [中文剂型], 匹配原始列名: {col}")
                    break
        
        if not form_assigned:
            # 高级特征自动探测
            for col in df.columns:
                sample = df[col].dropna().astype(str).head(10).str.upper()
                if sample.apply(lambda val: any(k in val for k in DOSAGE_CODE_MAP.keys())).any():
                    df['中文剂型'] = df[col].apply(map_form_val)
                    log_callback(f"    [+] 成功解析 [中文剂型], 自动探测匹配原始列名: {col}")
                    break

    # 2. 列名映射 (Raw -> Core 统一表头)
    rename_mapping = {}
    for col in df.columns:
        # 精确匹配
        if col in COLUMN_MAPPING:
            rename_mapping[col] = COLUMN_MAPPING[col]
        else:
            # 模糊匹配容错处理，比如"通用名单(INNN)" -> "molecule", "销售额(US$)" -> "sales_value"
            for k, v in COLUMN_MAPPING.items():
                if k in str(col):
                    rename_mapping[col] = v
                    break
                    
    df = df.rename(columns=rename_mapping)
    
    # 3. 列筛选与结构对齐（确保具备所有核心列）
    core_cols = list(COLUMN_TYPES.keys())
    # 额外保留 "检索药名"，用来追溯来源
    keep_cols = [c for c in df.columns if c in core_cols or c == '检索药名']
    
    for c in core_cols:
        if c not in df.columns:  # Rename后的 dataframe 中没有这个标准列
            df[c] = None
            log_callback(f"    [-] 预定义列 [{c}] 在源数据中未找到，已自动补齐空值结构")
            keep_cols.append(c)
            
    # 去重保留列序列
    keep_cols = list(dict.fromkeys(keep_cols))
    df = df[keep_cols]
    
    # 4. 类型转换及缺省容错处理
    for col, dtype in COLUMN_TYPES.items():
        if col in df.columns:
            try:
                if dtype == "string":
                    df[col] = df[col].fillna("").astype(str)
                    df[col] = df[col].replace("nan", "").replace("None", "").str.strip()
                elif dtype == "int":
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
                elif dtype == "float":
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
            except Exception as e:
                log_callback(f"    [-] 异常: 字段 [{col}] 转为 [{dtype}] 时发生错误: {e}")
                
    log_callback("[*] 标准化流程完毕，准备固化为 Parquet 缓存模型")
    
    # 5. 写入高速 Parquet 缓存
    os.makedirs(output_parquet_dir, exist_ok=True)
    # 我们固定写出一个干净的中转 parquet
    final_file = os.path.join(output_parquet_dir, "step2_cleaned.parquet")
    
    try:
        # 注意: 如果抛出缺失依赖的异常，需要用户在终端安装 pyarrow: pip install pyarrow
        df.to_parquet(final_file, engine='pyarrow', index=False)
        log_callback(f"[+++] 维度清洗与标准转型完成: {final_file} (共 {len(df)} 行衍生增量记录)")
        return True, final_file
    
    except ImportError as e:
        log_callback("[-] 错误：缺少 Parquet 支持库！请在终端运行: pip install pyarrow fastparquet")
        return False, ""
    except Exception as e:
        log_callback(f"[-] 保存清洗缓存 Parquet 失败: {e}")
        return False, ""

# ==========================================
# 本地测试入口
# ==========================================
if __name__ == "__main__":
    def test_logger(msg):
        print(f"[Step B] {msg}")
        
    print(">> 开始测试 Step B (Clean 模块)...")
    
    # 这里需要写死一个用于测试的 A模块生成的 Excel 路径（你需要指向你下载好的一份文件）
    test_excel = "TSM_Downloads/填写这里的一个实际存在的Excel文件如_StepA_Merged_2022-2025.xlsx"
    
    if os.path.exists(test_excel):
        success, parquet_path = clean_and_cache_data(
            input_excel_path=test_excel, 
            output_parquet_dir="Cache", 
            log_callback=test_logger
        )
        if success:
            print(f"\n>> 测试执行成功, C模块(标准模型化)可直接读取此 Parquet：{parquet_path}")
    else:
        print(f">> [提示] 测试文件不存在，未能完成独立测试流程 ({test_excel})")
