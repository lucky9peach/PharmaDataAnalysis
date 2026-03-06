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
        # 使用 openpyxl 读取避免告警
        df = pd.read_excel(input_excel_path)
    except Exception as e:
        log_callback(f"[-] 读取 Excel 失败: {e}")
        return False, ""
        
    log_callback(f"[*] 获取到 {len(df)} 行数据，执行头文件规范化与清洗...")
    
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
    basename = os.path.basename(input_excel_path)
    parquet_filename = os.path.splitext(basename)[0] + ".parquet"
    output_parquet_path = os.path.abspath(os.path.join(output_parquet_dir, parquet_filename))
    
    try:
        # 注意: 如果抛出缺失依赖的异常，需要用户在终端安装 pyarrow: pip install pyarrow
        df.to_parquet(output_parquet_path, engine='pyarrow', index=False)
        log_callback(f"[+++] Parquet 转换完成! 已缓存至: {output_parquet_path} (共 {len(df)} 行记录)")
        return True, output_parquet_path
    
    except ImportError as e:
        log_callback("[-] 错误：缺少 Parquet 支持库！请在终端运行: pip install pyarrow fastparquet")
        return False, ""
    except Exception as e:
        log_callback(f"[-] Parquet 生成失败: {e}")
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
