import os
import pandas as pd
import core_config

class StandardizationEngine:
    """
    Step 3: 数据标准化与 API 分片缓存模块
    将清洗后的全量源数据打散、标准化、并执行业务逻辑衍生核算，最终生成按药物 (API) 隔离的分片缓存。
    """
    def __init__(self, input_path="Cache/raw.parquet", output_dir="Cache/API_Views/"):
        self.input_path = input_path
        self.output_dir = output_dir

    def _flag_originator(self, api_name, corporation_name):
        """
        利用 core_config.ORIGINATOR_MAPPING 判断是否为原研药企业
        支持映射配置中定义为列表(多个原研厂)或单字符串的情形
        """
        # 安全获取映射字典（防止配置尚未完全写入）
        mapping = getattr(core_config, 'ORIGINATOR_MAPPING', {})
        originator_companies = mapping.get(api_name, [])
        
        if isinstance(originator_companies, str):
            originator_companies = [originator_companies]
            
        corp_str = str(corporation_name).upper()
        for org_company in originator_companies:
            if org_company.upper() in corp_str:
                return True
        return False

    def execute_standardization(self, progress_callback=None, log_callback=None):
        """
        执行主流程：读取 -> 重命名 -> 分组衍生 -> 落盘
        可以通过两组 callback 回调，与主 GUI 无缝整合进度与日志。
        """
        def log(msg):
            if log_callback:
                log_callback(msg)
            else:
                print(msg)
                
        def progress(current, total):
            if progress_callback:
                progress_callback(current, total)

        log("[*] 初始化标准化处理引擎 (StandardizationEngine)...")
        log(f"[*] 数据输入源: {self.input_path}")
        
        # 1. 尝试读取原始全量缓存
        if not os.path.exists(self.input_path):
            log(f"[-] 致命错误: 无法找到输入数据源 {self.input_path}")
            return False
            
        try:
            ingested_raw_df = pd.read_parquet(self.input_path)
        except Exception as e:
            log(f"[-] 读取 Parquet 失败: {e}")
            return False
            
        log(f"[+] 成功装载 {len(ingested_raw_df)} 行 ingested raw data。")

        # 2. 严格的命名空间转换 (映射到企业级命名)
        rename_dict = {
            "molecule": "api_name",
            "country": "market_region",
            "mah": "corporation_name",
            "units_small": "sales_volume_units",
            "sales_value": "sales_value_usd"
        }
        
        # 兼容性匹配（防止 B 模块传过来的还是中文物理列名的情况）
        fallback_rename = {
            "通用名单": "api_name",
            "国家": "market_region",
            "集团/企业": "corporation_name",
            "最小单包装销售数量": "sales_volume_units",
            "销售额": "sales_value_usd",
            "公斤": "api_kg"
        }
        
        for old_col, new_col in fallback_rename.items():
            if old_col in ingested_raw_df.columns and new_col not in ingested_raw_df.columns:
                rename_dict[old_col] = new_col

        ingested_raw_df = ingested_raw_df.rename(columns=rename_dict)
        
        # 强制将下载下发的输入药名词汇当做唯一的分类基准（Fix: API基于盐基导致无限裂变问题）
        if "检索药名" in ingested_raw_df.columns:
            ingested_raw_df["api_name"] = ingested_raw_df["检索药名"].astype(str).str.strip().str.upper()
        
        # 统一美国市场命名规范
        if "market_region" in ingested_raw_df.columns:
            ingested_raw_df["market_region"] = ingested_raw_df["market_region"].str.strip().str.upper()
            us_mask = ingested_raw_df["market_region"].isin(["US", "USA", "UNITED STATES", "US MARKET", "美国"])
            ingested_raw_df.loc[us_mask, "market_region"] = "美国"
        
        # 确保关键核心业务字段存在
        essential_cols = ["api_name", "market_region", "corporation_name", "sales_volume_units", "sales_value_usd"]
        for col in essential_cols:
            if col not in ingested_raw_df.columns:
                log(f"[-] 错误: 源数据缺失关键列 '{col}'")
                return False

        # 确保数值字段计算安全
        ingested_raw_df['sales_volume_units'] = pd.to_numeric(ingested_raw_df['sales_volume_units'], errors='coerce').fillna(0.0)
        ingested_raw_df['sales_value_usd'] = pd.to_numeric(ingested_raw_df['sales_value_usd'], errors='coerce').fillna(0.0)
        if 'api_kg' in ingested_raw_df.columns:
            ingested_raw_df['api_kg'] = pd.to_numeric(ingested_raw_df['api_kg'], errors='coerce').fillna(0.0)

        # 3. 按 API 分组处理准备
        unique_apis = ingested_raw_df["api_name"].dropna().unique().tolist()
        unique_apis = [api.strip() for api in unique_apis if str(api).strip() != ""]
        
        if not unique_apis:
            log("[-] 源数据中未发现有效的 api_name，退出引擎。")
            return False
            
        total_apis = len(unique_apis)
        log(f"[*] 侦测到 {total_apis} 个独立 API，开始计算衍生指标与分片写出...")
        
        # 清空旧缓存防止残留
        import shutil
        if os.path.exists(self.output_dir):
            shutil.rmtree(self.output_dir, ignore_errors=True)
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 开始循环处理每个 API
        for idx, current_api in enumerate(unique_apis):
            progress(idx, total_apis)
            
            # 分离出当前的 API 视图拷贝
            standardized_api_df = ingested_raw_df[ingested_raw_df["api_name"] == current_api].copy()
            
            if standardized_api_df.empty:
                continue
                
            record_count = len(standardized_api_df)
            log(f"    -> [{idx+1}/{total_apis}] 正在计算 API: {current_api} (共 {record_count} 条记录)")

            # ------- 4. 业务衍生字段计算区 -------
            
            # 衍生指标 A: 原研判定 (is_originator)
            standardized_api_df["is_originator"] = standardized_api_df["corporation_name"].apply(
                lambda corp: self._flag_originator(current_api, corp)
            )

            # 衍生指标 B: 预估出厂单价 (factory_price_est = 终端单价 * 0.3)
            # 计算单价：避免除以零产生的 Inf/NaN 问题
            unit_terminal_price = standardized_api_df["sales_value_usd"] / standardized_api_df["sales_volume_units"].replace({0: pd.NA})
            standardized_api_df["factory_price_est"] = (unit_terminal_price * 0.3).fillna(0.0)

            # 衍生指标 C: 市场权重/全球用量占比 (api_global_share)
            api_total_volume = standardized_api_df["sales_volume_units"].sum()
            if api_total_volume > 0:
                standardized_api_df["api_global_share"] = (standardized_api_df["sales_volume_units"] / api_total_volume).astype(float)
            else:
                standardized_api_df["api_global_share"] = 0.0

            # ------- 5. 分片回写保护 -------
            
            safe_filename = str(current_api).replace("/", "_").replace("\\", "_")
            out_file = os.path.join(self.output_dir, f"core_cache_{safe_filename}.parquet")
            
            try:
                standardized_api_df.to_parquet(out_file, engine='pyarrow', index=False)
            except Exception as e:
                log(f"    [-] Parquet 回写分片 {safe_filename} 失败: {e}")
                
        # 循环结束收尾
        progress(total_apis, total_apis)
        log("\n" + "="*50)
        log(f"[+++] 核心处理完成！全部分片( {total_apis} 份)生成于: {os.path.abspath(self.output_dir)}")
        return True

# ==========================================
# 本地联调/独立测试沙盒
# ==========================================
if __name__ == "__main__":
    def mock_progress(current, total):
        print(f"[进度状态]: {current}/{total} " + "=" * int(current/total * 20) if total else "")
        
    def mock_log(msg):
        print(f"[Engine] {msg}")

    print(">>> 启动测试沙盒 (Step 3: Standardization Engine) <<<")
    
    # 假设输入为 Step2 B 模块生成的缓存
    tester = StandardizationEngine(
        input_path="Cache/这里填入你的_Step2_Cache.parquet",
        output_dir="Cache/API_Views/"
    )
    
    # 实际测试取消下行注释
    # tester.execute_standardization(progress_callback=mock_progress, log_callback=mock_log)
