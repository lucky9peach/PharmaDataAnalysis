import os
import time
import random
import requests
import pandas as pd
from datetime import datetime
import configparser
import json
from concurrent.futures import ThreadPoolExecutor, as_completed

def parse_years(time_string):
    """解析用户输入的时间段，支持如 '2022-2025' 或 '2022,2023'"""
    years = []
    if isinstance(time_string, list):
        return [str(y) for y in time_string]
        
    time_string = str(time_string)
    if '-' in time_string:
        parts = time_string.split('-')
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            years = [str(y) for y in range(int(parts[0]), int(parts[1])+1)]
    elif ',' in time_string or '，' in time_string:
        time_string = time_string.replace('，', ',')
        years = [y.strip() for y in time_string.split(',') if y.strip().isdigit()]
    else:
        if time_string.strip().isdigit():
            years = [time_string.strip()]
    return years

def get_parsed_drug_names(drug_names):
    if isinstance(drug_names, str):
        for char in [',', ';', '，', '；']:
            drug_names = drug_names.replace(char, ' ')
        raw_names = [d.strip().upper() for d in drug_names.split() if d.strip()]
    else:
        raw_names = [d.strip().upper() for d in drug_names if str(d).strip()]
        
    ignore_suffixes = {
        'SODIUM', 'HYDROCHLORIDE', 'POTASSIUM', 'CALCIUM', 'SULFATE', 
        'PHOSPHATE', 'MESYLATE', 'MALEATE', 'ACETATE', 'CHLORIDE',
        'BROMIDE', 'NITRATE', 'CITRATE', 'TARTRATE', 'SUCCINATE',
        'FUMARATE', 'LACTATE', 'MALATE', 'TOSYLATE', 'BESYLATE',
        'SALICYLATE', 'ASCORBATE', 'CARBONATE', 'HYDROXIDE', 'OXIDE',
        'PEROXIDE', 'SULFIDE', 'FLUORIDE', 'IODIDE', 'SULPHATE',
        'DIPROPIΟNATE', 'VALERATE', 'BUTYRATE', 'PROPIONATE', 'CAPROATE',
        'ENANTHATE', 'CYPIONATE', 'DECANOATE', 'UNDECANOATE', 'LAURATE',
        'PALMITATE', 'STEARATE', 'OLEATE', 'LINOLEATE', 'LINOLENATE',
        'ARACHIDONATE', 'PAMOATE', 'NAPADISILATE', 'ESTOLATE', 'ALPHACALCIDOL'
    }
    
    final_names = []
    for name in raw_names:
        if name not in ignore_suffixes and len(name) > 2:
            final_names.append(name)
            
    return final_names

def read_cache_file(args):
    filepath, drug_name = args
    try:
        tmp_df = pd.read_excel(filepath)
        if not tmp_df.empty:
            if '检索药名' not in tmp_df.columns:
                tmp_df['检索药名'] = drug_name
            return tmp_df
    except Exception as e:
        return e
    return None

def download_and_aggregate_tsm(
    drug_names,
    time_period,
    base_url="http://work.progames.top:3000",
    username=None,
    password=None,
    output_base_dir="TSM_Downloads",
    log_callback=print,
    # === NEW: Skip fetching from network if set to True ===
    skip_downloads=False
):
    
    parsed_names = get_parsed_drug_names(drug_names)
    years_to_fetch = parse_years(time_period)
    
    if not parsed_names:
        log_callback("[-] 未提取到任何有效药物名称，请检查输入。")
        return False, ""
        
    if not years_to_fetch:
        log_callback("[-] 年份格式解析失败，请使用如 '2023', '2022-2025', '2022,2023' 格式。")
        return False, ""
        
    log_callback(f"[*] 解析目标药物: {parsed_names}")
    log_callback(f"[*] 解析目标年份: {years_to_fetch}")
    
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.abspath(os.path.join(BASE_DIR, output_base_dir, "raw_files"))
    os.makedirs(output_dir, exist_ok=True)
    log_callback(f"[*] 已建/检查数据下载保存目录: {output_dir}")
    
    endpoints_to_try = ["MIDS", "ATC"]
    fields_to_try = ["mol", "name"]
    
    # --- PHASE 1: SCAN LOCAL CACHE ---
    log_callback("[*] 正在极速扫描本地缓存...")
    local_files = []
    missing_tasks = []
    
    for drug_name in parsed_names:
        for current_year in years_to_fetch:
            found_cache = False
            for ep in endpoints_to_try:
                filename = f"TSM_{current_year}_{drug_name}_{ep}.xlsx"
                filepath = os.path.join(output_dir, filename)
                if os.path.exists(filepath) and os.path.getsize(filepath) > 1024:
                    local_files.append((filepath, drug_name))
                    found_cache = True
                    break
            if not found_cache:
                missing_tasks.append((drug_name, current_year))
                
    log_callback(f"[+] 扫描完毕，发现本地拥有 {len(local_files)} 个匹配文件，缺失 {len(missing_tasks)} 个任务。")
    all_dfs = []
    
    if local_files:
        log_callback("[*] 开始多线程并发读取本地数据...")
        with ThreadPoolExecutor(max_workers=8) as executor:
            future_to_file = {executor.submit(read_cache_file, nf): nf for nf in local_files}
            for idx, future in enumerate(as_completed(future_to_file)):
                res_df = future.result()
                if isinstance(res_df, pd.DataFrame):
                    all_dfs.append(res_df)
                elif isinstance(res_df, Exception):
                    log_callback(f"[-] 并发读取缓存时遇到异常: {res_df}")
                if (idx + 1) % 10 == 0:
                    log_callback(f"    ...已读取 {idx + 1}/{len(local_files)} 个本地文件")
        log_callback("[+] 本地数据并发装载完成。")
        
    if skip_downloads:
        if missing_tasks:
            log_callback("[!] 用户选择跳过云端获取，将仅聚合上方扫描到的本地数据。")
        missing_tasks = []
        
    if not missing_tasks and not all_dfs:
         log_callback("\n[-] 没有任何有效数据被读取，不仅本地缺失而且云端也被跳过/拒绝。")
         return False, ""
         
    # --- PHASE 2: LOGIN & DOWNLOAD MISSING ---
    if missing_tasks:
        if username is None or password is None:
            config = configparser.ConfigParser()
            config_path = os.path.join(BASE_DIR, "config.ini")
            if os.path.exists(config_path):
                try:
                    config.read(config_path, encoding='utf-8')
                    if 'Credentials' in config:
                        username = username or config['Credentials'].get('username', '').strip()
                        password = password or config['Credentials'].get('password', '').strip()
                except Exception as e:
                    pass
        if not username or not password:
            log_callback("[-] 缺失任务需要云端获取，但未提供账号或密码（config.ini），中止云端环节。")
            missing_tasks = []
        
        if missing_tasks:
            log_callback("[*] 开始处理缺失任务...")
            session = requests.Session()
            session.headers.update({
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/json, text/plain, */*",
                "Accept-Language": "zh-CN,zh;q=0.9",
            })
            
            login_url = f"{base_url}/login"
            log_callback(f"[*] 正在尝试登录: {login_url}")
            login_payload = {"username": username, "password": password}
            try:
                resp = session.post(login_url, json=login_payload, timeout=20)
                if resp.status_code == 200:
                    body = resp.json()
                    if body.get('code') == 200:
                        log_callback("[+] 登录成功。")
                    else:
                        log_callback(f"[-] 登录失败: {body.get('msg')}。 提取可能受限。")
                else:
                    log_callback(f"[-] 登录请求 HTTP {resp.status_code}")
            except Exception as e:
                log_callback(f"[-] 网络异常或登录超时: {e}")
                
            for missing_drug, missing_year in missing_tasks:
                log_callback("-" * 50)
                log_callback(f">>> 开始拉取: {missing_drug} | {missing_year} <<<")
                success_for_year = False
                
                for ep in endpoints_to_try:
                    target_route = f"/search{ep}"
                    download2_url = f"{base_url}{target_route}/download2"
                    download_url = f"{base_url}{target_route}/download"
                    
                    for field in fields_to_try:
                        for attempt in range(1, 4):
                            log_callback(f"[*] 尝试次数 [{attempt}/3] - 端点 {ep} 字段 {field}")
                            payload = {field: missing_drug, "year": missing_year}
                            try:
                                resp = session.post(download2_url, data=payload, timeout=240)
                                if resp.status_code == 200:
                                    result = resp.text.strip()
                                    is_json_err = False
                                    try:
                                        jdata = json.loads(result)
                                        is_json_err = True  
                                    except: 
                                        pass
                                    
                                    if is_json_err:
                                        log_callback(f"    [-] TSM数据库提示无此数据或发生错误。响应: {result[:50]}...")
                                        success_for_year = True  
                                        break
                                    if not is_json_err and result and len(result) < 200 and "<html" not in result.lower():
                                        file_url = f"{download_url}?{result}"
                                        try:
                                            dll_resp = session.get(file_url, timeout=180)
                                            content_size = len(dll_resp.content)
                                            
                                            if dll_resp.status_code == 200 and content_size > 1024:
                                                filename = f"TSM_{missing_year}_{missing_drug}_{ep}.xlsx"
                                                filepath = os.path.join(output_dir, filename)
                                                with open(filepath, "wb") as f:
                                                    f.write(dll_resp.content)
                                                    
                                                log_callback(f"[+++] {missing_year} {missing_drug} 成功拉取! 保存至 {filename} ({content_size} byte)")
                                                success_for_year = True
                                                sleep_t = random.uniform(1, 2)
                                                log_callback(f"[*] 礼貌等待 {sleep_t:.1f} 秒...")
                                                time.sleep(sleep_t)
                                                
                                                try:
                                                    tmp_df = pd.read_excel(filepath)
                                                    if tmp_df.empty:
                                                        log_callback(f"    [-] 表格为空，跳过合并。")
                                                        break
                                                    
                                                    if '检索药名' not in tmp_df.columns:
                                                        tmp_df['检索药名'] = missing_drug
                                                        
                                                    all_dfs.append(tmp_df)
                                                    log_callback(f"    [+] 成功读取 {len(tmp_df)} 行原始数据已拼装。")
                                                except Exception as e:
                                                    log_callback(f"    [-] 读取/处理新下载 {filename} 失败: {e}")
                                                break
                                            else:
                                                log_callback(f"    [-] 文件异常/过小 ({content_size} bytes)")
                                        except requests.exceptions.Timeout:
                                            log_callback(f"    [-] 下载文件时超时。")
                                        except Exception as e:
                                            log_callback(f"    [-] 下载发生不可预计的错误: {e}")
                            except requests.exceptions.Timeout:
                                log_callback(f"    [-] 队列查询超时。")
                            except Exception as e:
                                log_callback(f"    [-] 队列查询异常: {e}")
                                
                            if attempt == 3 and not success_for_year:
                                 log_callback(f"    [-] 已达到最大尝试次数，放弃云端重试。")
                                 
                        if success_for_year:
                            break
                    if success_for_year:
                        break
                
                if not success_for_year:
                    log_callback(f"[-] 警告: {missing_drug} 在 {missing_year} 拉取全部失败。已自动跳过。")
            
            try:
                logout_url = f"{base_url}/logout"
                session.get(logout_url, timeout=10)
                log_callback("[+] 已主动退出登录，账号安全。")
            except Exception as e:
                pass
            finally:
                session.close()

    # --- PHASE 3: AGGREGATE CACHE ---
    if not all_dfs:
        log_callback("\n[-] 没有任何有效数据被下载和读取，任务结束。")
        return False, ""
        
    log_callback("\n" + "="*50)
    log_callback(f"[*] 开始聚合所有抓取的原始数据，总共 {len(all_dfs)} 份表格数据...")
    try:
        merged_df = pd.concat(all_dfs, ignore_index=True)
        log_callback(f"[+] 数据合并完成，总计 {len(merged_df)} 行 x {len(merged_df.columns)} 列")
        
        # 将合并后的数据保存为 .csv 以突破 Excel xlsx 1,048,576 行的格式限制
        log_callback(f"[*] 正在将紧密数据缓存保存为 CSV（容量无上限）...")
        cache_dir = os.path.abspath(os.path.join(BASE_DIR, "Cache"))
        os.makedirs(cache_dir, exist_ok=True)
        
        # 使用 utf-8-sig 并用 csv，这样即使直接用 Excel 双击点开中文也不会产生乱码，同时完美规避 xlsx 报错
        final_file = os.path.join(cache_dir, "step1_latest.csv")
        merged_df.to_csv(final_file, index=False, encoding="utf-8-sig")
        log_callback(f"[+] 聚合 CSV 保存成功: \n{final_file}")
        
    except Exception as e:
        log_callback(f"[-] 保存聚合 CSV 失败: {e}\n[*] 若为文件被占用，请先关闭相关 Excel!")
        return False, ""
        
    return True, final_file
