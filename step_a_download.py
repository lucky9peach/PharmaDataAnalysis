import os
import time
import random
import requests
import pandas as pd
from datetime import datetime

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

def download_and_aggregate_tsm(
    drug_names,
    time_period,
    base_url="http://work.progames.top:3000",
    username="YLWZGJ2022",
    password="2026#ylwz",
    output_base_dir="TSM_Downloads",
    log_callback=print
):
    """
    Step A: 下载模块接口
    提供给总程序的调用接口。
    
    参数:
    - drug_names: string (以空格分隔) 或 list (药名列表)
    - time_period: string (如 '2022-2025' 或 '2022,2023') 或 list
    - base_url, username, password: 登录参数
    - output_base_dir: 存放所有下载数据的根目录
    - log_callback: 用于向主程序输出日志的回调函数，默认直接 print
    
    返回:
        (bool, str) -> (是否成功, 聚合后的Excel文件绝对路径，如果失败则为空字符串)
    """
    
    if isinstance(drug_names, str):
        drug_names = [d.strip().upper() for d in drug_names.split() if d.strip()]
    else:
        drug_names = [d.strip().upper() for d in drug_names if str(d).strip()]
        
    years_to_fetch = parse_years(time_period)
    
    if not drug_names:
        log_callback("[-] 未提取到任何有效药物名称，请检查输入。")
        return False, ""
        
    if not years_to_fetch:
        log_callback("[-] 年份格式解析失败，请使用如 '2023', '2022-2025', '2022,2023' 格式。")
        return False, ""
        
    # 使用当前脚本目录的绝对路径，防止改变工作目录导致缓存找不到
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.abspath(os.path.join(BASE_DIR, output_base_dir, "raw_files"))
    os.makedirs(output_dir, exist_ok=True)
    log_callback(f"[*] 已建/检查数据下载保存目录: {output_dir}")
    
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    })
    
    login_url = f"{base_url}/login"
    payload = {"username": username, "password": password}
    log_callback(f"[*] 正在尝试登录: {login_url}")
    
    try:
        resp = session.post(login_url, data=payload, timeout=30)
        if resp.status_code != 200:
            log_callback(f"[-] 登录失败，状态码: {resp.status_code}")
            return False, ""
        log_callback("[+] 登录成功。")
        # 短暂休息，避免频繁操作触发风控
        time.sleep(random.uniform(1, 2))
    except Exception as e:
        log_callback(f"[-] 登录发生异常: {e}")
        return False, ""
        
    # 指定爬取目标
    fields_to_try = ['mol']
    endpoints_to_try = ['MIDS']
    
    all_dfs = []
    
    log_callback(f"[*] 目标药物: {drug_names}")
    log_callback(f"[*] 目标年份: {years_to_fetch}")
    
    for drug_name in drug_names:
        for current_year in years_to_fetch:
            log_callback("-" * 50)
            log_callback(f">>> 开始拉取: {drug_name} | {current_year} <<<")
            success_for_year = False
            
            for ep in endpoints_to_try:
                target_route = f"/search{ep}"
                download2_url = f"{base_url}{target_route}/download2"
                download_url = f"{base_url}{target_route}/download"
                
                for field in fields_to_try:
                    filename = f"TSM_{current_year}_{drug_name}_{ep}.xlsx"
                    filepath = os.path.join(output_dir, filename)
                    
                    if os.path.exists(filepath) and os.path.getsize(filepath) > 1024:
                        log_callback(f"[+] 发现本地缓存: {filename}，跳过下载。")
                        try:
                            tmp_df = pd.read_excel(filepath)
                            if not tmp_df.empty:
                                if '检索药名' not in tmp_df.columns:
                                    tmp_df['检索药名'] = drug_name
                                all_dfs.append(tmp_df)
                            success_for_year = True
                            break
                        except Exception as e:
                            log_callback(f"[-] 读取本地缓存失败，将重新下载: {e}")
                            try: os.remove(filepath)
                            except: pass
                            
                    for attempt in range(1, 4):
                        log_callback(f"[*] 尝试次数 [{attempt}/3] - 端点 {ep} 字段 {field}")
                        payload = {field: drug_name, "year": current_year}
                        try:
                            resp = session.post(download2_url, data=payload, timeout=240)
                            if resp.status_code == 200:
                                result = resp.text.strip()
                                import json
                                is_json_err = False
                                try:
                                    jdata = json.loads(result)
                                    if str(jdata.get("ret")) == "0": is_json_err = True
                                except: pass
                                
                                # 判断是否是合理的UUID查询字符串
                                if not is_json_err and result and len(result) < 200 and "<html" not in result.lower():
                                    file_url = f"{download_url}?{result}"
                                    dll_resp = session.get(file_url, timeout=600)
                                    content_size = len(dll_resp.content)
                                    
                                    if dll_resp.status_code == 200 and content_size > 1024:
                                        filename = f"TSM_{current_year}_{drug_name}_{ep}.xlsx"
                                        filepath = os.path.join(output_dir, filename)
                                        with open(filepath, "wb") as f:
                                            f.write(dll_resp.content)
                                            
                                        log_callback(f"[+++] {current_year} {drug_name} 成功拉取! 保存至 {filename} ({content_size} byte)")
                                        success_for_year = True
                                        # 随机休眠 1-2 秒，保护账号防封号
                                        sleep_t = random.uniform(1, 2)
                                        log_callback(f"[*] 礼貌性等待 {sleep_t:.1f} 秒...")
                                        time.sleep(sleep_t)
                                        
                                        # 读取原始数据进行合并（暂不做业务清洗，仅堆叠以备后续模块处理）
                                        try:
                                            tmp_df = pd.read_excel(filepath)
                                            if tmp_df.empty:
                                                log_callback(f"    [-] 表格为空，跳过合并。")
                                                break
                                            
                                            if '检索药名' not in tmp_df.columns:
                                                tmp_df['检索药名'] = drug_name
                                                
                                            all_dfs.append(tmp_df)
                                            log_callback(f"    [+] 成功读取 {len(tmp_df)} 行原始数据准备合并。")
                                            
                                        except Exception as e:
                                            log_callback(f"    [-] 读取/处理 {filename} 失败: {e}")
                                            
                                        break
                                    else:
                                        log_callback(f"    [-] 文件异常/过小 ({content_size} bytes)")
                        except Exception as e:
                            log_callback(f"    [-] 请求异常: {e}")
                            
                    if success_for_year:
                        break
                if success_for_year:
                    break
                    
            if not success_for_year:
                log_callback(f"[-] 警告: 药物 {drug_name} 在年份 {current_year} 拉取失败！")

    # 所有下载任务结束后，主动注销登录保护账号安全
    try:
        logout_url = f"{base_url}/logout"
        session.get(logout_url, timeout=10)
        log_callback("[+] 已主动退出登录，账号安全。")
    except Exception as e:
        log_callback(f"[*] 退出登录时出现异常（不影响结果）: {e}")
    finally:
        session.close()

    if not all_dfs:
        log_callback("\n[-] 没有任何有效数据被下载和读取，任务结束。")
        return False, ""
        
    log_callback("\n" + "="*50)
    log_callback("[*] 开始聚合所有原始数据...")
    try:
        merged_df = pd.concat(all_dfs, ignore_index=True)
        # 固定缓存文件位置以供 Step 2 拾取，挂载到绝对路径
        cache_dir = os.path.join(BASE_DIR, "Cache")
        os.makedirs(cache_dir, exist_ok=True)
        merged_filepath = os.path.abspath(os.path.join(cache_dir, "step1_latest.xlsx"))
        
        merged_df.to_excel(merged_filepath, index=False)
        log_callback(f"[+++] 聚合 Excel 已保存至统一缓存路口: {merged_filepath} (共 {len(merged_df)} 行)")
        return True, merged_filepath
        
    except Exception as e:
        log_callback(f"[-] 保存聚合 Excel 失败: {e}")
        return False, ""

# 独立运行测试代码
if __name__ == "__main__":
    def my_logger(msg):
        # 模拟主程序的输出台
        print(f"[A模块] {msg}")
        
    # 可直接测试该接口
    print("开始测试 Step A 模块...")
    success, result_path = download_and_aggregate_tsm(
        drug_names="APIXABAN RIVAROXABAN",
        time_period="2022",
        output_base_dir="TSM_Downloads",
        log_callback=my_logger
    )
    
    if success:
        print(f"\n测试成功！下一步(Step B)可以使用该文件：\n{result_path}")
    else:
        print("\n测试失败。")
