import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import re
import os
import subprocess
import platform
import time
import math
from scipy import stats
from matplotlib.patches import Rectangle
import matplotlib.patches as mpatches
import matplotlib.lines as mlines

# ==========================================
# 1. 跨平台配置与样式
# ==========================================
system_name = platform.system()
if system_name == "Darwin": # MacOS
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'PingFang SC', 'Heiti TC', 'Arial']
else: # Windows
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial']

plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = '#F8F9FA'
plt.rcParams['axes.edgecolor'] = '#444444'
plt.rcParams['grid.color'] = '#E0E0E0'
plt.rcParams['grid.linestyle'] = '--'

# 形状映射 (Matplotlib Marker -> Unicode for Listbox)
MARKER_MAP = {
    'o': '●', # Circle
    's': '■', # Square
    '^': '▲', # Triangle Up
    'D': '◆', # Diamond
    'v': '▼', # Triangle Down
    '<': '◀', 
    '>': '▶',
    'p': '⬟', 
    '*': '★', 
    'h': '⬢', 
    'X': '✖',
    'd': '♦'
}
MARKERS_LIST = list(MARKER_MAP.keys())

# 配色
DISTINCT_COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#000080', '#800000', '#008080', '#808000', '#000000'
]

# 咨询风格配色
COLORS = {
    'primary': '#0052CC', 'secondary': '#36B37E', 'danger': '#FF5630', 
    'warning': '#FFAB00', 'neutral': '#6B778C', 
    'tiers': {
        'Tier 1': '#0052CC', 'Tier 2': '#36B37E', 
        'Tier 3': '#FFAB00', 'Tier 4': '#FF5630', 'Wait': '#6B778C'
    }
}

SALES_CACHE_FILE = "sales_data_cache.pkl"
EMA_CACHE_FILE = "ema_data_cache.pkl"

Target_Markets = [
    "UNITED KINGDOM", "GERMANY", "FRANCE", "ITALY", "SPAIN", 
    "NETHERLANDS", "BELGIUM", "SWEDEN", "SWITZERLAND", "AUSTRIA", 
    "POLAND", "PORTUGAL", "GREECE", "CZECH REPUBLIC", "HUNGARY",
    "ROMANIA", "IRELAND", "DENMARK", "FINLAND", "NORWAY"
]
EU5 = ["UNITED KINGDOM", "GERMANY", "FRANCE", "ITALY", "SPAIN"]

# ==========================================
# 2. 原研药知识库 (Originator Config)
# ==========================================
ORIGINATOR_CONFIG = {
    "APIXABAN": ["BRISTOL", "PFIZER", "SQUIBB"],
    "CITALOPRAM": ["LUNDBECK"],
    "RIVAROXABAN": ["BAYER"],
    "TADALAFIL": ["LILLY", "ICOS", "GLAXO"],
    "TICAGRELOR": ["ASTRAZENECA"],
    "TOBISILATE": ["ETHAMSYLATE"], 
}

# ==========================================
# 2.5 欧洲国家市场特征词典 (Country Profiles)
# ==========================================
COUNTRY_PROFILES = {
    "GERMANY": {
        "Type": "高度集约，保险主导 (Sickness Funds)",
        "Mechanism": "贴牌折扣协议 (Discount Agreements)",
        "Substitution": "药房强制替换为签订了折扣协议的仿制药",
        "Insight": "极度看重价格，赢者通吃格局(Tender驱动)。一旦失去折扣协议，本地份额几乎归零。新玩家需在两年一次的招标期抓住断供机会，通常是大玩家(Sandoz/Ratiopharm/Aliud/1A)的战场。"
    },
    "UNITED KINGDOM": {
        "Type": "自由定价但有利润封顶，药房主导",
        "Mechanism": "药房利润分成限制 (Drug Tariff)",
        "Substitution": "医生按INN(通用名)开处方，药店自行选择采购品牌",
        "Insight": "市场极度分散。药房会选择进货价最低、折扣最高的品牌以赚取最大差价。关系网在批发商(Wholesalers)层面最重要，而不是医生。"
    },
    "SWEDEN": {
        "Type": "单一支付方招标 (Monthly Tender)",
        "Mechanism": "每月大洗牌，赢者拿走几乎100%增量",
        "Substitution": "强制最高报销范围内最便宜的仿制药 (PvA)",
        "Insight": "极其极致的唯价格论市场。由于是定期重新竞标，一旦“Product of the Month”中标当月通吃，落标则销量暴跌。小公司往往靠极低价突袭。"
    },
    "FRANCE": {
        "Type": "内部参考定价，仿制药群体目标",
        "Mechanism": "参考组定价 (Tarif Forfaitaire de Responsabilité - TFR)",
        "Substitution": "药师有替代权，但患者需补差价如果选择贵药（近期政策更严）",
        "Insight": "法国人喜欢“国货”(如 Biogaran, Zentiva)。仿制药渗透率曾被认为较低，但政府正强制推高。首仿有先发优势，价格战不如德国惨烈，通常温和降价。"
    },
    "ITALY": {
        "Type": "区域化采购，原研偏好严重",
        "Mechanism": "大区公立医院招标 + 零售药房参考价",
        "Substitution": "药师可替换，但原研药品牌忠诚度极高",
        "Insight": "意大利是典型的“品牌仿制药”(Branded Generics)市场。患者和医生特别认牌子，导致纯靠低价很难立刻撬动份额，销售队伍的客情维护依然不可或缺。"
    },
    "SPAIN": {
        "Type": "快速价格联动大区制度",
        "Mechanism": "参考价格系统 (Sistema de Precios de Referencia)",
        "Substitution": "平价原则，所有药必须降到同组最低价才能报销",
        "Insight": "一旦有首仿以极低价上市，原研药和所有现有品牌必须在极短时间内跟着降价到同一水平。这导致大家最后拼的是药店的隐性折扣(Bonus)和供货稳定性。"
    },
    "POLAND": {
        "Type": "患者自付比例较高，品牌认知导向",
        "Mechanism": "报销清单与固定报销额",
        "Substitution": "药师必须告知最便宜的替代药，但患者有决定权",
        "Insight": "作为东欧大国，价格敏感度高，但患者对本地传统大厂（如 Polpharma, Adamed）极具信任。外来者往往需要找本地Top代理商做渠道包销。"
    },
    "NETHERLANDS": {
        "Type": "极端的保险公司偏好系统 (Preference Policy)",
        "Mechanism": "荷兰四大医保巨头掌握绝对话语权",
        "Substitution": "只报销保险公司指定的首选药物 (Preference Drug)",
        "Insight": "比德国还极端的“寡头采买”。一旦被最大的几家保险单子剔除，就无药可救。荷兰几乎消灭了所有的温和药房溢价空间。"
    }
}

# ==========================================
# 3. 核心工具函数
# ==========================================

def smart_price_parser(x):
    if pd.isna(x) or x == '': return 0.0
    if isinstance(x, (int, float)): return float(x)
    s = str(x).strip()
    s = re.sub(r'[^\d,.-]', '', s) 
    if not s: return 0.0
    try:
        last_dot = s.rfind('.')
        last_comma = s.rfind(',')
        if last_dot > last_comma: s = s.replace(',', '')
        elif last_comma > last_dot: s = s.replace('.', '').replace(',', '.')
        return float(s)
    except: return 0.0

def copy_to_clipboard(filepath):
    sys_type = platform.system()
    abs_path = os.path.abspath(filepath)
    try:
        if sys_type == "Darwin":
            script = f'set the clipboard to (read (POSIX file "{abs_path}") as JPEG picture)'
            subprocess.run(["osascript", "-e", script])
            return True
        elif sys_type == "Windows":
            cmd = f"powershell -c \"Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile('{abs_path}'))\""
            subprocess.run(cmd, shell=True)
            return True
    except: pass
    return False

# ==========================================
# 4. 主程序类
# ==========================================

class PharmaAnalyzerProV24:
    def __init__(self, root):
        self.root = root
        self.root.title(f"CMO Strategy V24 ({platform.system()}): Final Stable")
        self.root.geometry("1400x950")
        
        self.df_sales_raw = None
        self.df_sales_clean = None
        self.df_ema = None
        
        self.selected_companies_for_originator = [] 
        self.current_countries = [] 
        self.highlighted_country = None 
        
        self.pie_batch_index = 0
        self.pie_batch_size = 6
        
        self._init_gui()
        self._check_caches()

    def _init_gui(self):
        control_frame = ttk.Frame(self.root, padding="5")
        control_frame.pack(fill=tk.X)
        
        # 1. Sources
        f1 = ttk.LabelFrame(control_frame, text="1. Sources")
        f1.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        self.btn_sales = ttk.Button(f1, text="Load Sales", command=self.load_sales_file)
        self.btn_sales.pack(padx=5, pady=2)
        self.btn_ema = ttk.Button(f1, text="Load EMA", command=self.load_ema_data)
        self.btn_ema.pack(padx=5, pady=2)
        ttk.Button(f1, text="Refresh DB", command=self.force_reload).pack(padx=5, pady=2)
        
        # 2. Drug
        f2 = ttk.LabelFrame(control_frame, text="2. Drug Selection")
        f2.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        self.combo_drugs = ttk.Combobox(f2, state="readonly", width=20)
        self.combo_drugs.pack(padx=5, pady=5)
        self.combo_drugs.bind("<<ComboboxSelected>>", self.on_drug_selected)
        
        # 3. Filter
        f3 = ttk.LabelFrame(control_frame, text="3. Originator Config")
        f3.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        self.btn_companies = ttk.Button(f3, text="Verify Originators", command=self.open_company_selector, state=tk.DISABLED)
        self.btn_companies.pack(padx=5, pady=5)
        self.lbl_originator_status = ttk.Label(f3, text="Auto-Detect: Off", foreground="gray")
        self.lbl_originator_status.pack()
        
        # 4. Analysis
        f5 = ttk.LabelFrame(control_frame, text="4. Strategy Engine")
        f5.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        self.btn_analyze = ttk.Button(f5, text="Run Analysis (Single Drug)", command=self.run_auto_analysis, state=tk.DISABLED)
        self.btn_analyze.pack(padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="Checking caches...")
        ttk.Label(control_frame, textvariable=self.status_var, foreground="blue").pack(side=tk.LEFT, padx=20)

        # Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_summary = ttk.Frame(self.notebook)
        self.tab_trend = ttk.Frame(self.notebook)
        self.tab_pen = ttk.Frame(self.notebook)
        self.tab_pie = ttk.Frame(self.notebook)
        self.tab_matrix = ttk.Frame(self.notebook)
        self.tab_ma = ttk.Frame(self.notebook)
        self.tab_predict = ttk.Frame(self.notebook) 
        
        self.notebook.add(self.tab_summary, text="0. 战略决策 (Global)")
        self.notebook.add(self.tab_trend, text="1. 趋势分析 (Trend)")
        self.notebook.add(self.tab_pen, text="2. 原研渗透率 (Penetration)")
        self.notebook.add(self.tab_pie, text="3. 仿制药格局 (Structure)")
        self.notebook.add(self.tab_matrix, text="4. 机会矩阵 (Matrix)")
        self.notebook.add(self.tab_ma, text="5. 潜在威胁 (Threats)")
        self.notebook.add(self.tab_predict, text="6. 智能预测 (Prediction)")

    # ==========================================
    # 4. 数据处理 (Data Processing)
    # ==========================================

    def _check_caches(self):
        if os.path.exists(EMA_CACHE_FILE):
            try:
                self.df_ema = pd.read_pickle(EMA_CACHE_FILE)
                self.btn_ema.config(text="EMA Cached ✓", state=tk.DISABLED)
            except: pass
        if os.path.exists(SALES_CACHE_FILE):
            try:
                self.df_sales_raw = pd.read_pickle(SALES_CACHE_FILE)
                self.btn_sales.config(text="Sales Cached ✓", state=tk.DISABLED)
                self._populate_drugs()
                self.status_var.set("Cached data loaded.")
            except: pass

    def force_reload(self):
        if os.path.exists(EMA_CACHE_FILE): os.remove(EMA_CACHE_FILE)
        if os.path.exists("ema_data_cache.xlsx"): os.remove("ema_data_cache.xlsx")
        if os.path.exists(SALES_CACHE_FILE): os.remove(SALES_CACHE_FILE)
        if os.path.exists("sales_raw_cache.xlsx"): os.remove("sales_raw_cache.xlsx")
        self.df_sales_raw = None
        self.df_ema = None
        self.btn_sales.config(text="Load Sales", state=tk.NORMAL)
        self.btn_ema.config(text="Load EMA", state=tk.NORMAL)
        self.combo_drugs.set('')
        self.combo_drugs['values'] = []
        self.status_var.set("Cache cleared.")

    def load_sales_file(self):
        file_path = filedialog.askopenfilename()
        if not file_path: return
        self.status_var.set("Reading Excel...")
        self.root.update()
        try:
            xl = pd.ExcelFile(file_path)
            sheet = next((s for s in xl.sheet_names if "原始" in s), xl.sheet_names[0])
            self.df_sales_raw = pd.read_excel(file_path, sheet_name=sheet, dtype=str)
            self.df_sales_raw.columns = [str(c).strip() for c in self.df_sales_raw.columns]
            if '通用名单' not in self.df_sales_raw.columns and len(self.df_sales_raw.columns)>4:
                 self.df_sales_raw.rename(columns={self.df_sales_raw.columns[4]: '通用名单'}, inplace=True)
            
            self.df_sales_raw.to_pickle(SALES_CACHE_FILE)
            self._populate_drugs()
            self.btn_sales.config(text="Sales Cached ✓", state=tk.DISABLED)
            self.status_var.set(f"Loaded {len(self.df_sales_raw)} rows.")
        except Exception as e: messagebox.showerror("Error", str(e))

    def _populate_drugs(self):
        if self.df_sales_raw is not None:
            drugs = sorted(self.df_sales_raw['通用名单'].dropna().unique().tolist())
            self.combo_drugs['values'] = drugs
            if drugs: self.combo_drugs.current(0)

    def load_ema_data(self):
        import requests
        url = "https://www.ema.europa.eu/sites/default/files/Medicines_output_european_public_assessment_reports.xlsx"
        local_excel = "ema_data_cache.xlsx"
        
        self.status_var.set("Downloading EMA Database...")
        self.root.update()
        
        try:
            if not os.path.exists(local_excel):
                resp = requests.get(url, timeout=60, verify=False)
                if resp.status_code == 200:
                    with open(local_excel, "wb") as f:
                        f.write(resp.content)
                else:
                    messagebox.showerror("Error", f"Failed to download EMA data. Status code: {resp.status_code}")
                    self.status_var.set("EMA Download Failed.")
                    return
            
            self.status_var.set("Processing EMA Database...")
            self.root.update()
            
            df = pd.read_excel(local_excel, header=19, usecols=[0, 1, 3, 4])
            df.columns = ['Product', 'Substance', 'Country', 'MAH']
            df['Country'] = df['Country'].astype(str).str.upper().str.strip()
            df['Substance'] = df['Substance'].astype(str).str.upper()
            df.to_pickle(EMA_CACHE_FILE)
            self.df_ema = df
            self.btn_ema.config(text="EMA Cached ✓", state=tk.DISABLED)
            self.status_var.set("EMA Ready.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed processing EMA data: {str(e)}")
            self.status_var.set("EMA Processing Failed.")

    def on_drug_selected(self, event=None):
        drug = self.combo_drugs.get()
        if not drug or self.df_sales_raw is None: return
        
        self.status_var.set("Cleaning Data...")
        self.root.update()
        
        # 1. Filter Drug
        df = self.df_sales_raw[self.df_sales_raw['通用名单'] == drug].copy()
        
        # 2. Map Columns (Defensive)
        if '国家' in df.columns: col_country = df['国家']
        elif 'Country' in df.columns: col_country = df['Country']
        else: 
            messagebox.showerror("Error", "找不到[国家]列")
            return
            
        col_year = df['年份'] if '年份' in df.columns else df.get('Year', pd.Series(['']*len(df)))
        col_comp = df['集团/企业'] if '集团/企业' in df.columns else df.get('Company', df.get('生产企业', pd.Series(['Unknown']*len(df))))

        if '最小单包装销售数量 粒' in df.columns: col_vol = df['最小单包装销售数量 粒']
        elif 'Volume' in df.columns: col_vol = df['Volume']
        else: col_vol = pd.Series([0]*len(df), index=df.index)
        
        if '销售额' in df.columns: col_sales = df['销售额']
        elif 'Sales' in df.columns: col_sales = df['Sales']
        else: col_sales = pd.Series([0]*len(df), index=df.index)

        if '每粒单价' in df.columns: col_price = df['每粒单价']
        else: col_price = pd.Series([0]*len(df), index=df.index)

        # 3. Clean
        df_clean = pd.DataFrame()
        df_clean['国家'] = col_country.astype(str).str.upper().str.strip()
        df_clean['年份'] = col_year.astype(str).str.strip()
        df_clean['集团/企业'] = col_comp.astype(str).str.strip()
        df_clean['Volume_Clean'] = col_vol.apply(smart_price_parser)
        df_clean['Sales_Clean'] = col_sales.apply(smart_price_parser)
        df_clean['Price_Clean'] = col_price.apply(smart_price_parser)
        df_clean['Factory_Price'] = df_clean['Price_Clean'] * 0.3
        
        df_clean = df_clean[df_clean['国家'].isin(Target_Markets)]
        self.df_sales_clean = df_clean[df_clean['Sales_Clean'] > 0].copy()
        
        if self.df_sales_clean.empty:
            self.status_var.set("No valid sales data.")
            return

        # 4. Auto-Detect Originator
        self.selected_companies_for_originator = []
        drug_upper = drug.upper()
        found_config = False
        
        for key, comps in ORIGINATOR_CONFIG.items():
            if key in drug_upper:
                avail_comps = self.df_sales_clean['集团/企业'].unique()
                for ac in avail_comps:
                    for target in comps:
                        if target.upper() in ac.upper():
                            self.selected_companies_for_originator.append(ac)
                            found_config = True
        
        if found_config:
            self.lbl_originator_status.config(text=f"Config Match: {len(self.selected_companies_for_originator)} Found", foreground="green")
        else:
            self.lbl_originator_status.config(text="Config Miss: Please select manually", foreground="orange")

        self.btn_companies['state'] = tk.NORMAL
        self.btn_analyze['state'] = tk.NORMAL
        self.status_var.set(f"Ready. {len(self.df_sales_clean)} rows.")

    def open_company_selector(self):
        if self.df_sales_clean is None: return
        win = Toplevel(self.root)
        win.title("Step 3: Identify Originator")
        win.geometry("500x600")
        
        ttk.Label(win, text="Check companies to set as 'Originator'.", font=('Arial', 10, 'bold')).pack(pady=5)
        
        stats = self.df_sales_clean.groupby('集团/企业').agg({
            'Factory_Price': 'max', 
            'Sales_Clean': 'sum'
        }).sort_values('Factory_Price', ascending=False)
        
        frm = ttk.Frame(win); frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        canvas = tk.Canvas(frm); sb = ttk.Scrollbar(frm, orient="vertical", command=canvas.yview)
        sc = ttk.Frame(canvas)
        sc.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=sc, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.check_vars = {}
        for idx, (comp, row) in enumerate(stats.iterrows()):
            price = row['Factory_Price']
            is_checked = comp in self.selected_companies_for_originator
            var = tk.BooleanVar(value=is_checked)
            self.check_vars[comp] = var
            bg = "#FFEBEE" if idx < 3 else None
            txt = f"{comp} | Price:{price:.2f}"
            tk.Checkbutton(sc, text=txt, variable=var, bg=bg, anchor='w').pack(fill='x', padx=2, pady=1)
        
        def confirm():
            self.selected_companies_for_originator = [c for c,v in self.check_vars.items() if v.get()]
            win.destroy()
        ttk.Button(win, text="Confirm", command=confirm).pack(pady=10)

    def run_auto_analysis(self):
        if self.df_sales_clean is None: return
        c_stats = self.df_sales_clean.groupby('国家')['Sales_Clean'].sum().sort_values(ascending=False)
        self.current_countries = c_stats.head(15).index.tolist()
        self.pie_batch_index = 0
        self.highlighted_country = None
        self.generate_all_charts()

    # ==========================================
    # 5. 分析主逻辑
    # ==========================================

    def generate_all_charts(self):
        self.status_var.set("Analyzing...")
        df = self.df_sales_clean.copy()
        
        # 自动预测因子：不需手动，根据历史趋势与RunRate自动混合
        # 此处不直接修改 df 的 2025 值，而在绘图函数中动态计算
        
        # 准备数据：仿制药专用 (剔除原研)
        df_gen = df[~df['集团/企业'].isin(self.selected_companies_for_originator)].copy()
        
        # 0. Summary (with Global Scan Button)
        self.render_chart(self.tab_summary, "Executive Summary", lambda f: self.draw_summary_table(f, df))
        
        # 1. Trend (YTD Actuals) - Interactive Legend
        self.render_chart_with_legend(self.tab_trend, "Trend", lambda f: self.draw_trend(f, df))
        
        # 2. Penetration (Overview)
        self.render_chart(self.tab_pen, "Penetration", lambda f: self.draw_penetration(f, df))
        
        # 3. Pie (Generics Only)
        self.render_pie_tab(df_gen)
        
        # 4. Matrix (Full Data)
        self.render_matrix_tab(self.tab_matrix, df)
        
        # 5. Threats
        self.render_chart(self.tab_ma, "Threats", lambda f: self.draw_ma_threats(f, df))
        
        # 6. Prediction (Bayesian)
        self.render_prediction_tab(self.tab_predict, df)
        
        self.status_var.set("Done.")

    def render_chart(self, parent, name, func):
        for w in parent.winfo_children(): w.destroy()
        bf = ttk.Frame(parent); bf.pack(fill=tk.X, padx=10, pady=5)
        
        # Algo Text
        algo_frame = ttk.Frame(parent, relief="sunken", borderwidth=1)
        algo_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        algo_lbl = ttk.Label(algo_frame, text=self.get_algo_explanation(name), foreground="#555", font=('Arial', 9, 'italic'), padding=5)
        algo_lbl.pack(anchor='w')

        tf = ttk.Frame(parent); tf.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        txt = tk.Text(tf, height=10, font=('Arial', 10), bg="#F5F5F5", relief="flat", wrap=tk.WORD)
        txt.pack(fill=tk.BOTH)
        
        fig = plt.Figure(figsize=(10, 6), dpi=100)
        res_txt = func(fig)
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        txt.insert(tk.END, res_txt)
        
        def cp():
            path = os.path.abspath(f"temp_{name}.jpg")
            fig.savefig(path, bbox_inches='tight', dpi=150)
            if copy_to_clipboard(path): messagebox.showinfo("OK", "Copied!")
        ttk.Button(bf, text="📋 Copy", command=cp).pack(side=tk.RIGHT)
        
        if name == "Executive Summary":
            ttk.Button(bf, text="🌍 Run Global Scan (All Drugs)", command=self.run_global_scan).pack(side=tk.RIGHT, padx=10)
            
        ttk.Label(bf, text=f"{name} Analysis", font=('bold')).pack(side=tk.LEFT)

    def get_algo_explanation(self, name):
        """返回各个图表的算法白皮书"""
        explanations = {
            "Executive Summary": "【算法】基于波士顿矩阵逻辑，将市场分为4个Tier。一级筛选看销量(Volume)，二级筛选看价格(Price)。Tier 1 = 双优。",
            "Trend": "【逻辑】展示原始销售数据(YTD)。数据未经线性处理，真实反映季节性波动和价格衰减趋势。",
            "Penetration": "【逻辑】计算原研药(Originator)与仿制药的销售额占比。原研占比越高，理论替代空间越大(Tier 4)。",
            "Matrix": "【逻辑】气泡大小=总销售额。X轴=出厂价，Y轴=销量。辅助线为中位数(Median)。不同形状代表不同Tier。",
            "Threats": "【数据源】左侧=销售报表活跃公司数；右侧=EMA数据库批准但未销售公司数(Ghost MAHs)。",
            "Prediction": "【算法】气泡矩阵模型。X轴=出厂单价，Y轴=销量，气泡大小=总销售额。用于直观筛选「量大价优」的好市场。",
            "Structure": "【逻辑】HHI指数分析。仅统计仿制药内部竞争，剔除原研药干扰。"
        }
        return explanations.get(name, "")

    def run_global_scan(self):
        """遍历所有药物，统计最佳国家"""
        if self.df_sales_raw is None: return
        
        msg = messagebox.showinfo("Global Scan", "Scanning all drugs... This may take a minute.")
        all_drugs = self.df_sales_raw['通用名单'].unique()
        country_scores = {} # Country -> Score
        
        for d in all_drugs:
            sub = self.df_sales_raw[self.df_sales_raw['通用名单']==d]
            if '国家' in sub.columns: 
                ctry = sub['国家'].astype(str).str.upper().str.strip()
                sub = sub[ctry.isin(Target_Markets)]
                if sub.empty: continue
                
                if '最小单包装销售数量 粒' in sub.columns: vol = sub['最小单包装销售数量 粒'].apply(smart_price_parser)
                else: continue
                if '销售额' in sub.columns: val = sub['销售额'].apply(smart_price_parser)
                else: continue
                
                agg = pd.DataFrame({'C': ctry[sub.index], 'Vol': vol, 'Val': val})
                agg = agg.groupby('C').sum()
                
                med_v = agg['Vol'].median()
                med_p = (agg['Val']/agg['Vol']).median()
                
                for c, r in agg.iterrows():
                    p = r['Val']/r['Vol'] if r['Vol']>0 else 0
                    if r['Vol'] > med_v and p > med_p: # Tier 1
                        country_scores[c] = country_scores.get(c, 0) + 1
        
        res = sorted(country_scores.items(), key=lambda x:x[1], reverse=True)
        top_txt = "\n".join([f"{k}: {v} Tier-1 Hits" for k,v in res[:10]])
        messagebox.showinfo("Global Portfolio Result", f"Top Strategic Markets across ALL drugs:\n\n{top_txt}")

    # --- Trend Legend Interactive ---
    def render_chart_with_legend(self, parent, name, func):
        for w in parent.winfo_children(): w.destroy()
        bf = ttk.Frame(parent); bf.pack(fill=tk.X, padx=10, pady=5)
        
        main_frame = ttk.Frame(parent)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Algo Text
        algo_frame = ttk.Frame(parent, relief="sunken", borderwidth=1)
        algo_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        algo_lbl = ttk.Label(algo_frame, text=self.get_algo_explanation(name), foreground="#555", font=('Arial', 9, 'italic'), padding=5)
        algo_lbl.pack(anchor='w')

        tf = ttk.Frame(parent); tf.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        txt = tk.Text(tf, height=8, font=('Arial', 10), bg="#F5F5F5", relief="flat", wrap=tk.WORD)
        txt.pack(fill=tk.BOTH)
        
        fig = plt.Figure(figsize=(10, 6), dpi=100)
        res_txt, leg_map = func(fig) # Draw Initial
        
        canvas = FigureCanvasTkAgg(fig, master=main_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        txt.insert(tk.END, res_txt)
        
        self._current_leg_map = leg_map
        
        # Interaction
        def on_pick(event):
            legline = event.artist
            # Only handle if artist is in our legend map
            if legline in self._current_leg_map:
                country_name = self._current_leg_map[legline]
                
                # Toggle highlight
                if country_name == self.highlighted_country:
                    self.highlighted_country = None
                else:
                    self.highlighted_country = country_name
                    # Global Profile Hook
                    if country_name.upper() in COUNTRY_PROFILES:
                        self.show_country_profile(country_name.upper())
                        
                fig.clf()
                _, new_leg_map = func(fig)
                self._current_leg_map = new_leg_map
                canvas.draw()
                
        canvas.mpl_connect('pick_event', on_pick)
        
        def reset_highlight():
            self.highlighted_country = None
            fig.clf()
            _, new_leg_map = func(fig)
            self._current_leg_map = new_leg_map
            canvas.draw()
            
        ttk.Button(bf, text="↺ Reset Highlight", command=reset_highlight).pack(side=tk.RIGHT, padx=5)
        
        def cp():
            export_fig = plt.Figure(figsize=(14, 7), dpi=150)
            prev_hl = self.highlighted_country
            self.highlighted_country = None # Reset for export
            _t, _ = func(export_fig)
            self.highlighted_country = prev_hl
            path = os.path.abspath(f"temp_{name}.jpg")
            export_fig.savefig(path, bbox_inches='tight')
            if copy_to_clipboard(path): messagebox.showinfo("OK", "Copied!")
            
        ttk.Button(bf, text="📋 Copy", command=cp).pack(side=tk.RIGHT)
        ttk.Label(bf, text=f"{name} Analysis (Click Legend to Highlight)", font=('bold')).pack(side=tk.LEFT)

    def render_pie_tab(self, df):
        parent = self.tab_pie
        for w in parent.winfo_children(): w.destroy()
        bf = ttk.Frame(parent); bf.pack(fill=tk.X, padx=10, pady=5)
        
        algo_frame = ttk.Frame(parent, relief="sunken", borderwidth=1)
        algo_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        algo_lbl = ttk.Label(algo_frame, text=self.get_algo_explanation("Structure"), foreground="#555", font=('Arial', 9, 'italic'), padding=5)
        algo_lbl.pack(anchor='w')

        tf = ttk.Frame(parent); tf.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        txt = tk.Text(tf, height=10, font=('Arial', 10), bg="#F5F5F5", relief="flat", wrap=tk.WORD)
        txt.pack(fill=tk.BOTH)
        
        # Scrollable area for matplotlib
        scroll_frame = ttk.Frame(parent)
        scroll_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        canvas_scroll = tk.Canvas(scroll_frame)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas_scroll.yview)
        scrollable_container = ttk.Frame(canvas_scroll)
        
        scrollable_container.bind(
            "<Configure>",
            lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
        )
        
        canvas_scroll.create_window((0, 0), window=scrollable_container, anchor="nw", width=1200)
        canvas_scroll.configure(yscrollcommand=scrollbar.set)
        
        canvas_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        fig = plt.Figure(figsize=(12, 6), dpi=100)
        total_c = len(self.current_countries)
        batch_countries = self.current_countries
        
        res_txt = self.draw_pie_batch(fig, df, batch_countries)
        canvas = FigureCanvasTkAgg(fig, master=scrollable_container)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Ensure scrollable container fits width of canvas
        def _on_canvas_configure(event):
            canvas_scroll.itemconfig(canvas_scroll.find_withtag("all")[0], width=event.width)
        canvas_scroll.bind("<Configure>", _on_canvas_configure)
        
        txt.insert(tk.END, res_txt)
        
        def on_pie_pick(event):
            try:
                # Event on the axes
                ax = event.inaxes
                if ax and ax.get_title():
                    # The title might have \n(HHI: ...), so we split and take the first part
                    title = ax.get_title().split('\n')[0].strip()
                    if title.upper() in COUNTRY_PROFILES:
                        self.show_country_profile(title.upper())
            except Exception:
                pass
                
        canvas.mpl_connect('button_press_event', on_pie_pick)
        
        def cp():
            path = os.path.abspath(f"temp_pie.jpg")
            fig.savefig(path, bbox_inches='tight', dpi=150)
            copy_to_clipboard(path)

        ttk.Button(bf, text="📋 Copy", command=cp).pack(side=tk.RIGHT)

    def render_matrix_tab(self, parent, df):
        for w in parent.winfo_children(): w.destroy()
        
        ctrl_f = ttk.Frame(parent)
        ctrl_f.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(ctrl_f, text="X轴 (X-Axis):").pack(side=tk.LEFT, padx=5)
        x_combo = ttk.Combobox(ctrl_f, values=["市场绝对规模", "原研替代空间", "价格降幅"], state="readonly", width=15)
        x_combo.current(0)
        x_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(ctrl_f, text="Y轴 (Y-Axis):").pack(side=tk.LEFT, padx=5)
        y_combo = ttk.Combobox(ctrl_f, values=["市场增速", "竞争者数量", "集中度HHI"], state="readonly", width=15)
        y_combo.current(0)
        y_combo.pack(side=tk.LEFT, padx=5)
        
        bf = ttk.Frame(parent); bf.pack(fill=tk.X, padx=10, pady=5)
        algo_frame = ttk.Frame(parent, relief="sunken", borderwidth=1)
        algo_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        algo_lbl = ttk.Label(algo_frame, text=self.get_algo_explanation("Matrix"), foreground="#555", font=('Arial', 9, 'italic'), padding=5)
        algo_lbl.pack(anchor='w')

        tf = ttk.Frame(parent); tf.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        txt = tk.Text(tf, height=8, font=('Arial', 10), bg="#F5F5F5", relief="flat", wrap=tk.WORD)
        txt.pack(fill=tk.BOTH)
        
        fig = plt.Figure(figsize=(10, 6), dpi=100)
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        def update_matrix(event=None):
            fig.clf()
            x_val = x_combo.get()
            y_val = y_combo.get()
            res_txt = self.draw_dynamic_matrix(fig, df, x_val, y_val)
            canvas.draw()
            txt.delete("1.0", tk.END)
            txt.insert(tk.END, res_txt)
            
        x_combo.bind("<<ComboboxSelected>>", update_matrix)
        y_combo.bind("<<ComboboxSelected>>", update_matrix)
        
        update_matrix()

        def on_pick(event):
            ind = event.ind[0]
            if hasattr(self, '_matrix_countries') and ind < len(self._matrix_countries):
                country = self._matrix_countries[ind]
                if country in self.current_countries:
                    idx = self.current_countries.index(country)
                    self.pie_batch_index = idx // self.pie_batch_size
                    
                    # 弹出国家百科卡片
                    if country.upper() in COUNTRY_PROFILES:
                        self.show_country_profile(country.upper())
                    
                    self.render_pie_tab(df[~df['集团/企业'].isin(self.selected_companies_for_originator)])
                    self.notebook.select(self.tab_pie)
        
        canvas.mpl_connect('pick_event', on_pick)
        
        def cp():
            path = os.path.abspath(f"temp_matrix.jpg")
            fig.savefig(path, bbox_inches='tight', dpi=150)
            copy_to_clipboard(path)
            
        ttk.Button(bf, text="📋 Copy", command=cp).pack(side=tk.RIGHT)
        ttk.Label(bf, text="Matrix Analysis", font=('bold')).pack(side=tk.LEFT)

    def show_country_profile(self, country):
        """显示国家市场档案弹窗"""
        profile = COUNTRY_PROFILES.get(country)
        if hasattr(self, 'profile_win') and self.profile_win.winfo_exists():
            self.profile_win.destroy()
            
        win = Toplevel(self.root)
        self.profile_win = win
        win.title(f"📖 欧洲准入战略局: {country} Market Profile")
        win.geometry("550x300")
        # 使其保持在顶部且不模态，这样用户可以边看边操作
        win.attributes('-topmost', 1) 
        
        main_f = ttk.Frame(win, padding=15)
        main_f.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(main_f, text=f"{country} 市场微观特征", font=('Arial', 14, 'bold'), foreground=COLORS['primary']).pack(anchor='w', pady=(0,10))
        
        # Content
        text = tk.Text(main_f, wrap=tk.WORD, font=('Arial', 11), bg="#FAFAFA", relief="flat")
        text.pack(fill=tk.BOTH, expand=True)
        
        content = f"📌 核心机制 (Archetype):\n{profile['Type']}\n\n"
        content += f"💰 报销与定价 (Reimbursement):\n{profile['Mechanism']}\n\n"
        content += f"🔄 替代政策 (Substitution):\n{profile['Substitution']}\n\n"
        content += f"⚠️ 战略洞察 (Key Insight):\n{profile['Insight']}"
        
        text.insert("1.0", content)
        text.config(state=tk.DISABLED) # Read-only
        
        ttk.Button(main_f, text="关闭", command=win.destroy).pack(pady=(10,0))

    def render_prediction_tab(self, parent, df):
        for w in parent.winfo_children(): w.destroy()
        
        bf = ttk.Frame(parent); bf.pack(fill=tk.X, padx=10, pady=5)
        algo_frame = ttk.Frame(parent, relief="sunken", borderwidth=1)
        algo_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=2)
        algo_lbl = ttk.Label(algo_frame, text=self.get_algo_explanation("Prediction"), foreground="#555", font=('Arial', 9, 'italic'), padding=5)
        algo_lbl.pack(anchor='w')

        tf = ttk.Frame(parent); tf.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        txt = tk.Text(tf, height=8, font=('Arial', 10), bg="#F5F5F5", relief="flat", wrap=tk.WORD)
        txt.pack(fill=tk.BOTH)
        
        fig = plt.Figure(figsize=(10, 6), dpi=100)
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        def update_pred(event=None):
            fig.clf()
            res_txt = self.draw_prediction(fig, df)
            canvas.draw()
            txt.delete("1.0", tk.END)
            txt.insert(tk.END, res_txt)
            
        update_pred()
        
        def cp():
            path = os.path.abspath(f"temp_predict.jpg")
            fig.savefig(path, bbox_inches='tight', dpi=150)
            copy_to_clipboard(path)
            
        ttk.Button(bf, text="📋 Copy", command=cp).pack(side=tk.RIGHT)
        ttk.Label(bf, text="Market Prediction Bubble Matrix", font=('bold')).pack(side=tk.LEFT)

    # ==========================================
    # 6. 绘图与分析逻辑 (Logic)
    # ==========================================

    def draw_summary_table(self, fig, df):
        ax = fig.add_subplot(111)
        ax.axis('off')
        
        data = []
        for c in self.current_countries:
            sub = df[df['国家']==c]
            total_vol = sub['Volume_Clean'].sum()
            total_sales = sub['Sales_Clean'].sum()
            avg_price = (total_sales / total_vol * 0.3) if total_vol > 0 else 0
            
            orig = sub[sub['集团/企业'].isin(self.selected_companies_for_originator)]['Sales_Clean'].sum()
            orig_share = orig / total_sales if total_sales > 0 else 0
            data.append({'C': c, 'Vol': total_vol, 'Price': avg_price, 'Orig': orig_share})
            
        df_metrics = pd.DataFrame(data)
        if df_metrics.empty: return "无数据"
        
        med_vol = df_metrics['Vol'].median()
        med_price = df_metrics['Price'].median()
        
        def get_tier(row):
            if row['Orig'] > 0.8: return "Tier 4"
            if row['Vol'] > med_vol and row['Price'] > med_price: return "Tier 1"
            if row['Vol'] > med_vol: return "Tier 2"
            if row['Price'] > med_price: return "Tier 3"
            return "Wait"
            
        df_metrics['Tier'] = df_metrics.apply(get_tier, axis=1)
        df_metrics = df_metrics.sort_values(['Vol', 'Price'], ascending=False)
        
        # 移除表头
        table_data = [] # No Header
        cell_colors = []
        
        for _, r in df_metrics.iterrows():
            row_c = [r['C'], r['Tier'], f"{r['Vol']/1e6:.1f}M", f"{r['Price']:.2f}", f"{r['Orig']:.0%}"]
            table_data.append(row_c)
            base_c = COLORS['tiers'].get(r['Tier'], "#FFFFFF")
            # Transparent colors for table
            if r['Tier'] == "Tier 1": base_c = "#E3F2FD" 
            if r['Tier'] == "Tier 2": base_c = "#E8F5E9"
            if r['Tier'] == "Tier 3": base_c = "#FFF3E0"
            if r['Tier'] == "Tier 4": base_c = "#FFEBEE"
            if r['Tier'] == "Wait": base_c = "#F5F5F5"
            cell_colors.append([base_c]*5)
            
        # 调整表格样式
        the_table = ax.table(cellText=table_data, cellColours=cell_colors, loc='center', cellLoc='left', colWidths=[0.25, 0.2, 0.2, 0.15, 0.2])
        the_table.auto_set_font_size(False); the_table.set_fontsize(10); the_table.scale(1, 1.8)
        
        text = "【战略决策建议】\n"
        t1 = df_metrics[df_metrics['Tier'].str.contains("Tier 1")]['C'].tolist()
        t4 = df_metrics[df_metrics['Tier'].str.contains("Tier 4")]['C'].tolist()
        text += f"• 核心首选 (Tier 1): {', '.join(t1)} (量大价优)\n"
        if t4: text += f"• 战略机会 (Tier 4): {', '.join(t4)} (原研垄断)\n\n"
        text += "📊 【数据口径】: 抽取最新周期的整体销量(Volume)和预估出厂价(Price)。原研占比反映了理论替代空间。\n"
        text += f"💡 【汇报话术】: “老板，基于量价双优的原则，{', '.join(t1[:2])}等国是我们的Tier-1核心票仓；而像{', '.join(t4[:2])}虽然目前量不大，但是原研药占据绝对主导，缺乏仿制药竞争，存在切入蓝海的巨大降维打击空间。”"
        return text

    def draw_trend(self, fig, df):
        ax1 = fig.add_subplot(211); ax2 = fig.add_subplot(212, sharex=ax1)
        df_sub = df[df['国家'].isin(self.current_countries)]
        trend = df_sub.groupby(['国家','年份']).apply(lambda x: pd.Series({
            'Vol': x['Volume_Clean'].sum(), 
            'Price': (x['Sales_Clean'].sum()/x['Volume_Clean'].sum()*0.3) if x['Volume_Clean'].sum()>0 else 0
        })).reset_index()
        trend['年份'] = trend['年份'].astype(str).str.replace(r'\.0$', '', regex=True)
        
        labels1 = []
        
        for i, c in enumerate(self.current_countries):
            sub = trend[trend['国家']==c].sort_values('年份')
            if sub.empty: continue
            yrs = sub['年份'].tolist(); vols = sub['Vol'].tolist(); prs = sub['Price'].tolist()
            
            # Calculate CAGR for Vol
            cagr_label = ""
            if len(vols) >= 2 and vols[0] > 0:
                n = len(vols) - 1
                cagr = (vols[-1] / vols[0]) ** (1/n) - 1
                if cagr > 0.15: stage = "快速放量期"
                elif cagr > 0: stage = "导入/平台期"
                else: stage = "衰退期"
                cagr_label = f" ({stage} {cagr:.0%})"

            color = DISTINCT_COLORS[i % len(DISTINCT_COLORS)]
            marker = MARKERS_LIST[i % len(MARKERS_LIST)]
            
            # Highlight logic
            alpha = 1.0
            lw = 2
            if self.highlighted_country and c != self.highlighted_country:
                alpha = 0.1 
                lw = 1
            elif self.highlighted_country == c:
                lw = 4
            
            kw = {'color': color, 'marker': marker, 'markersize': 6, 'linewidth': lw, 'alpha': alpha}
            
            ax1.plot(yrs, vols, label=c + cagr_label, **kw)
            ax2.plot(yrs, prs, label=c + cagr_label, **kw)
            labels1.append(c)
        
        ax1.set_title("Volume Trend (Raw YTD)"); ax2.set_title("Price Trend")
        
        # Draw legend on ax1 and make pickable
        leg = ax1.legend(loc='center left', bbox_to_anchor=(1.02, 0.5), fontsize=8, frameon=False)
        leg_map = {}
        if leg:
            for legline, country in zip(leg.get_lines(), labels1):
                legline.set_picker(True)
                legline.set_pickradius(10)
                leg_map[legline] = country
                
        fig.subplots_adjust(left=0.1, right=0.7, top=0.9, bottom=0.1)
        
        text = "【图表解读】\n   展示历史销量与价格走势。右侧图例附带自动生命周期判定。点击图例可高亮对应国家。\n"
        text += "📊 【数据口径】: 取用历年 Volume_Clean 汇总，价格基于销售额/数量折算。\n"
        text += "💡 【汇报话术】: “老板，看左侧的价格曲线是否已经走平。如果价格已经进入‘L型’底部（平台期），说明价格战基本见底，是我们可以带着成本优势吃利润的安全区。重点关注标为【快速放量期】的国家。”"
        return text, leg_map

    def draw_penetration(self, fig, df):
        ax = fig.add_subplot(111)
        res = []
        opps = []
        for c in self.current_countries:
            sub = df[df['国家']==c]
            total = sub['Sales_Clean'].sum()
            orig = sub[sub['集团/企业'].isin(self.selected_companies_for_originator)]['Sales_Clean'].sum()
            if total>0:
                gen_share = (total-orig)/total*100
                res.append({'C':c, 'G':gen_share, 'O':orig/total*100})
                if gen_share < 20: opps.append(f"{c}(原研{orig/total:.0%})")
        
        if not res: return "无数据"
        pd.DataFrame(res).set_index('C').sort_values('G').plot(kind='barh', stacked=True, ax=ax, color=[COLORS['secondary'], COLORS['danger']])
        ax.legend(["Generic", "Originator"], bbox_to_anchor=(1.02, 1), loc='upper left')
        fig.subplots_adjust(right=0.8)
        
        text = "【图表解读】\n   红色=原研药占比，绿色=仿制药占比。\n"
        if opps: text += f"   • 最佳替代切入点: {', '.join(opps)}。\n\n"
        text += "📊 【数据口径】: 匹配 Originator Config 配置中的企业作为原研药。计算结果为纯销售额的份额占比。\n"
        text += "💡 【汇报话术】: “老板，红色柱子越长的地方，原研品牌粘性越强（如意大利）。这类市场我们需要找强势本地代理商；红色非常短的市场，说明已经是纯粹的红海大混战了。”"
        return text

    def draw_pie_batch(self, fig, df, countries):
        if not countries:
            return "无数据"
        cols = 3
        rows = (len(countries) + cols - 1) // cols
        fig.set_size_inches(12, max(6, rows * 5))
        axes = fig.subplots(rows, cols).flatten()
        fig.subplots_adjust(wspace=0.6, hspace=0.8)
        
        leaders = []
        for i, ax in enumerate(axes):
            if i < len(countries):
                c = countries[i]
                sub = df[df['国家']==c]
                d = sub.groupby('集团/企业')['Sales_Clean'].sum().sort_values(ascending=False)
                if d.sum() > 0:
                    total_sales = d.sum()
                    shares = d / total_sales * 100
                    hhi = (shares ** 2).sum()
                    top1 = shares.iloc[0] / 100
                    
                    leaders.append(f"{c}(Top1 {top1:.0%})")
                    # Keep Top 5 + Others logic
                    if len(d)>5: d = pd.concat([d.iloc[:5], pd.Series([d.iloc[5:].sum()], index=['Others'])])
                    
                    wedges, texts, autotexts = ax.pie(d, labels=None, autopct='%1.0f%%', textprops={'fontsize':8}, radius=0.7)
                    ax.axis('equal')
                    # Add legend
                    legend_labels = [f"{label} ({val/d.sum()*100:.1f}%)" for label, val in zip(d.index, d.values)]
                    ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(0.95, 0.5), fontsize=7, frameon=False)
                    
                    title = f"{c}\n(HHI: {hhi:.0f})"
                    ax.set_title(title, pad=15)
                    
                    if hhi > 2500:
                        ax.text(0, -1.3, "高度寡头垄断", color='white', ha='center', fontsize=8, fontweight='bold', bbox=dict(facecolor='red', alpha=0.8, edgecolor='none', boxstyle='round,pad=0.3'))
                    elif hhi < 1500:
                        ax.text(0, -1.3, "市场分散，存在机会", color='white', ha='center', fontsize=8, fontweight='bold', bbox=dict(facecolor='green', alpha=0.8, edgecolor='none', boxstyle='round,pad=0.3'))
                else: ax.axis('off')
            else: ax.axis('off')
            
        fig.subplots_adjust(wspace=0.6, hspace=0.8)
        
        text = f"【仿制药格局（已剔除原研药干扰）】\n   • 领头羊: {'; '.join(leaders[:3])}。\n\n"
        text += "📊 【数据口径】: 仅计算非原研药企业的存量博弈份额。HHI(市场集中度)指数 = 仿制药企业市占率平方和。\n"
        text += "💡 【汇报话术】: “老板，HHI > 2500的地方（标红）代表已经形成了极强的寡头垄断，比如有本地大厂控盘，贸然进入大概率当炮灰；我们挑 HHI < 1500 的‘绿色’分散市场做横向盘整最有利可图。”"
        return text

    def draw_dynamic_matrix(self, fig, df, x_metric, y_metric):
        ax = fig.add_subplot(111)
        data = []
        df_sub = df[df['国家'].isin(self.current_countries)]
        
        for c in self.current_countries:
            sub = df_sub[df_sub['国家']==c]
            if sub.empty: continue
            
            tot_sales = sub['Sales_Clean'].sum()
            tot_vol = sub['Volume_Clean'].sum()
            
            orig_sales = sub[sub['集团/企业'].isin(self.selected_companies_for_originator)]['Sales_Clean'].sum()
            orig_share = orig_sales / tot_sales if tot_sales > 0 else 0
            
            gen_sub = sub[~sub['集团/企业'].isin(self.selected_companies_for_originator)]
            comps = gen_sub['集团/企业'].nunique()
            gen_sales = gen_sub['Sales_Clean'].sum()
            if gen_sales > 0:
                shares = gen_sub.groupby('集团/企业')['Sales_Clean'].sum() / gen_sales * 100
                hhi = (shares**2).sum()
            else:
                hhi = 0
                
            ts = sub.groupby('年份').apply(lambda x: pd.Series({
                'V': x['Volume_Clean'].sum(),
                'P': x['Sales_Clean'].sum()/x['Volume_Clean'].sum() if x['Volume_Clean'].sum()>0 else 0
            })).sort_index()
            
            cagr = 0; price_drop = 0
            if len(ts) >= 2:
                v0 = ts.iloc[0]['V']; v1 = ts.iloc[-1]['V']
                p0 = ts.iloc[0]['P']; p1 = ts.iloc[-1]['P']
                if v0 > 0: cagr = (v1/v0)**(1/(len(ts)-1)) - 1
                if p0 > 0: price_drop = (p0 - p1) / p0  # Positive means dropped
                
            data.append({
                'C': c,
                '市场绝对规模': tot_sales,
                '原研替代空间': 1 - orig_share,
                '价格降幅': price_drop,
                '市场增速': cagr,
                '竞争者数量': comps,
                '集中度HHI': hhi,
                'Vol': tot_vol # bubble size
            })
            
        df_mat = pd.DataFrame(data)
        if df_mat.empty: return "无数据"
        
        x = df_mat[x_metric]
        y = df_mat[y_metric]
        sizes = df_mat['Vol']
        sf = 2000 / sizes.max() if sizes.max() > 0 else 1
        
        self._matrix_countries = df_mat['C'].tolist()
        
        # Make sure sizes are at least visible
        s_arr = [max(20, v) for v in sizes*sf]
        scatter = ax.scatter(x, y, s=s_arr, alpha=0.6, edgecolors='k', picker=True)
        
        for i, r in df_mat.iterrows():
            ax.text(r[x_metric], r[y_metric], r['C'], fontsize=9, ha='center', va='center')
            
        ax.axvline(x.median(), ls=':', color='gray')
        ax.axhline(y.median(), ls=':', color='gray')
        ax.set_xlabel(x_metric)
        ax.set_ylabel(y_metric)
        
        if x_metric in ["原研替代空间", "价格降幅", "市场增速"]:
            ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda val, pos: f"{val:.0%}"))
        if y_metric in ["原研替代空间", "价格降幅", "市场增速"]:
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, pos: f"{val:.0%}"))
            
        text = f"【多因素矩阵】  X轴: {x_metric} | Y轴: {y_metric} | 气泡大小: 整体销量\n\n"
        text += "📊 【数据口径】: X和Y轴数据由上方控件动态生成，自动过滤空数据。\n"
        text += "💡 【汇报话术】: “老板，这个图是我们的战略导航。选【规模】对比【增速】，能帮我们找明星级重磅市场；选【原研替代空间】对比【竞争者数量】能帮我们寻找降维打击的‘蓝海’；直接点击这上面的重磅国家气泡，就能自动跳到该国的对手盘进行尽调（Drill-down）！”"
        return text

    def draw_ma_threats(self, fig, df):
        # 手动布局代替 gridspec tight_layout
        ax = fig.add_axes([0.1, 0.1, 0.6, 0.8]) # Left Chart
        ax_list = fig.add_axes([0.75, 0.1, 0.2, 0.8]) # Right Text
        ax_list.axis('off')
        
        res = []
        ghost_txt = []
        for c in self.current_countries:
            act_comps = set(df[(df['国家']==c)&(df['Sales_Clean']>0)]['集团/企业'])
            act = len(act_comps)
            ma_comps = set()
            if self.df_ema is not None:
                ma_comps = set(self.df_ema[(self.df_ema['Country']==c)&(self.df_ema['Substance'].str.contains(self.combo_drugs.get(), na=False))]['MAH'])
            ghosts = list(ma_comps - act_comps)
            res.append({'C':c, 'A':act, 'S':len(ghosts)})
            if ghosts: ghost_txt.append(f"{c}: {', '.join(ghosts[:3])}...")

        if res:
            pd.DataFrame(res).set_index('C').sort_values('A').plot(kind='barh', stacked=True, ax=ax, color=[COLORS['secondary'], COLORS['neutral']])
            ax.legend(["Active", "Sleeping"], loc='lower right')
        
        y = 1.0
        ax_list.text(0, y, "【Ghost List】", fontweight='bold'); y-=0.05
        for t in ghost_txt[:15]: ax_list.text(0, y, f"• {t}", fontsize=8); y-=0.06
        
        text = "【潜在威胁】 灰色条代表获得了批准(MA)但是根本没有卖货的'幽灵玩家'(Ghosts)。\n\n"
        text += "📊 【数据口径】: 销售系统（Active）与 EMA 审批注册数据库（MA Headers）的比对差集。\n"
        text += "💡 【汇报话术】: “老板，右边这列幽灵名单（Ghost List）全是手握批文在观望的狼。如果一个国家灰色条特别长，说明一旦市场价格上涨，随时会有十几家公司带着批文杀进来引发价格战雪崩，不能盲目乐观。”"
        return text

    def draw_bubble_matrix(self, fig, df):
        ax = fig.add_subplot(111)
        df_sub = df[df['国家'].isin(self.current_countries)]
        
        data = []
        for c in self.current_countries:
            sub = df_sub[df_sub['国家']==c]
            if sub.empty: continue
            
            tot_vol = sub['Volume_Clean'].sum()
            tot_sales = sub['Sales_Clean'].sum()
            avg_price = (tot_sales / tot_vol * 0.3) if tot_vol > 0 else 0
            
            data.append({
                'Country': c,
                'Volume': tot_vol,
                'Sales': tot_sales,
                'Price': avg_price
            })
            
        if not data:
            ax.text(0.5, 0.5, "无数据", ha='center', va='center', color='gray')
            return "无数据"
            
        df_p = pd.DataFrame(data)
        
        x = df_p['Price']
        y = df_p['Volume']
        sizes = df_p['Sales']
        
        max_size = sizes.max()
        sf = 3000 / max_size if max_size > 0 else 1
        s_arr = [max(50, v) for v in sizes * sf]
        
        # Color based on distinct colors for pretty visualization
        colors = [DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(df_p))]
        
        scatter = ax.scatter(x, y, s=s_arr, alpha=0.6, c=colors, edgecolors='white', picker=True)
        
        for i, r in df_p.iterrows():
            ax.text(r['Price'], r['Volume'], r['Country'], fontsize=9, ha='center', va='center')
            
        med_x = x.median() if not x.empty else 0
        med_y = y.median() if not y.empty else 0
        
        ax.axvline(med_x, ls='--', color='gray', alpha=0.7, label=f'Price Median ({med_x:.2f})')
        ax.axhline(med_y, ls='--', color='gray', alpha=0.7, label=f'Volume Median ({med_y:,.0f})')
        
        ax.set_xlabel("预估出厂单价 (Factory Price, USD)")
        ax.set_ylabel("总销量 (Volume, Units)")
        ax.set_title("核心市场战略机遇矩阵 (气泡大小: 销售额)")
        ax.grid(True, linestyle=':', alpha=0.4)
        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left')
        
        fig.subplots_adjust(bottom=0.15, right=0.8)
        text = "【市场预测气泡矩阵】 横坐标: 出厂均价 | 纵坐标: 销量 | 气泡大小: 总体销售额\n\n"
        text += "📊 【数据口径】: 抽取所有历史总销量与对应出厂单价截面进行投影映射。虚线为中位线。\n"
        text += "💡 【汇报话术】: “老板，这个图右上角的市场（价格双高），就是最优质的‘肥肉市场’，利润丰厚且容量巨大。左下角的是典型低价内卷且规模鸡肋的‘红海’，建议直接放弃。”"
        
        return text

    def draw_prediction(self, fig, df, sim_price=0.4, sim_hhi=1200):
        # 原版混合预测Hybrid
        ax = fig.add_subplot(111)
        df_sub = df[df['国家'].isin(self.current_countries)]
        summary = []
        res_data = []
        
        for c in self.current_countries:
            sub = df_sub[df_sub['国家']==c].groupby('年份').apply(lambda x: pd.Series({
                'Vol': x['Volume_Clean'].sum()
            })).reset_index().sort_values('年份')
            
            if len(sub) < 3: continue
            try:
                x = sub['年份'].astype(str).str.extract(r'(\d{4})')[0].astype(int)
                y = sub['Vol']
                
                # Hybrid Prediction Logic
                x_fit = x[x < 2025]; y_fit = y[x < 2025]
                
                if len(x_fit) >= 2:
                    slope, intercept, r_val, _, _ = stats.linregress(x_fit, y_fit)
                    pred_linear = slope * 2025 + intercept
                    actual_2025_h1 = y[x==2025].values[0] if 2025 in x.values else 0
                    pred_runrate = actual_2025_h1 * 2.0
                    r_sq = r_val ** 2
                    weight_linear = 0.7 if r_sq > 0.8 else 0.3
                    final_pred = weight_linear * pred_linear + (1 - weight_linear) * pred_runrate
                    
                    if final_pred > 0:
                        pace = actual_2025_h1 / final_pred
                        res_data.append({'C': c, 'Pace': pace * 100})
                        status = "超预期" if pace > 0.55 else "滞后"
                        summary.append(f"• {c}: 预测{final_pred/1e6:.1f}M，当前{status}({pace:.0%})")
            except: pass

        if res_data:
            df_p = pd.DataFrame(res_data).set_index('C').sort_values('Pace')
            # 颜色逻辑：大于50%绿色，小于50%橙色
            colors = [COLORS['secondary'] if p > 50 else COLORS['warning'] for p in df_p['Pace']]
            df_p['Pace'].plot(kind='barh', ax=ax, color=colors)
            
            max_p = df_p['Pace'].max()
            ax.set_xlim(0, max(100, max_p * 1.1))
            
            for container in ax.containers:
                ax.bar_label(container, fmt='%.0f%%', padding=3)
                
            ax.axvline(50, ls='--', color='red', label='H1 Target (50%)')
            ax.set_title("2025 Forecast Achievement (Hybrid Model: Trend + RunRate)")
            ax.legend()
            
        fig.subplots_adjust(right=0.8)
        text = "【智能预测】 基于历史趋势与H1跑率的混合预测。红色虚线为50%半年度达标线。\n"
        text += "\n".join(summary)
        
        text += f"\n\n【What-If 情景模拟】 假定定价为原研药的 {sim_price:.0%}，目标环境 HHI={sim_hhi}:\n"
        expected_share = (1 - sim_price) * (2000 / max(sim_hhi, 500)) * 0.15
        expected_share = max(0.0, float(min(expected_share, 0.45)))
        text += f"   ➤ 预计在上述环境下，本产品第3年市场份额可达: {expected_share:.1%}\n\n"
        text += "💡 【汇报话术】: “老板，按照上面我拖拽设定的激进价格打法（原研价折让率），结合该国的垄断难度（HHI滑块），系统用测算公式预估我们第三年能抢下这么多份额，这就给我们明年制定欧洲整体销量盘子（Budget）托住了底。”"
        
        return text

if __name__ == "__main__":
    root = tk.Tk()
    app = PharmaAnalyzerProV24(root)
    root.mainloop()