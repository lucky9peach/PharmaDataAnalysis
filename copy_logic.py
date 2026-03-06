import re

with open("欧洲市场预测promax.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

# Extract constants
constants = []
recording = False
for line in lines:
    if "MARKER_MAP =" in line or "Target_Markets =" in line or "ORIGINATOR_CONFIG =" in line or "COUNTRY_PROFILES =" in line:
        recording = True
    if recording:
        constants.append(line)
        if line.strip() == "}":
            if "COUNTRY_PROFILES" in ''.join(constants[-10:]):
                recording = False
                break

# Extract the class methods we need
class_code = [
    "class AnalysisEngineV24:\n",
    "    def __init__(self):\n",
    "        self.selected_companies_for_originator = []\n",
    "        self.current_countries = []\n",
    "        self.highlighted_country = None\n",
    "        self.df_sales_clean = None\n",
    "        self.pie_batch_index = 0\n",
    "        self.pie_batch_size = 6\n\n"
]

in_class = False
method_lines = []
for line in lines:
    if line.startswith("    def draw_summary_table(self, fig, df):"):
        in_class = True
    if in_class and line.startswith("    def draw_ma_threats"):
        # We don't need ma_threats
        in_class = False
    if in_class:
        if line.startswith("    def on_drug_selected"): # safety boundary
            break
        method_lines.append(line)
        
    if not in_class and line.startswith("    def draw_prediction"):
        in_class = True

with open("step4_visualizer.py", "a", encoding="utf-8") as f:
    f.write("\n\n# ==========================================\n")
    f.write("# 提取自 欧洲市场预测promax.py 的分析引擎引擎配置库\n")
    f.write("# ==========================================\n")
    for l in lines[34:144]: # MARKER MAP to COUNTRY_PROFILE end
        f.write(l)
    
    f.write("\n\n" + "".join(class_code))
    for l in method_lines:
        f.write(l)
        
print("Extraction complete.")
