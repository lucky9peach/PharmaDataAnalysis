import re

with open('step4_visualizer.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Pattern to find generate_chart methods that DON'T have self.last_df = df right inside
pattern = r'(def generate_chart\(self, df: pd\.DataFrame, filters: dict(?:, \*\*kwargs)?\):\n(?:\s+populate_combo = kwargs.get\(\'populate_combo\', True\)\n)?)(\s+self\.figure\.clear\(\)|\s+self\.last_df = df)'

def replacer(match):
    if 'self.last_df = df' in match.group(2):
        return match.group(0) # Already has it
    indent = match.group(2).split('self.figure.clear')[0]
    return match.group(1) + indent + "self.last_df = df" + indent + "self.last_filters = filters" + match.group(2)

new_content = re.sub(pattern, replacer, content)

with open('step4_visualizer.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Fixed step4_visualizer.py")
