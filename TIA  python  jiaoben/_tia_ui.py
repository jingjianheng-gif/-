import uiautomation as auto
import win32clipboard
import time
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 1. 激活 TIA Portal
tia = auto.WindowControl(searchDepth=2, ClassName='WindowsForms10.Window.8.app.0.12ab327_r8_ad1')
if not tia.Exists(3):
    print("ERROR: TIA Portal not found")
    sys.exit(1)

print("激活 TIA Portal...")
tia.SetFocus()
time.sleep(0.5)

# 2. 尝试查找包含变量表数据的控件
# TIA Portal 的变量表通常是 WinForms DataGridView 或 WPF DataGrid
# 尝试通过 AutomationId 或 Name 来找

def crawl_all(ctrl, depth=0, max_depth=8):
    if depth > max_depth:
        return None
    try:
        ct = ctrl.ControlTypeName
        cls = (ctrl.ClassName or '')
        name = (ctrl.Name or '')[:40]
        aid = (ctrl.AutomationId or '')[:40]

        # 找 DataGrid / 表格
        if 'DataGrid' in ct or 'DataGrid' in cls or 'GridView' in cls or 'GridControl' in cls:
            print(f"FOUND GRID: [{ct}] {cls} - {name}")
            return ctrl

        # 找包含变量名的文本
        if ct == 'Edit' and 'tag' in name.lower():
            print(f"  Filter box: {name}")

        for child in ctrl.GetChildren():
            result = crawl_all(child, depth+1, max_depth)
            if result:
                return result
    except:
        pass
    return None

print("搜索变量表控件...")
grid = crawl_all(tia)
if grid:
    print(f"\n找到变量表! 尝试读取...")
    try:
        # 尝试获取行数
        rows = grid.GetChildren()
        print(f"行数: {len(rows)}")
        for r in rows[:3]:
            cells = r.GetChildren()
            vals = [(c.Name or '')[:20] for c in cells[:6]]
            print(f"  {vals}")
    except Exception as e:
        print(f"读取失败: {e}")
else:
    print("\n未能定位到变量表 DataGrid。")
    print("请确认 TIA Portal 中 PLC 变量表已打开并可见。")
    print("\n如果已打开，将尝试键盘导出...")
