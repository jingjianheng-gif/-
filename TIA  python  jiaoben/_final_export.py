#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""补充 DB 块变量到一号挤出机变量清单"""
import sys, os, re, datetime
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = Path(__file__).parent.resolve()
OUTPUT_DIR = SCRIPT_DIR / "导出结果"
MD_PATH = SCRIPT_DIR / "设备说明" / "一号挤出机变量清单.md"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── DB 块定义 ──
# 从 PEData.plf 解析出的 DB 编号:
# DB14=flange, DB15=connector, DB19=melt, DB20=watertemp, DB35=current, DB44=air_pump, DB45=alarm

DB_BLOCKS = {
    14: {
        "name": "e1_flange_db",
        "desc": "法兰温度PID",
        "channel": "16#0005",
        "members": [
            ("LADDR", "HW_IO", 0.0, "模块硬件标识符"),
            ("Channel", "Word", 2.0, "通道号(0-7)"),
            ("Working_T_SV", "Real", 4.0, "设定温度"),
            ("Temperature_L_SV", "Real", 8.0, "温度低报警偏差设定值"),
            ("Temperature_H_SV", "Real", 12.0, "温度高报警偏差设定值"),
            ("Batch_Temp_L_SV", "Real", 16.0, "批次控制温度低报警偏差设定值"),
            ("Batch_Temp_H_SV", "Real", 20.0, "批次控制温度高报警偏差设定值"),
            ("Temperature_HH_SV", "Real", 24.0, "温度超高偏差设定值（切断接触器）"),
            ("CotrolByte", "Byte", 28.0, "控制字"),
            ("Cycle", "Word", 30.0, "脉冲输出周期"),
            ("Kp", "Real", 32.0, "比例系数"),
            ("Ti", "Int", 36.0, "积分时间"),
            ("Td", "Int", 38.0, "微分时间"),
            ("Mp", "Int", 40.0, "较准值"),
            ("Batch_Enable", "Bool", 42.0, "批次控制"),
            ("E_Stop", "Bool", 42.1, "急停"),
            ("PID_Enable", "Bool", 42.2, "PID使能"),
            ("Fan_Fault", "Bool", 42.3, "风机故障"),
            ("Time_Switch_Off_SV", "Time", 44.0, "接触器切断延时设定"),
        ]
    },
    15: {
        "name": "e1_connector_db",
        "desc": "连接器温度PID",
        "channel": "16#0006",
        "members": [
            ("LADDR", "HW_IO", 0.0, "模块硬件标识符"),
            ("Channel", "Word", 2.0, "通道号(0-7)"),
            ("Working_T_SV", "Real", 4.0, "设定温度"),
            ("Temperature_L_SV", "Real", 8.0, "温度低报警偏差设定值"),
            ("Temperature_H_SV", "Real", 12.0, "温度高报警偏差设定值"),
            ("Batch_Temp_L_SV", "Real", 16.0, "批次控制温度低报警偏差设定值"),
            ("Batch_Temp_H_SV", "Real", 20.0, "批次控制温度高报警偏差设定值"),
            ("Temperature_HH_SV", "Real", 24.0, "温度超高偏差设定值（切断接触器）"),
            ("CotrolByte", "Byte", 28.0, "控制字"),
            ("Cycle", "Word", 30.0, "脉冲输出周期"),
            ("Kp", "Real", 32.0, "比例系数"),
            ("Ti", "Int", 36.0, "积分时间"),
            ("Td", "Int", 38.0, "微分时间"),
            ("Mp", "Int", 40.0, "较准值"),
            ("Batch_Enable", "Bool", 42.0, "批次控制"),
            ("E_Stop", "Bool", 42.1, "急停"),
            ("PID_Enable", "Bool", 42.2, "PID使能"),
            ("Fan_Fault", "Bool", 42.3, "风机故障"),
            ("Time_Switch_Off_SV", "Time", 44.0, "接触器切断延时设定"),
        ]
    },
    20: {
        "name": "e1_watertemp_db",
        "desc": "水温PID",
        "channel": "16#0004",
        "members": [
            ("LADDR", "HW_IO", 0.0, "模块硬件标识符"),
            ("Channel", "Word", 2.0, "通道号(0-7)"),
            ("Working_T_SV", "Real", 4.0, "设定温度"),
            ("Temperature_L_SV", "Real", 8.0, "温度低报警偏差设定值"),
            ("Temperature_H_SV", "Real", 12.0, "温度高报警偏差设定值"),
            ("Batch_Temp_L_SV", "Real", 16.0, "批次控制温度低报警偏差设定值"),
            ("Batch_Temp_H_SV", "Real", 20.0, "批次控制温度高报警偏差设定值"),
            ("Temperature_HH_SV", "Real", 24.0, "温度超高偏差设定值（切断接触器）"),
            ("CotrolByte", "Byte", 28.0, "控制字"),
            ("Cycle", "Word", 30.0, "脉冲输出周期"),
            ("Kp", "Real", 32.0, "比例系数"),
            ("Ti", "Int", 36.0, "积分时间"),
            ("Td", "Int", 38.0, "微分时间"),
            ("Mp", "Int", 40.0, "较准值"),
            ("Batch_Enable", "Bool", 42.0, "批次控制"),
            ("E_Stop", "Bool", 42.1, "急停"),
            ("PID_Enable", "Bool", 42.2, "PID使能"),
            ("Fan_Fault", "Bool", 42.3, "风机故障"),
            ("Time_Switch_Off_SV", "Time", 44.0, "接触器切断延时设定"),
        ]
    },
}


def calc_modbus(db_num, byte_offset, bit_offset, dtype):
    """
    计算 DB 块变量的 Modbus 地址。
    
    西门子 S7 Modbus 标准映射（用于保持寄存器 4xxxx）:
      - 字类型 (Int/Word):   HR_addr = DB基址 + byte_offset/2
      - 双字类型 (Real/DInt): HR_addr = DB基址 + byte_offset/2 (占2个寄存器)
      - Bool类型:             线圈 = DB线圈基址 + byte_offset*8 + bit_offset
      - Byte类型:             HR_addr = DB基址 + byte_offset/2
    
    注意: 实际地址取决于网关设备中 DB 块的起始偏移配置。
          此处假设每个 DB 块映射到连续的 Modbus 区域。
    """
    if dtype in ('Bool',):
        # 线圈映射（需要知道 DB 块的线圈基址）
        coil_base = db_num * 100  # 估算值
        coil_addr = coil_base + int(byte_offset * 8 + bit_offset)
        return (str(coil_addr).zfill(5), "线圈 (0xxxx)", "FC01/05/15")
    elif dtype in ('Byte', 'Word', 'Int'):
        hr_base = 40000 + db_num * 10
        hr_addr = hr_base + int(byte_offset // 2)
        return (str(hr_addr), "保持寄存器 (4xxxx)", "FC03/06/16")
    elif dtype in ('Real', 'DInt', 'Time', 'DWord'):
        hr_base = 40000 + db_num * 10
        hr_addr = hr_base + int(byte_offset // 2)
        return (str(hr_addr), f"保持寄存器 (4xxxx)-{dtype}", "FC03/06/16 (2 regs)")
    else:
        hr_base = 40000 + db_num * 10
        hr_addr = hr_base + int(byte_offset // 2)
        return (str(hr_addr), "保持寄存器 (4xxxx)", "FC03/06/16")


# ── 读取已有 Markdown ──
with open(MD_PATH, 'r', encoding='utf-8') as f:
    md_content = f.read()

# ── 生成 DB 块表格 ──
db_lines = []
db_lines.append("## 六、DB 块参数（含 Modbus 地址）")
db_lines.append("")
db_lines.append("> ⚠️ **DB 块的 Modbus 地址为估算值**，基于 DB编号×偏移的简单映射。")
db_lines.append("> 实际地址取决于 CoTrust CTH2-277PN / ADFweb HD67607 网关设备的 I/O 映射表配置。")
db_lines.append("> 请以 TIA Portal 设备组态中的实际映射为准！")
db_lines.append("")

for db_num in sorted(DB_BLOCKS.keys()):
    block = DB_BLOCKS[db_num]
    db_lines.append(f"### {block['name']} [DB{db_num}] — {block['desc']}")
    db_lines.append("")
    db_lines.append("| 变量名 | 类型 | 偏移 | Modbus 地址 | 区域 | 说明 |")
    db_lines.append("|--------|------|------|-----------|------|------|")
    
    for name, dtype, offset, desc in block['members']:
        # 解析 offset: "42.0" → byte=42, bit=0
        if '.' in str(offset):
            bo, bit_off = str(offset).split('.')
            bo, bit_off = int(bo), int(bit_off)
        else:
            bo, bit_off = int(offset), 0
        
        mb_addr, mb_area, fc = calc_modbus(db_num, bo, bit_off, dtype)
        off_str = f"{bo}.{bit_off}" if dtype == 'Bool' else str(bo)
        db_lines.append(f"| `{name}` | {dtype} | {off_str} | {mb_addr} | {mb_area} | {desc} |")
    
    db_lines.append("")

# ── 加入 Modbus 换算规则补充 ──
db_lines.append("### DB 块 Modbus 地址估算公式")
db_lines.append("")
db_lines.append("| 数据类型 | 地址公式 | 说明 |")
db_lines.append("|---------|---------|------|")
db_lines.append("| Word / Int | `(40000 + DB号×10) + 偏移/2` | 单寄存器 |")
db_lines.append("| Real / DInt / Time | `(40000 + DB号×10) + 偏移/2` | 双寄存器(2 regs) |")
db_lines.append("| Bool | `DB号×100 + 偏移×8 + 位` | 线圈 (0xxxx) |")
db_lines.append("")
db_lines.append("> 示例: DB15.offset=4.0, Real → 40000+15×10+4/2 = **40152**（2 regs）")

# ── 更新 Markdown ──
# 找到 "## 五、" 之后插入 DB 块内容
new_db_section = "\n".join(db_lines)
if "## 六、" in md_content:
    md_content = md_content[:md_content.index("## 六、")] + new_db_section
else:
    md_content = md_content.rstrip() + "\n\n---\n\n" + new_db_section + "\n"

with open(MD_PATH, 'w', encoding='utf-8') as f:
    f.write(md_content)

print(f"[✓] Markdown 已更新: {MD_PATH}")

# ── 打印摘要 ──
print("\n=== DB 块 Modbus 地址（关键变量）===")
for db_num in sorted(DB_BLOCKS.keys()):
    block = DB_BLOCKS[db_num]
    print(f"\n{block['name']} [DB{db_num}]:")
    for name, dtype, offset, desc in block['members']:
        if name in ('Working_T_SV', 'Actual_T', 'Kp', 'Ti', 'Td', 'PID_Enable'):
            bo = int(str(offset).split('.')[0])
            bit_off = int(str(offset).split('.')[1]) if '.' in str(offset) else 0
            mb_addr, mb_area, fc = calc_modbus(db_num, bo, bit_off, dtype)
            print(f"  {name:25s} {dtype:8s} offset={offset} → {mb_addr} {mb_area}")
