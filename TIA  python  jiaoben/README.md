# TIA Portal 变量导出工具集

> **操作原则：所有脚本仅执行读取操作，不修改任何项目数据。**

## 目录结构

```
python  jiaoben/
├── README.md                        ← 本说明文件
├── check_environment.py             ← 环境检查脚本（先运行这个）
├── export_tia_variables.py          ← Openness API 自动导出（方案 A，首选）
├── export_e1_modbus.py              ← e1_ 变量 Modbus 地址导出 🆕
├── parse_exported_variables.py      ← 手动导出解析（方案 B，备选）
├── generate_test_data.py            ← 生成模拟测试表格
├── 手动导出/                        ← 存放 TIA Portal 手动导出的 Excel 文件
└── 导出结果/                        ← 所有输出文件存放位置
```

---

## 🆕 方案 M：一号挤出机 Modbus 地址导出

专门针对 JSB-25-081B 项目一号挤出机（e1），通过 Openness API 获取所有 `e1_` 变量的逻辑地址，自动换算为 Modbus 地址，并更新设备说明文档。

### 操作步骤

```powershell
# 前置：启动 TIA Portal，打开 JSB-25-081B TPV包纱管1.0 项目

# 运行导出（一键完成）
python export_e1_modbus.py
```

### 输出
- `导出结果/一号挤出机_e1_Modbus地址_<时间戳>.xlsx` — 完整的 Excel 变量表（含 Modbus 地址）
- `设备说明/一号挤出机变量清单.md` — **自动更新**的 Markdown 清单（原文件自动备份为 .bak）

### Modbus 地址换算规则

| 逻辑地址类型 | Modbus 区域 | 地址公式 | 功能码 |
|------------|-----------|---------|-------|
| `%Ix.y` | 离散输入 (1xxxx) | `10001 + x*8 + y` | FC02 |
| `%Qx.y` | 线圈 (0xxxx) | `1 + x*8 + y` | FC01/05/15 |
| `%Mx.y` | 线圈 (0xxxx)-M区 | `1 + x*8 + y` | FC01/05/15 |
| `%MWx` | 保持寄存器 (4xxxx) | `40001 + x` | FC03/06/16 |
| `%MDx` | 保持寄存器 (4xxxx)-双字 | `40001 + x` | FC03/06/16 (2 regs) |
| `%IWx` | 输入寄存器 (3xxxx) | `30001 + x` | FC04 |
| `%DBn.DBWx` | 保持寄存器 (4xxxx) | DB 基址 + 偏移 | FC03/06/16 |

> ⚠️ 实际 Modbus 地址取决于 CoTrust CTH2-277PN / ADFweb HD67607 网关设备的地址映射配置。以上为西门子标准映射规则推算值。

---

## 方案 A：Openness API 自动导出（首选）

### 前置条件
1. TIA Portal **Professional** 或 **Advanced** 版本
2. 安装时勾选了 **Openness** 组件
3. Python 3.8+ 64位（与 TIA Portal 架构一致）

### 操作步骤

```powershell
# 第1步：检查环境
python check_environment.py

# 如果缺依赖，安装它们：
python check_environment.py --install

# 第2步：启动 TIA Portal，打开需要导出的项目

# 第3步：运行导出
python export_tia_variables.py
```

### 输出
- `导出结果/TIA_Variables_<项目名>_<时间戳>.xlsx`
- 包含 4 个 Sheet：PLC变量表、DB块结构、HMI变量、IO映射

---

## 方案 B：手动导出 + 解析（备选）

适用于 TIA Portal Basic 版本或无 Openness 组件的情况。

### 操作步骤

```powershell
# 第1步：在 TIA Portal 中手动导出变量表
#   打开 PLC 变量表 → Ctrl+A → Ctrl+C → 粘贴到 Excel → 保存 .xlsx
#   同样操作 HMI 变量表
#   将文件放入: python  jiaoben/手动导出/

# 第2步：运行解析脚本
python parse_exported_variables.py

# 也可以指定自定义目录：
python parse_exported_variables.py "D:\我的导出文件"
```

### 输出
- `导出结果/变量汇总_<时间戳>.xlsx`
- 包含 3 个 Sheet：全部变量、按分类分组、按数据类型

---

## 测试验证

```powershell
# 生成模拟的 TIA Portal 变量表（在 手动导出/ 目录下生成 3 个 .xlsx）
python generate_test_data.py

# 运行解析脚本验证功能
python parse_exported_variables.py

# 检查输出
explorer 导出结果
```

生成的测试数据包含：
| 文件 | 内容 | 变量数 |
|------|------|--------|
| `PLC变量表_模拟.xlsx` | 30个I/Q/M/DB 变量 + 6个系统变量 | 36 |
| `HMI变量_模拟.xlsx` | 16个 HMI 连接变量 | 16 |
| `DB块_模拟.xlsx` | 12个配方参数 + 10个机器参数 | 22 |

---

## 方案对比

| 特性 | 方案 M (Modbus) | 方案 A (Openness API) | 方案 B (手动+解析) |
|------|----------------|----------------------|-------------------|
| 自动化程度 | 全自动 | 全自动 | 半自动 |
| 数据完整性 | e1_ 变量 + Modbus | 高（PLC/HMI/DB/IO） | 中（仅导出内容） |
| TIA Portal 版本 | Professional/Advanced | Professional/Advanced | 所有版本 |
| 额外依赖 | pythonnet + Openness | pythonnet + Openness | 仅 openpyxl |
| 操作复杂度 | 一键运行 | 一键运行 | 需手动导出多个文件 |
| Modbus 地址 | ✅ 自动计算 | ❌ | ❌ |

---

## 项目信息

| 项目 | TIA 版本 | 主设备 |
|------|---------|--------|
| JSB-25-081B TPV包纱管1.0 | V17 | inoex/Sikora/海康/基恩士等 8 个第三方设备 |
| 针织机_V19 | V19（从V16升级） | S7-1200 + KTP700 + G120 |

## 故障排除

**Q: 运行 `check_environment.py` 提示找不到 Openness DLL？**
A: TIA Portal Basic 不支持 Openness。请使用方案 B。或安装 TIA Portal Professional。

**Q: `pythonnet` 安装失败？**
A: 确保 Python 是 64 位版本。运行 `python -c "import struct; print(struct.calcsize('P') * 8)"` 确认。

**Q: 导出的文件中缺少某些变量？**
A: 方案 B 需要手动导出每个变量表和 DB 块。方案 A 会自动遍历所有。

**Q: 变量地址显示为空？**
A: 如果变量使用符号访问（无绝对地址），Address 列会显示为符号名。

**Q: Modbus 地址与实际设备不符？**
A: 脚本按西门子标准映射规则推算。实际地址取决于网关设备组态，请以 TIA Portal 中 CoTrust CTH2-277PN / ADFweb HD67607 的 I/O 映射表为准。