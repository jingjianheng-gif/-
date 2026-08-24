import os, re

LUCENE = r"C:\Users\Administrator\Documents\Project\JSB-25-081B（瑞源橡塑）\JSB-25--081B(瑞源橡塑）TPV包纱管1.0\IM\SearchIndex"

# .fdt = 字段数据, .fdx = 字段索引
# 每个文档存储了变量名、类型、地址、注释等
fdt = os.path.join(LUCENE, '_1s22.fdt')
with open(fdt, 'rb') as f:
    data = f.read()
print(f'.fdt 大小: {len(data)} bytes')

# Lucene .fdt 格式: [doc_count][doc1_len][doc1_fields...][doc2_len]...
# 每个 field: [field_number(VInt)][bits][value]
# 直接搜索 e1_
matches = []
pos = 0
while pos < len(data) - 5:
    # 找 e1_ 字节序列
    if data[pos:pos+3] == b'e1_':
        # 向前向后扩展找完整字符串
        start = pos
        while start > 0 and 32 <= data[start-1] < 127:
            start -= 1
        end = pos + 3
        while end < len(data) and (32 <= data[end] < 127 or data[end] == ord('_')):
            end += 1
        name = data[start:end].decode('ascii', errors='ignore').strip()
        if name.startswith('e1_') and len(name) > 5:
            # 看看后续字节中是否包含地址信息
            # 尝试找附近的 %I, %Q, %MW 等模式
            nearby = data[end:end+200]
            nearby_text = nearby.decode('ascii', errors='replace')
            # 找地址
            addr_match = re.search(r'%\w[\w.]*', nearby_text)
            addr = addr_match.group(0) if addr_match else ''
            matches.append((pos, name, addr))
        pos = end
    else:
        pos += 1

print(f'找到 {len(matches)} 处 e1_ (在 .fdt 中)')
# 去重
seen = set()
for pos, name, addr in matches:
    if name not in seen:
        seen.add(name)
        print(f'  {name} 地址={addr}')
