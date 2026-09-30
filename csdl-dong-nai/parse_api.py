# -*- coding: utf-8 -*-
"""Bóc toàn bộ Request body của mục VI (đầu mục 39-75) từ tài liệu API CSDL Đồng Nai."""
import io, json, re, sys

SRC = "_rs/api_full.txt"
lines = io.open(SRC, encoding="utf-8").read().split("\n")
lines = [l for l in lines if not l.startswith("=== PAGE ")]

# --- tìm vị trí tiêu đề đầu mục "NN. ..." trong mục VI
head_re = re.compile(r"^\s*(\d{1,3})\.\s+(\S.*?)\s*$")
heads = []
for i, l in enumerate(lines):
    m = head_re.match(l)
    if not m:
        continue
    n = int(m.group(1))
    if 1 <= n <= 80 and ("nhận" in m.group(2) or "Nhận" in m.group(2)):
        heads.append((i, n, m.group(2).strip()))

# giữ dãy tăng dần cuối cùng (mục VI), bỏ trùng ở mục lục
seq = []
for i, n, t in heads:
    while seq and seq[-1][1] >= n:
        seq.pop()
    seq.append((i, n, t))
items = {n: (i, t) for i, n, t in seq}

FIELD = re.compile(r'^\s*[“"\']([A-Z0-9_]+)[”"\']\s*:\s*(.*)$')


def parse_body(start, end):
    """Đọc block Request body, trả về [(path, name, batbuoc)]"""
    # tìm dòng 'Request body'
    b = None
    for i in range(start, end):
        if lines[i].strip().startswith("Request body"):
            b = i + 1
            break
    if b is None:
        return None, None
    stack, out = [], []
    depth = 0
    started = False
    for i in range(b, end):
        s = lines[i].strip()
        if not s:
            continue
        if re.match(r"^(Success Response|Response|Error Response|Notes|Sample Call|Danh m)", s):
            break
        m = FIELD.match(s)
        if m:
            name, rest = m.group(1), m.group(2)
            if rest.startswith("{"):
                stack.append(name); depth += 1
            elif rest.startswith("["):
                stack.append(name + "[]"); depth += 1
            else:
                bb = "bắt buộc" in rest
                out.append((".".join([x for x in stack if x]), name, bb))
            continue
        # dấu ngoặc trần
        for ch in s:
            if ch == "{":
                if not started:
                    started = True
                else:
                    stack.append(None); depth += 1
            elif ch in "}]":
                if stack:
                    stack.pop()
                    depth -= 1
                elif started:
                    depth -= 1
        if started and depth < 0:
            break
    return out, b


# --- ranh giới mỗi đầu mục
keys = sorted(items)
res = {}
for k in range(39, 76):
    if k not in items:
        print("THIEU DAU MUC", k); continue
    st = items[k][0]
    nxt = [items[x][0] for x in keys if items[x][0] > st]
    en = min(nxt) if nxt else len(lines)
    body, _ = parse_body(st, en)
    res[k] = {"stt": k, "ten": items[k][1], "fields": [
        {"path": p, "name": n, "bb": bb} for p, n, bb in (body or [])]}

json.dump(res, io.open("_rs/api_fields.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
tot = 0
for k in sorted(res):
    c = len(res[k]["fields"])
    tot += c
    print(k, c, res[k]["ten"][:60])
print("TONG TRUONG =", tot)
