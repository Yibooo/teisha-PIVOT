"""スライドの重なり・はみ出しチェック
使い方: python3 check.py layout.json study.pdf
- 描画された文字（pdftotext -bbox）を、記録した各テキスト枠に割り当てて「実際の文字の範囲」を出す
- どのテキスト枠にも入らない文字 → はみ出し
- 異なるテキスト枠の実際の文字範囲どうしが重なる → 文字の重なり
- 図形・画像・文字範囲が「部分的に」重なる（どちらかが他方を含むのはOK）→ 配置の重なり
- 本文の一番下がフッター線（7.0in）より下 → フッターとの衝突
"""
import json, subprocess, sys, re, html
from collections import defaultdict

layout, pdf = sys.argv[1], sys.argv[2]
reg = json.load(open(layout, encoding="utf-8"))
out = subprocess.run(["pdftotext", "-bbox", pdf, "-"], capture_output=True, text=True).stdout

pages, cur = [], None
for line in out.splitlines():
    if "<page " in line:
        cur = []; pages.append(cur)
    m = re.search(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*)</word>', line)
    if m and cur is not None:
        x0, y0, x1, y1 = (float(m.group(i)) / 72 for i in range(1, 5))
        cur.append((x0, y0, x1, y1, html.unescape(m.group(5))))

TOL = 0.04
def inside(px, py, r, t=TOL):
    return r["x"] - t <= px <= r["x"] + r["w"] + t and r["y"] - t <= py <= r["y"] + r["h"] + t
def inter(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0]); h = min(a[3], b[3]) - max(a[1], b[1])
    return (w, h) if w > 0 and h > 0 else None
def contains(a, b, t=TOL):
    return a[0] - t <= b[0] and a[1] - t <= b[1] and a[2] + t >= b[2] and a[3] + t >= b[3]
def box(r): return (r["x"], r["y"], r["x"] + r["w"], r["y"] + r["h"])

issues = []
by = defaultdict(list)
for r in reg: by[r["slide"]].append(r)

for sn, words in enumerate(pages, 1):
    els = by.get(sn, [])
    texts = [e for e in els if e["kind"] in ("text", "overlay")]
    ext = {}
    for (x0, y0, x1, y1, t) in words:
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if all(ord(ch) >= 0x2B00 and not ('\u3000' <= ch <= '\u9fff' or '\uff00' <= ch <= '\uffef') for ch in t):
            continue  # 絵文字だけの語は、フォントの外枠が大きく誤検出になるので除外
        cands = [i for i, e in enumerate(texts) if inside(cx, cy, e)]
        # 重なった枠に吸い込まれないよう、文字列が一致する枠を優先
        norm = lambda z: re.sub(r"\s", "", z)
        match = [i for i in cands if norm(t) and norm(t) in norm(texts[i].get("str", ""))]
        if match: cands = match
        if not cands:
            issues.append((sn, "はみ出し", f"どの枠にも入らない文字「{t}」 at ({cx:.2f},{cy:.2f})"))
            continue
        i = min(cands, key=lambda i: texts[i]["w"] * texts[i]["h"])
        b = ext.get(i)
        ext[i] = (x0, y0, x1, y1) if b is None else (min(b[0], x0), min(b[1], y0), max(b[2], x1), max(b[3], y1))
    # 文字範囲どうしの重なり
    keys = list(ext)
    for a in range(len(keys)):
        for b in range(a + 1, len(keys)):
            ia = inter(ext[keys[a]], ext[keys[b]])
            if ia and ia[0] > 0.03 and ia[1] > 0.03:
                ea, eb = texts[keys[a]], texts[keys[b]]
                kind = "番号が文字を隠す" if "overlay" in (ea["kind"], eb["kind"]) else "文字の重なり"
                issues.append((sn, kind, f"「{ea.get('str','')[:18]}」×「{eb.get('str','')[:18]}」"))
    # 配置の重なり（図形・画像・実際の文字範囲）
    solids = [(box(e), e["kind"], "") for e in els if e["kind"] in ("shape", "image")]
    solids += [(ext[i], "text", texts[i].get("str", "")[:18]) for i in ext if texts[i]["kind"] == "text"]
    for a in range(len(solids)):
        for b in range(a + 1, len(solids)):
            A, B = solids[a], solids[b]
            if A[1] == "text" and B[1] == "text": continue
            ii = inter(A[0], B[0])
            if not ii or ii[0] < 0.05 or ii[1] < 0.05: continue
            if contains(A[0], B[0]) or contains(B[0], A[0]): continue
            issues.append((sn, "配置の重なり", f"{A[1]}{A[2] and '「'+A[2]+'」'} × {B[1]}{B[2] and '「'+B[2]+'」'}  {[round(v,2) for v in A[0]]} / {[round(v,2) for v in B[0]]}"))
    # フッター
    body = [ext[i] for i in ext if texts[i]["y"] < 6.99] + [box(e) for e in els if e["kind"] in ("shape", "image") and e["y"] < 6.99]
    if body:
        low = max(b[3] for b in body)
        if low > 6.97:
            issues.append((sn, "フッター衝突", f"本文の下端 {low:.2f}in > 6.97in"))

for i in issues: print(f"p{i[0]:>2} [{i[1]}] {i[2]}")
print(f"--- {len(pages)}ページ / 指摘 {len(issues)}件")
