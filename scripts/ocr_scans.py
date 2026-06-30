"""
扫描件 OCR 匹配工具
用法: python ocr_scans.py <工作目录>

两阶段 OCR：先快速 psm=6 扫全部（~1分），只对未命中做深搜。
通过发票号精确匹配 + 金额兜底 + 关键词分类 + 消元法，
自动生成 scan_mapping.json。

依赖: pytesseract Pillow
Tesseract: https://github.com/UB-Mannheim/tesseract/wiki (Windows 安装包)
"""
import csv, json, os, re, sys
from pathlib import Path

import pytesseract
from PIL import Image, ImageOps, ImageFilter

from _ocr_common import configure_tesseract, MATCH_TOLERANCE

configure_tesseract()

WORK = Path(sys.argv[1]).resolve()
SCAN_DIR = WORK / '扫描件'


def load_rows():
    rows = {}
    with open(WORK / '数据.csv', encoding='utf-8-sig') as f:
        for i, r in enumerate(csv.reader(f)):
            if not r or not any(r):
                continue
            rows[3 + i] = {
                'inv': r[10] or '',
                'amt': float(r[5]),
                'cat': classify(r[8], r[11]),
            }
    return rows


def classify(cat, desc):
    if cat == '住宿费':
        return '住宿'
    if any(k in desc for k in ['高铁', '二等座']):
        return '高铁'
    if any(k in desc for k in ['打车', '出行', '滴滴', 'T3', '如祺', '风韵', '曹操', '享道', '高德']):
        return '打车'
    return '其他'


def ocr_one(img, psm_list, upscale=1, invert=False):
    """单张图片跑 Tesseract，返回 (inv_nums, amounts, raw_text)"""
    w, h = img.size
    if upscale > 1:
        img = img.resize((w * upscale, h * upscale), Image.LANCZOS)
        img = ImageOps.autocontrast(img, cutoff=5)
        img = img.filter(ImageFilter.SHARPEN)

    all_text = []
    for psm in psm_list:
        im = ImageOps.invert(img) if invert else img
        try:
            t = pytesseract.image_to_string(im, lang='chi_sim+eng', config=f'--psm {psm}')
            all_text.append(t)
        except Exception:
            pass
    full = '\n'.join(all_text)
    normalized = re.sub(r'(\d)\s+(\d)', r'\1\2', full)

    inv_nums = set()
    for m in re.findall(r'\b(\d{17,20})\b', normalized):
        inv_nums.add(m)
    for m in re.findall(r'\b(\d{8,16})\b', normalized):
        inv_nums.add(m)

    amounts = set()
    for m in re.findall(r'[¥￥]\s*(\d+\.?\d*)', full):
        try:
            v = float(m)
            if 1 < v < 9999: amounts.add(v)
        except ValueError: pass
    for m in re.findall(r'(?<![.\d])(\d+\.\d{2})(?!\d)', full):
        try:
            v = float(m)
            if 1 < v < 9999: amounts.add(v)
        except ValueError: pass

    return sorted(inv_nums, key=len, reverse=True), sorted(amounts), full


def match_by_invoice(inv_nums, rows):
    """发票号精确匹配 → (row or None)"""
    for num in inv_nums:
        for row, info in rows.items():
            expected = info['inv']
            if not expected:
                continue
            if num == expected:
                return row
            if len(num) >= 15 and len(expected) >= 15 and num[-15:] == expected[-15:]:
                return row
    return None


def match_by_amount(amounts, rows):
    """金额最近邻匹配 → (row, dist) or (None, 999)"""
    best_row, best_dist = None, 999
    for amt in amounts:
        for row, info in rows.items():
            d = abs(amt - info['amt'])
            if d < best_dist and d < 3:
                best_row, best_dist = row, d
    return best_row, best_dist


def classify_scan(text, w, h):
    trip_kw = ['DIDI TRAVEL', 'AMAP ITINERARY', 'TRIP TABLE',
               '行程', '上车', '下车', '公里', '分钟', '出行']
    inv_kw = ['发票', '发票代码', '发票号码', '增值税', '国家税务总局']
    train_kw = ['二等座', '火车', '车次', '席别']
    hotel_kw = ['住宿', '酒店', '宾馆', '房费']

    if any(k in text for k in train_kw): return '高铁票'
    if any(k in text for k in hotel_kw): return '住宿票'
    if any(k in text for k in trip_kw): return '行程单'
    if any(k in text for k in inv_kw): return '发票'
    return '发票' if w > h else '?'


def main():
    if not SCAN_DIR.exists():
        print(f'扫描件目录不存在: {SCAN_DIR}')
        sys.exit(1)

    rows = load_rows()
    files = sorted([f for f in os.listdir(SCAN_DIR)
                    if f.lower().endswith(('.jpg', '.jpeg', '.png'))
                    and not f.startswith('_')])
    print(f'数据行: {len(rows)}  扫描件: {len(files)} 张\n')

    scan_map = {}
    matched = set()
    unmatched = []

    # ── Phase 1: 快速扫（1x, psm=6, 不反色）──
    print('── Phase 1: 快速 OCR (psm=6) ──')
    for fn in files:
        img = Image.open(SCAN_DIR / fn).convert('L')
        w, h = img.size
        inv_nums, amounts, text = ocr_one(img, psm_list=[6], upscale=1, invert=False)

        row = match_by_invoice(inv_nums, rows)
        if row:
            scan_map.setdefault(row, []).append(fn)
            matched.add(fn)
            print(f'  {fn[:40]} [INV] → R{row}')
        else:
            row, dist = match_by_amount(amounts, rows)
            if row and dist <= 2:
                scan_map.setdefault(row, []).append(fn)
                matched.add(fn)
                print(f'  {fn[:40]} [AMT] → R{row} Δ{dist:.2f}')
            else:
                unmatched.append((fn, img, w, h, inv_nums, amounts, text))
                print(f'  {fn[:40]} [?] 待深搜 inv={inv_nums[:2]} amt={amounts[:3]}')

    # ── Phase 2: 深搜（只有未命中的，2x, psm=[3,6,7], 正常+反色）──
    if unmatched:
        print(f'\n── Phase 2: 深搜 {len(unmatched)} 张 ──')
        still_unmatched = []
        for fn, img, w, h, inv_nums, amounts, text in unmatched:
            # 深搜: 2x upscale + more PSMs
            inv_nums2, amounts2, text2 = ocr_one(img, psm_list=[3, 6, 7], upscale=2, invert=False)
            inv_nums3, amounts3, text3 = ocr_one(img, psm_list=[3, 6, 7], upscale=2, invert=True)

            all_inv = list(dict.fromkeys(inv_nums + inv_nums2 + inv_nums3))
            all_amt = sorted(set(amounts + amounts2 + amounts3))
            all_text = text + '\n' + text2 + '\n' + text3

            row = match_by_invoice(all_inv, rows)
            if row:
                scan_map.setdefault(row, []).append(fn)
                matched.add(fn)
                print(f'  {fn[:40]} [INV2] → R{row}')
                continue

            row, dist = match_by_amount(all_amt, rows)
            if row and dist <= 2:
                scan_map.setdefault(row, []).append(fn)
                matched.add(fn)
                print(f'  {fn[:40]} [AMT2] → R{row} Δ{dist:.2f}')
                continue

            still_unmatched.append((fn, w, h, all_text))

        unmatched = still_unmatched

    # ── Phase 3: 消元法 ──
    if unmatched:
        print(f'\n── Phase 3: 消元法 {len(unmatched)} 张 ──')
        typed = [(fn, classify_scan(text, w, h)) for fn, w, h, text in unmatched]

        for row in sorted(rows):
            info = rows[row]
            need_trip = 1 if info['cat'] == '打车' else 0
            existing = scan_map.get(row, [])
            have_inv = 1 if existing else 0
            have_trip = 1 if len(existing) >= 2 else 0

            if have_inv == 0:
                candidates = [t for t in typed if t[1] in ('发票', '高铁票', '住宿票', '?') and t[0] not in matched]
                if candidates:
                    pick = candidates[0]
                    scan_map.setdefault(row, []).append(pick[0])
                    matched.add(pick[0])
                    typed.remove(pick)
                    print(f'  {pick[0][:40]} [ELIM] → R{row} (发票)')

            if need_trip and have_trip == 0:
                # 行程单 OCR 不稳定，'发票' 类型的横版也可能是行程单
                candidates = [t for t in typed if t[0] not in matched]
                if candidates:
                    pick = candidates[0]
                    scan_map.setdefault(row, []).append(pick[0])
                    matched.add(pick[0])
                    typed.remove(pick)
                    print(f'  {pick[0][:40]} [ELIM] → R{row} (行程单)')

        remaining = [t for t in typed if t[0] not in matched]
        if remaining:
            print(f'  WARN: {len(remaining)} 张无法自动分配: {[t[0][:25] for t in remaining]}')

    # 确保所有行有键
    for r in rows:
        scan_map.setdefault(r, [])

    # 写入
    output = {str(k): v for k, v in sorted(scan_map.items())}
    with open(WORK / 'scan_mapping.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f'\n── 结果 ──')
    total = 0
    for r in sorted(scan_map, key=int):
        items = scan_map[r]
        info = rows.get(int(r), {})
        need = 2 if info.get('cat') == '打车' else 1
        ok = 'OK' if len(items) >= need else f'MISS {need - len(items)}'
        print(f'  R{r} [{info.get("cat", "?")}] ¥{info.get("amt", 0):.2f} {len(items)}/{need} {ok}')
        total += len(items)
    print(f'\n已分配: {total}/{len(files)}')
    print(f'Saved: {WORK / "scan_mapping.json"}')


if __name__ == '__main__':
    main()
