"""
付款截图 OCR 匹配工具
用法: python ocr_payments.py <工作目录>

遍历付款截图目录，OCR 识别付款金额，按金额匹配到数据行，
写入 payment_mapping.json。支持拼单（两笔小付组合为一笔报销）。

依赖: pytesseract Pillow
"""
import csv, json, os, re, sys
from itertools import combinations
from pathlib import Path

import pytesseract
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from _ocr_common import configure_tesseract, MATCH_TOLERANCE

configure_tesseract()

WORK = Path(sys.argv[1]).resolve()
PAY_DIR = WORK / '付款截图'


def load_rows():
    rows = {}
    with open(WORK / '数据.csv', encoding='utf-8-sig') as f:
        for i, r in enumerate(csv.reader(f)):
            if not r or not any(r):
                continue
            rows[3 + i] = float(r[5])
    return rows


def ocr_amount(path):
    """OCR 单张付款截图，返回提取到的所有金额"""
    img = Image.open(path).convert('L')
    w, h = img.size

    # 付款截图是竖版 1152×2560，金额在中间偏上
    # 裁剪金额区
    crop = img.crop((0, h // 3, w, h // 2))
    crop2 = crop.resize((crop.width * 2, crop.height * 2), Image.LANCZOS)
    crop2 = ImageOps.autocontrast(crop2, cutoff=5)
    crop2 = crop2.filter(ImageFilter.SHARPEN)

    amounts = set()
    for inv in [False, True]:
        im = ImageOps.invert(crop2) if inv else crop2
        for psm in [6, 3]:
            try:
                t = pytesseract.image_to_string(im, lang='chi_sim+eng', config=f'--psm {psm}')
                for m in re.findall(r'[¥￥]\s*(\d+\.?\d*)', t):
                    v = float(m)
                    if 1 < v < 9999:
                        amounts.add(v)
                for m in re.findall(r'(?<![.\d])(\d+\.\d{2})(?!\d)', t):
                    v = float(m)
                    if 1 < v < 9999:
                        amounts.add(v)
            except Exception:
                pass

    return sorted(amounts, reverse=True)


def match_payments(files, rows):
    """
    两阶段匹配：
    1. 单张精确匹配（精确 ± 容差）
    2. 剩余双张组合匹配（拼单）
    """
    # Step 1: OCR all files
    ocr_results = {}
    for f in files:
        amounts = ocr_amount(PAY_DIR / f)
        ocr_results[f] = amounts
        main_amt = amounts[0] if amounts else 0
        print(f'  {f[:40]} OCR={amounts[:4]}')

    # Step 2: Single match
    mapping = {}
    matched_files = set()
    unmatched_ocr = {}

    for f, amounts in ocr_results.items():
        best_row, best_dist = None, 999
        best_amt = 0
        for amt in amounts:
            for row, target in rows.items():
                if row in mapping:
                    continue  # row already covered
                d = abs(amt - target)
                if d < best_dist and d <= MATCH_TOLERANCE:
                    best_row, best_dist, best_amt = row, d, amt
        if best_row:
            mapping[best_row] = [f]
            matched_files.add(f)
            print(f'  → R{best_row} ¥{rows[best_row]:.2f} (OCR={best_amt:.2f} Δ{best_dist:.2f})')
        else:
            unmatched_ocr[f] = amounts

    # Step 3: Combination match (拼单)
    unmatched = [f for f in files if f not in matched_files]
    if len(unmatched) >= 2:
        uncovered = {r: t for r, t in rows.items() if r not in mapping}
        for r, target in uncovered.items():
            best_pair, best_dist = None, 999
            for f1, f2 in combinations(unmatched, 2):
                a1 = ocr_results[f1][0] if ocr_results[f1] else 0
                a2 = ocr_results[f2][0] if ocr_results[f2] else 0
                d = abs(a1 + a2 - target)
                if d < best_dist and d <= MATCH_TOLERANCE:
                    best_pair, best_dist = (f1, f2), d
            if best_pair:
                mapping[r] = list(best_pair)
                matched_files.update(best_pair)
                a1 = ocr_results[best_pair[0]][0] if ocr_results[best_pair[0]] else 0
                a2 = ocr_results[best_pair[1]][0] if ocr_results[best_pair[1]] else 0
                print(f'  → R{r} ¥{target:.2f} 拼单: {a1:.2f}+{a2:.2f}={a1 + a2:.2f}')

    return mapping


def main():
    if not PAY_DIR.exists():
        print(f'付款截图目录不存在: {PAY_DIR}')
        sys.exit(1)

    files = sorted([f for f in os.listdir(PAY_DIR)
                    if f.lower().endswith(('.jpg', '.jpeg', '.png'))
                    and not f.startswith('_')])

    if not files:
        print('无付款截图文件')
        return

    rows = load_rows()
    print(f'数据行: {len(rows)}')
    print(f'付款截图: {len(files)} 张')
    print()

    # 检查已有映射，跳过已匹配的
    existing = {}
    existing_path = WORK / 'payment_mapping.json'
    if existing_path.exists():
        with open(existing_path, encoding='utf-8') as f:
            existing = json.load(f)
        if existing:
            print(f'已有映射: {len(existing)} 行，跳过已匹配')

    # 只 OCR 新增文件
    in_use = {pf for v in existing.values() for pf in v} if existing else set()
    new_files = [f for f in files if f not in in_use]

    if not new_files:
        print('无新增付款截图')
        return

    print(f'新增: {len(new_files)} 张')
    print()
    print('── OCR 识别 ──')

    new_mapping = match_payments(new_files, rows)

    # 合并
    for r, files_list in new_mapping.items():
        r_str = str(r)
        existing[r_str] = files_list

    # 确保所有行都有键
    for r in rows:
        existing.setdefault(str(r), [])

    with open(existing_path, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    matched = sum(1 for v in existing.values() if v)
    print(f'\n已覆盖: {matched}/{len(rows)} 行')
    print(f'Saved: {existing_path}')


if __name__ == '__main__':
    main()
