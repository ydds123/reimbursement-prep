"""
报销 Excel 生成工具
用法: python build.py <工作目录>

职责：读数据 → 配 PDF → 嵌图片 → 检查 → 输出 Excel
Claude 只需准备 数据.csv、payment_mapping.json、scan_mapping.json，
不需改动本脚本。
"""
import csv, io, json, os, re, sys
from copy import copy
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

# ═══════════════════════════════════════════════════════════════════════
# 配置
# ═══════════════════════════════════════════════════════════════════════

WORK = Path(sys.argv[1]).resolve()
SKILL_DIR = Path(__file__).resolve().parent
TEMPLATE = SKILL_DIR / '模板.xlsx'
DPI = 600
DISP_SINGLE = 200          # 单 PDF 显示宽度
DISP_PAIR   = 155          # 双 PDF 显示宽度
DISP_EXTRA  = 135          # 付款截图 / 扫描件显示宽度
GAP = 10

# 文件材料要求以 references/document-checklist.md 为单一事实来源。
# 此字典仅用于运行时自检，修改材料要求时请同步更新 document-checklist.md。
EXPECTED = {
    '高铁': {'发票': 1, '扫描件': 1, '付款截图': 1},
    '住宿': {'发票': 1, '扫描件': 1, '付款截图': 1},
    '打车': {'发票': 1, '行程单': 1, '扫描件': 1, '付款截图': 1},
    '其他': {'发票': 1, '扫描件': 1, '付款截图': 1},
}

# ═══════════════════════════════════════════════════════════════════════
# 1. 加载数据
# ═══════════════════════════════════════════════════════════════════════

def load_csv():
    p = WORK / '数据.csv'
    if not p.exists():
        die(f'{p} 不存在')
    rows = []
    with open(p, encoding='utf-8-sig') as f:
        rdr = csv.reader(f)
        for linenum, r in enumerate(rdr, start=1):
            if not r or not any(r):
                continue
            if len(r) != 13:
                die(f'数据.csv 第 {linenum} 行有 {len(r)} 列（应为 13 列），请检查 CSV 文件')
            try:
                r[5] = float(r[5])
            except (ValueError, IndexError):
                die(f'数据.csv 第 {linenum} 行金额列（第6列）无法转为数字: "{r[5]}"')
            if not re.match(r'^\d{4}-\d{2}$', str(r[6])):
                die(f'数据.csv 第 {linenum} 行月份格式错误 "{r[6]}"（应为 YYYY-MM）')
            if not str(r[12]).strip():
                die(f'数据.csv 第 {linenum} 行排序键（第13列）为空，每行必须填写排序键')
            if r[9] == '':
                r[9] = None
            rows.append(r)
    if not rows:
        die('数据.csv 没有有效数据行')
    rows.sort(key=lambda r: r[12])
    return rows

# ═══════════════════════════════════════════════════════════════════════
# 2. PDF 自动匹配
# ═══════════════════════════════════════════════════════════════════════

def match_pdfs(rows):
    global pdf_paths
    pdfs = sorted(Path(WORK).rglob('*.pdf'))
    pdf_paths = {p.name: p for p in pdfs}
    mapping = {}
    assigned = set()

    # ── 高铁票：纯数字文件名 → K 列票号 ──
    for pdf in pdfs:
        m = re.match(r'^(\d+)\.pdf$', pdf.name)
        if m:
            ticket = m.group(1)
            for i, rd in enumerate(rows):
                if rd[10] == ticket:
                    mapping[3 + i] = [pdf.name]
                    assigned.add(pdf.name)
                    break

    # ── 住宿 dzfp_ → K 列发票号 ──
    for pdf in pdfs:
        if pdf.name.startswith('dzfp_') and pdf.name not in assigned:
            m = re.match(r'dzfp_(\d+)', pdf.name)
            if m:
                invoice = m.group(1)
                for i, rd in enumerate(rows):
                    if rd[10] == invoice:
                        mapping[3 + i] = [pdf.name]
                        assigned.add(pdf.name)
                        break

    # ── 高德聚合（先于滴滴，因为文件名含精确平台+金额）──
    gaode_inv  = [p for p in pdfs if '高德打车电子发票' in p.name and p.name not in assigned]
    gaode_trv  = [p for p in pdfs if '高德打车电子行程单' in p.name and p.name not in assigned]
    PLATFORMS = ['T3', '如祺', '风韵', '曹操', '享道']

    for inv in gaode_inv:
        plat = next((k for k in PLATFORMS if k in inv.name), None)
        trv = next((t for t in gaode_trv if plat and plat in t.name and t.name not in assigned), None)
        m_amt = re.search(r'(\d+\.\d{2})', inv.name)
        if m_amt and plat:
            amt = float(m_amt.group(1))
            for i, rd in enumerate(rows):
                r = 3 + i
                if r in mapping:
                    continue
                if abs(float(rd[5]) - amt) < 2 and plat in rd[11]:
                    mapping[r] = [inv.name] + ([trv.name] if trv else [])
                    assigned.update([inv.name] + ([trv.name] if trv else []))
                    break

    # ── 滴滴：无法通过文件名确定归属，收集待人工确认 ──
    didi_inv  = sorted([p for p in pdfs if '滴滴电子发票' in p.name and p.name not in assigned])
    didi_trv  = sorted([p for p in pdfs if '滴滴出行行程报销单' in p.name and p.name not in assigned])
    didi_rows = sorted([3 + i for i in range(len(rows))
                        if (3 + i) not in mapping
                        and any(k in rows[i][11] for k in ['滴滴', '打车'])],
                       key=lambda r: rows[r - 3][12])

    if didi_inv or didi_trv:
        review = {
            "_note": "以下滴滴 PDF 无法通过文件名规则确定归属行，请人工建立映射后写入 pdf_mapping.json",
            "unmatched_didi_invoices": [p.name for p in didi_inv],
            "unmatched_didi_travels": [p.name for p in didi_trv],
            "unmatched_taxi_rows": [{"row": r, "amount": rows[r - 3][5], "desc": rows[r - 3][11]}
                                    for r in didi_rows],
        }
        review_path = WORK / 'pdf_mapping_review.json'
        with open(review_path, 'w', encoding='utf-8') as f:
            json.dump(review, f, ensure_ascii=False, indent=2)
        print(f'  ⚠ 滴滴 PDF {len(didi_inv)} 张发票 + {len(didi_trv)} 张行程单 无法自动匹配')
        print(f'  → 已生成 {review_path}，请人工确认后写入 pdf_mapping.json')

    unmatched = [p.name for p in pdfs if p.name not in assigned]
    unmapped  = [3 + i for i in range(len(rows)) if (3 + i) not in mapping]
    return mapping, unmatched, unmapped

# ═══════════════════════════════════════════════════════════════════════
# 3. JSON 映射
# ═══════════════════════════════════════════════════════════════════════

def load_json_map(name):
    p = WORK / name
    if p.exists():
        with open(p, encoding='utf-8') as f:
            return {int(k): v for k, v in json.load(f).items()}
    return {}

# ═══════════════════════════════════════════════════════════════════════
# 4. 填充数据
# ═══════════════════════════════════════════════════════════════════════

def fill_sheet(ws, rows):
    ref_row = 7
    for i, rd in enumerate(rows):
        r = 3 + i
        for j, val in enumerate(rd[:12]):
            src = ws.cell(row=ref_row, column=j + 1)
            dst = ws.cell(row=r, column=j + 1)
            dst.value = val
            if src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = copy(src.number_format)
                dst.alignment = copy(src.alignment)

    n = len(rows)
    s, b, p = 3 + n, 4 + n, 5 + n  # sum, backup, pay rows

    ws.cell(row=s, column=1).value = '合计'
    ws.cell(row=s, column=6).value = f'=SUM(F3:F{2 + n})'
    ws.cell(row=s, column=7).value = '/'
    ws.cell(row=s, column=12).value = '/'
    ws.cell(row=b, column=1).value = '本次核销备用金金额'
    ws.cell(row=b, column=7).value = '之前领用过备用金的，本次发生的费用冲抵备用金的金额'
    ws.cell(row=p, column=1).value = '核销完备用金后需支付金额'
    ws.cell(row=p, column=6).value = f'=F{s}-F{b}'
    ws.cell(row=p, column=7).value = '/'

    for a, b_col, c, d_col in [(s, 'E', s, 'K'), (b, 'E', b, 'K'), (p, 'E', p, 'K')]:
        ws.merge_cells(f'A{a}:{b_col}{a}')
        ws.merge_cells(f'G{a}:{d_col}{a}')
    ws.merge_cells(f'L{s}:L{p}')

    for r in range(3, 3 + n):
        c = ws.cell(row=r, column=6)
        if c.value is not None and not str(c.value).startswith('='):
            c.number_format = '0.00'
    for r in range(p + 1, 50):
        for c in range(1, 13):
            ws.cell(row=r, column=c).value = None

    return s, b, p

# ═══════════════════════════════════════════════════════════════════════
# 5. 嵌入图片
# ═══════════════════════════════════════════════════════════════════════

# ── PDF lookup (populated in main) ──
pdf_paths = {}

def render_pdf(name):
    path = pdf_paths.get(name, WORK / name)
    doc = pdfium.PdfDocument(str(path))
    bmp = doc[0].render(scale=DPI / 72)
    img = bmp.to_pil()
    doc.close()
    return img

def embed_images(ws, pdf_map, pay_map, scan_map):
    pay_dir = WORK / '付款截图'
    scan_dir = WORK / '扫描件'

    for row_num in sorted(set(pdf_map) | set(pay_map) | set(scan_map)):
        pdfs  = pdf_map.get(row_num, [])
        pays  = pay_map.get(row_num, [])
        scans = scan_map.get(row_num, [])

        items = []
        if len(pdfs) == 1:
            items.append((render_pdf(pdfs[0]), DISP_SINGLE))
        elif len(pdfs) >= 2:
            for pf in pdfs:
                items.append((render_pdf(pf), DISP_PAIR))
        for pf in pays:
            items.append((Image.open(str(pay_dir / pf)), DISP_EXTRA))
        for sf in scans:
            items.append((Image.open(str(scan_dir / sf)), DISP_EXTRA))

        if not items:
            continue

        x_off, max_dh = 0, 0
        for img, dw in items:
            dh = int(img.height * dw / img.width)
            max_dh = max(max_dh, dh)
            buf = io.BytesIO()
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img.save(buf, format='JPEG', quality=92)
            buf.seek(0)
            xl = XLImage(buf)
            xl.width = dw
            xl.height = dh
            if x_off == 0:
                xl.anchor = f'J{row_num}'
            else:
                anchor = OneCellAnchor(
                    _from=AnchorMarker(col=9, colOff=pixels_to_EMU(x_off), row=row_num - 1, rowOff=0),
                    ext=XDRPositiveSize2D(pixels_to_EMU(dw), pixels_to_EMU(dh)))
                xl.anchor = anchor
            ws.add_image(xl)
            x_off += dw + GAP

        ws.row_dimensions[row_num].height = max_dh * 0.75 + 25
        print(f'  R{row_num}: {"+".join(str(dw) for _,dw in items)} h={max_dh}')

# ═══════════════════════════════════════════════════════════════════════
# 6. 文档完整性检查
# ═══════════════════════════════════════════════════════════════════════

def classify(rd):
    cat, desc = rd[8], rd[11]
    if cat == '住宿费':
        return '住宿'
    if any(k in desc for k in ['高铁', '二等座']):
        return '高铁'
    if any(k in desc for k in ['打车', '出行', '滴滴', 'T3', '如祺', '风韵', '曹操', '享道', '高德']):
        return '打车'
    return '其他'

def pdf_kind(fname):
    if '电子发票' in fname or fname.startswith('dzfp_'):
        return '发票'
    if '行程单' in fname or '行程报销单' in fname:
        return '行程单'
    return '发票'

def check_completeness(rows, pdf_map, pay_map, scan_map):
    print('\n── 文档完整性检查 ──')
    header = f'  {"行":4s} {"金额":>7s} {"项目":24s} {"类型":4s} {"发票":5s} {"行程单":6s} {"付款截图":7s} {"扫描件":6s}'
    print(header)
    print('  ' + '-' * (len(header) - 2))
    for i, rd in enumerate(rows):
        r = 3 + i
        ctype = classify(rd)
        exp = EXPECTED[ctype]

        present = {}
        for pf in pdf_map.get(r, []):
            k = pdf_kind(pf)
            present[k] = present.get(k, 0) + 1
        present['付款截图'] = present.get('付款截图', 0) + len(pay_map.get(r, []))
        present['扫描件'] = present.get('扫描件', 0) + len(scan_map.get(r, []))

        def cell(k):
            want = exp.get(k, 0)
            if want == 0:
                return '  —  '
            got = present.get(k, 0)
            return '  ✓  ' if got >= want else f' \033[33m✗\033[0m  '

        line = f'  R{r:<2d} ¥{rd[5]:>6.2f} {rd[11][:24]:24s} [{ctype:4s}] {cell("发票")}  {cell("行程单")}  {cell("付款截图")}  {cell("扫描件")}'
        print(line)
    print()

# ═══════════════════════════════════════════════════════════════════════
# 7. 文件覆盖自检
# ═══════════════════════════════════════════════════════════════════════

def _row_has_images(row_num, pdf_map, pay_map, scan_map):
    """某行是否至少有一张真实图片。空数组不算。"""
    return (
        len(pdf_map.get(row_num, [])) > 0
        or len(pay_map.get(row_num, [])) > 0
        or len(scan_map.get(row_num, [])) > 0
    )


def file_audit(rows, pdf_map, pay_map, scan_map, unmatched_pdfs):
    errors = []
    warnings = []

    data_rows = set(range(3, 3 + len(rows)))
    imaged = {r for r in data_rows if _row_has_images(r, pdf_map, pay_map, scan_map)}
    for r in sorted(data_rows - imaged):
        errors.append(f'R{r} 无任何图片')

    for f in unmatched_pdfs:
        errors.append(f'孤儿 PDF: {f}')

    pdf_in_use = {pf for sub in pdf_map.values() for pf in sub}
    for pf in sorted(pdf_in_use):
        if pf not in pdf_paths:
            errors.append(f'PDF 不存在: {pf}')

    # 付款截图/扫描件孤儿不阻止，仅警告
    for subdir, label, the_map in [
        ('付款截图', '付款截图', pay_map),
        ('扫描件', '扫描件', scan_map),
    ]:
        d = WORK / subdir
        on_disk = {f.name for f in d.glob('*') if f.suffix.lower() in ('.jpg', '.png', '.jpeg')} if d.exists() else set()
        in_use  = {pf for sub in the_map.values() for pf in sub}
        for f in sorted(on_disk - in_use):
            if not f.startswith('_'):  # skip temp/preview files
                warnings.append(f'孤儿 {label}: {f}')
        for f in sorted(in_use - on_disk):
            errors.append(f'{label} 不存在: {f}')

    if warnings:
        for w in warnings:
            print(f'  \033[33mWARN: {w}\033[0m')

    if errors:
        for e in errors:
            print(f'  ERROR: {e}')
        die('自检未通过，已阻止保存。')

    print(f'  self-check passed')
    print(f'  行: {len(imaged)}/{len(data_rows)}  '
          f'PDF: {len(pdf_in_use)}  '
          f'付款截图: {sum(len(v) for v in pay_map.values())}  '
          f'扫描件: {sum(len(v) for v in scan_map.values())}')

# ═══════════════════════════════════════════════════════════════════════
# 主流程
# ═══════════════════════════════════════════════════════════════════════

def die(msg):
    print(f'\n  \033[31m{msg}\033[0m')
    sys.exit(1)

def main():
    print(f'工作目录: {WORK}')

    # 1. 数据
    rows = load_csv()
    year_month = rows[0][6]
    # 校验月份格式
    for i, rd in enumerate(rows):
        if not re.match(r'^\d{4}-\d{2}$', rd[6]):
            print(f'  WARN: R{3+i} 月份格式异常 "{rd[6]}"，应为 YYYY-MM')
    print(f'  数据行: {len(rows)}  文件名月份: {year_month}')

    # 2. PDF
    pdf_map, unmatched_pdfs, unmapped_rows = match_pdfs(rows)
    if unmatched_pdfs:
        print(f'  未匹配 PDF: {unmatched_pdfs}')
    if unmapped_rows:
        print(f'  无 PDF 的行: R{unmapped_rows}')

    # 3. JSON 覆盖（自动匹配之后，允许手动修正）
    pdf_map.update(load_json_map('pdf_mapping.json'))
    pay_map  = load_json_map('payment_mapping.json')
    scan_map = load_json_map('scan_mapping.json')

    # 4. 文档检查（保存前先看一眼缺什么）
    check_completeness(rows, pdf_map, pay_map, scan_map)

    # 5. 加载模板
    if not TEMPLATE.exists():
        die(f'模板不存在: {TEMPLATE}')
    wb = openpyxl.load_workbook(str(TEMPLATE))
    ws = wb.active

    # 清空旧数据和图片
    for mc in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mc))
    for r in range(3, 50):
        for c in range(1, 13):
            ws.cell(row=r, column=c).value = None
    ws._images = []
    ws._drawing = None

    # 6. 填数据
    fill_sheet(ws, rows)

    # 7. 嵌图片
    (WORK / '付款截图').mkdir(exist_ok=True)
    (WORK / '扫描件').mkdir(exist_ok=True)
    embed_images(ws, pdf_map, pay_map, scan_map)

    # 8. 自检
    file_audit(rows, pdf_map, pay_map, scan_map, unmatched_pdfs)

    # 9. 保存
    dst = WORK / f'报销-{year_month[2:4]}-{year_month[5:]}.xlsx'
    try:
        wb.save(str(dst))
        print(f'\nSaved: {dst}')
    except PermissionError:
        fallback = WORK / f'报销-{year_month[2:4]}-{year_month[5:]}-NEW.xlsx'
        wb.save(str(fallback))
        print(f'\nTarget locked. Saved: {fallback}')

if __name__ == '__main__':
    main()
