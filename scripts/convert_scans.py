"""
扫描件格式转换工具（Read 工具兼容）
用法: python convert_scans.py <工作目录>

手机相机直出 JPEG（Xiaomi 等）不被 Read 工具支持。
此脚本将扫描件嵌入 1152×2560 白底画布 + 复用付款截图的 EXIF/ICC，
使其可被 Read 工具渲染。

输出到 付款截图/_scans/ 目录。
"""
import io, os, sys
from pathlib import Path

from PIL import Image

WORK = Path(sys.argv[1]).resolve()
SCAN_DIR = WORK / '扫描件'
PAY_DIR = WORK / '付款截图'
OUT_DIR = PAY_DIR / '_scans'


def main():
    if not SCAN_DIR.exists():
        print(f'扫描件目录不存在: {SCAN_DIR}')
        sys.exit(1)

    # 找一张付款截图作为 EXIF/ICC 来源
    pay_files = sorted([f for f in os.listdir(PAY_DIR)
                        if f.lower().endswith('.jpg') and not f.startswith('_')])
    if not pay_files:
        print('无付款截图可用作 EXIF/ICC 来源')
        sys.exit(1)

    ref = Image.open(PAY_DIR / pay_files[0])
    exif = ref.info.get('exif')
    icc = ref.info.get('icc_profile')
    print(f'参考图: {pay_files[0]} exif={exif is not None} icc={icc is not None}')

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    scan_files = sorted([f for f in os.listdir(SCAN_DIR)
                         if f.lower().endswith('.jpg') and not f.startswith('_')])

    for fn in scan_files:
        scan = Image.open(SCAN_DIR / fn)
        w, h = scan.size
        sw = 1152
        sh = int(h * 1152 / w)
        scan2 = scan.resize((sw, sh), Image.LANCZOS).convert('RGB')
        canvas = Image.new('RGB', (1152, 2560), (255, 255, 255))
        y = (2560 - sh) // 2
        canvas.paste(scan2, (0, y))

        buf = io.BytesIO()
        canvas.save(buf, 'JPEG', quality=92, exif=exif, icc_profile=icc)

        out_path = OUT_DIR / fn
        with open(out_path, 'wb') as fh:
            fh.write(buf.getvalue())

        kb = os.path.getsize(out_path) // 1024
        print(f'  {fn[:35]} {w}x{h} → {canvas.size} {kb}KB')

    print(f'\nDone: {len(scan_files)} 张 → {OUT_DIR}')


if __name__ == '__main__':
    main()
