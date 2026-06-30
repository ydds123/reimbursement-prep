# 常见问题

| 问题 | 原因 | 解决 |
|------|------|------|
| `PermissionError` | Excel 窗口打开着目标文件 | 关掉 Excel 再跑；脚本自动 fallback 到 `-NEW.xlsx` |
| 付款截图无法自动匹配 | Read 工具不可用或 OCR 金额不清晰 | 优先 Tesseract 裁剪金额区 OCR；Read 工具可用时辅助验证 |
| 某行没图 | `row_pdfs` 漏了 | 自检会报 ERROR 阻止保存 |
| 孤儿 PDF | 下了发票但没配行 | 自检会报；删掉无用 PDF 或用到的行补上映射 |
| Read 看不到扫描件 | 手机相机 JPEG 不被 Read 工具解码器支持 | 用 Pillow 嵌入 1152×2560 白底画布 + 复用付款截图 EXIF/ICC |
| 多余的扫描件 | 备份/重拍 | **不要删**，只是不映射到行，检查时显示为警告 |
| Tesseract 未安装 | 首次使用 | `pip install pytesseract` + 下载安装 Tesseract OCR Windows 版 |
| OCR 检出 0 张 | 预处理不够 | 确保 2x 放大、自动对比度、多 PSM 模式、数字空格归一化 |
