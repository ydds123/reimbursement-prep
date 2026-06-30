# 踩坑记录

2026-06-24 积累，按发现顺序排列。

1. **Read 工具整体不可靠**：不同会话中表现不一致——有时能渲染白底画布转换后的图片，有时对所有图片（包括 100×100 测试图）全部返回 `[Unsupported Image]`。**不能作为唯一依赖路径**，必须有 Tesseract 兜底。

2. **DeepSeek 代理吞图**：Python SDK 经本地代理 → DeepSeek 视觉模型，首张准后续全错（把发票当化妆品、火车票当存折）。原因是代理对多张/大图有截断。不可作为可靠方案。

3. **Tesseract 可用且有条件准确**：之前认为 Tesseract 不准（54.30→4.30），是因为在低分辨率压缩图上跑。在 2560×1920 原始扫描件上做 2x 放大 + autocontrast + sharpen + 多 PSM 后，**发票号（纯数字）识别率约 70%**。金额识别仍然不准，但发票号是更好的锚点。

4. **不用 PaddleOCR**：v3 在 Windows 上有 ONEDNN 兼容性问题（`NotImplementedError: ConvertPirAttribute2RuntimeAttribute`），安装后无法运行。

5. **PDF 可能在子目录**：用户可能把 PDF 放进 `内容/` 等子目录，`glob('*.pdf')` 会漏掉，必须用 `rglob('**/*.pdf')`。

6. **永远不要自动删文件**：扫描件有备份/重拍/多余是正常的，自检改成警告而非 ERROR。用户自行决定是否清理。

7. **API 密钥**：从 `~/.claude/.credentials.json` 动态读取，绝不硬编码到脚本里。直连 Anthropic API 在国内被墙，走代理才是可用路径。

8. **拼单**：一笔报销（¥61.00）由两笔付款（¥48.19+¥13.81）组成，需手动在 `payment_mapping.json` 和 `scan_mapping.json` 中标注为数组。

9. **发票号精确匹配是王牌策略**：扫描件 OCR 的核心价值不是金额而是发票号。17-20 位发票号是全局唯一定位符，精确匹配即可确定归属。金额作为兜底（±3 元容差）。对于行程单（无发票号），靠关键词（DIDI TRAVEL, AMAP ITINERARY）+ 消元法分配。

10. **扫描件目录要非递归搜索**：`rglob('*')` 会把 `_scans/`、`_crops/` 等临时子目录的文件全当孤儿报 WARN。`build.py` 的 `file_audit` 用 `glob('*')` 仅搜索一级目录。

11. **文件名格式应为 `报销-YY-MM.xlsx`**：不包含姓名，从 CSV 的 YYYY-MM 列提取。原 `报销-{NAME}{YY}-{MM}` 会拼成 `吴松海26-06`。
