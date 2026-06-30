---
name: reimbursement-prep
description: 报销发票材料准备 —— 从模板 Excel、费用数据、发票 PDF、付款截图、扫描件，生成带嵌入图片的完整报销表，含文档完整性检查。触发词：报销、发票材料、报销Excel、费用报销、差旅报销、月度报销。不触发：仅查询报销政策、仅询问发票格式、研发文档撰写。近邻区分：prd-skills（文档撰写）vs 本 skill（报销材料生成）。
metadata:
  author: rd001
  mode: Production
  version: 1.3.0
  updated: 2026-06-24
---

# 报销发票材料准备

## 概述

将月度项目差旅费的电子发票 PDF、行程单、付款截图、发票扫描件整合为可直接提交财务的 Excel。**Tesseract OCR 发票号精确匹配**（准确率 ~70%）为扫描件主路径，Read 工具 / AI 视觉为辅助。可批量处理 200+ 张。

## 前置条件

- Python: `pypdfium2` `Pillow` `openpyxl` `pytesseract`
- Tesseract OCR: [Windows 安装包](https://github.com/UB-Mannheim/tesseract/wiki)
- 模板 Excel：`scripts/模板.xlsx`

## 目录规范

```
{工作目录}/
├── 数据.csv                     # 费用数据（12 列无表头）
├── 报销-{YY}-{MM}.xlsx          # 输出
├── 内容/ 或 *.pdf               # 电子发票/行程单 PDF
├── payment_mapping.json         # 付款截图→行号
├── scan_mapping.json            # 扫描件→行号
├── 付款截图/                    # 微信/支付宝 .jpg
└── 扫描件/                      # 手机相机直出 .jpg
```

## 工作流

### 阶段 1：环境确认

1. Python 依赖：`pypdfium2 openpyxl Pillow pytesseract`
2. 确认工作目录
3. 如为月度复用：从上期工作目录复制 `payment_mapping.json` 和 `scan_mapping.json` 作为起点，跳过阶段 4-5 中对已映射文件的重复识别

### 阶段 2：数据准备

`数据.csv` 13 列无表头：
```
申请人, 使用人, 区域, 报销部门, 建设单位, 金额, YYYY-MM, 一级分类, 二级分类, , 发票号码, 报销明细, 排序键
```
- 第 10 列留空；排序键形如 `2026-06-15-a`，每行必填

### 阶段 3：发票 PDF 匹配

`build.py` 按文件名模式自动匹配，无需干预：

| 类型 | 规则 |
|------|------|
| 高铁 | `{票号}.pdf` → K 列发票号 |
| 住宿 | `dzfp_{发票号}_*.pdf` → K 列发票号 |
| 滴滴 | `滴滴电子发票 (N).pdf` + `滴滴出行行程报销单 (N).pdf`，按日期顺序 |
| 高德聚合 | `【{平台}-{金额}】高德打车电子发票.pdf` + 行程单，平台名匹配 |

### 阶段 4：付款截图

**方案 A — Read 工具（推荐，准确率最高）：**
Claude 用 `Read` 逐张看图 → 识别金额 → 按行匹配（±2 元容差，支持两两拼单）→ 写入 `payment_mapping.json`。

**方案 B — OCR 兜底（弱，检出率约 10%）：**
```powershell
& python "{skill目录}\scripts\ocr_payments.py" "{工作目录}"
```
付款截图压缩度高 + 竖版排版复杂，Tesseract 对其效果很差。**实际场景下月度复用上期映射是更可靠的起点**（见阶段 1）。

**增量处理：** 默认合并已有 `payment_mapping.json`，跳过已映射文件和已覆盖行。如需重新识别所有截图，加 `--overwrite`：
```powershell
& python "{skill目录}\scripts\ocr_payments.py" "{工作目录}" --overwrite
```

### 阶段 5：扫描件

```powershell
& python "{skill目录}\scripts\ocr_scans.py" "{工作目录}"
```

**增量处理：** 默认合并已有 `scan_mapping.json`，跳过已映射文件和已覆盖行。如需重新识别所有扫描件，加 `--overwrite`：
```powershell
& python "{skill目录}\scripts\ocr_scans.py" "{工作目录}" --overwrite
```

脚本内置 **3 阶段 Pipeline**（详见 [扫描件匹配策略](references/scan-matching-strategy.md)）：

| 阶段 | 策略 | 覆盖 |
|------|------|------|
| Phase 1 快速 | psm=6, 1x, 不反色 → 发票号精确/尾 15 位匹配 + 金额兜底 | ~60% |
| Phase 2 深搜 | 仅未命中：2x + psm=[3,6,7] + 反色 → 同匹配策略 | ~80% |
| Phase 3 消元 | 剩余按行需求分配：先补发票位，再补行程单位 | ~100% |

消元法不依赖 OCR 关键词分类——任何未分配扫描件均可补位（避免 `DIDI TRAVEL` 被 `'D'` 单字符误判为高铁票）。

人工看图校验：
```powershell
& python "{skill目录}\scripts\convert_scans.py" "{工作目录}"
```

### 阶段 6：文档完整性检查

| 类型 | 发票 | 行程单 | 付款截图 | 扫描件 | 判定 |
|------|------|--------|---------|--------|------|
| 高铁 | 1 | — | 1 | 1 | 描述含"高铁/二等座" |
| 住宿 | 1 | — | 1 | 1 | 二级分类="住宿费" |
| 打车 | 1 | 1 | 1 | 1 | 描述含打车/滴滴/T3/如祺/风韵等 |
| 其他 | 1 | — | 1 | 1 | 兜底 |

缺项警告，不阻止生成。

### 阶段 7：生成 Excel

```powershell
& python "{skill目录}\scripts\build.py" "{工作目录}"
```

### 阶段 8：自检

**致命（exit 1）：** 数据行无图、孤儿 PDF、映射 PDF 不存在。
**警告：** 文档缺项、孤儿付款截图/扫描件。

## 约束

- 图片不压缩像素，仅 `xl.width/xl.height` 控制显示
- 不改源模板；`build.py` 固定不变，Claude 只更新 JSON 映射
- 不自动删文件。扫描件缺/多不阻止生成
- Tesseract 发票号匹配是扫描件主路径；Read 工具辅助
- PDF 用 `rglob('**/*.pdf')`，截图/扫描件目录用 `glob('*')`
- 付款截图 OCR 实用性低（~10%），月度复用优先复制上期映射

### 安全边界（output contract）

- **文件写入范围**：脚本只在 `<工作目录>` 内读写（`数据.csv`、`*.json`、`报销-{YY}-{MM}.xlsx`），不触及 skill 自身目录或系统目录
- **模板保护**：`scripts/模板.xlsx` 只读，写入始终生成新文件
- **文件覆盖**：`PermissionError`（Excel 正在被查看）时自动 fallback 到 `报销-{YY}-{MM}-NEW.xlsx`，不覆盖已打开的文件
- **半成品回滚**：生成失败时保留中途写入的 `.xlsx`（用户可检查调试），不自动删除
- **映射文件**：`payment_mapping.json` 和 `scan_mapping.json` 每次运行时合并（不覆盖已有映射），支持增量处理

### 输出风险画像

生成的是**财务报销提交文件**，以下场景输出可能"看起来对但实际不对"：

| 风险 | 信号 | 用户应检查 |
|------|------|-----------|
| 金额恰好匹配到错误行 | 两笔不同费用金额相同（如 ¥96.00 高铁和 ¥96.00 住宿） | 核对发票嵌入行是否与费用说明一致 |
| 发票号尾 15 位撞库 | OCR 提取的 15 位与其他发票号的尾 15 位相同（极低概率） | 核对发票号全量匹配 |
| 消元法配错 | 打车行的扫描件被分配给非打车行 | 检查打车行是否每行都有行程单截图 |
| 拼单金额≠报销金额 | 两笔付款加起来碰巧等于另一笔报销金额 | 核对付款截图商户名是否与费用行一致 |
| Tesseract 读空 | 模糊扫描件完全无法 OCR → 消元法随机分配 | 扫描件数量 > 行需求数时检查孤儿列表 |

以上风险**不阻止生成**——自检 fatal 只拦截"行无图"和"PDF 孤儿"两类确定性错误。其余交给用户肉眼确认。

## 参考

- [build.py](scripts/build.py) | [ocr_scans.py](scripts/ocr_scans.py) | [ocr_payments.py](scripts/ocr_payments.py) | [convert_scans.py](scripts/convert_scans.py)
- [扫描件匹配策略](references/scan-matching-strategy.md) | [踩坑记录](references/lessons-learned.md) | [常见问题](references/faq.md)
- [真实案例](references/real-case.md) | [文档要求清单](references/document-checklist.md)
