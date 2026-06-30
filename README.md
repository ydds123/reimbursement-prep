# reimbursement-prep

> 报销发票材料准备 —— 从模板 Excel、费用数据、发票 PDF、付款截图、扫描件，自动生成带嵌入图片的完整报销表。

## 一句话

把 200 多张 PDF 发票、付款截图、手机扫描件和 CSV 费用数据整合成一份可直接交财务的 Excel，5 分钟内完成。

## 适用场景

- 月度/季度项目差旅费报销
- 批量电子发票整理
- 发票号与费用行的自动匹配
- 扫描件 OCR 识别与归档

## 快速开始

### 环境

```bash
pip install pypdfium2 Pillow openpyxl pytesseract
```

Windows 需额外安装 [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)。

### 工作目录结构

```
{工作目录}/
├── 数据.csv                  # 12 列无表头费用数据
├── 内容/  或  *.pdf           # 电子发票/行程单 PDF
├── 付款截图/                  # 微信/支付宝 .jpg
├── 扫描件/                    # 手机相机直出 .jpg
├── payment_mapping.json       # 付款截图→行号映射
└── scan_mapping.json          # 扫描件→行号映射
```

### 运行

通过 Claude Code 加载本 skill：

```
/reimbursement-prep
```

然后按提示提供工作目录路径。Skill 会自动：

1. 检查依赖环境
2. PDF 发票按文件名规则匹配到费用行
3. 付款截图通过 AI 视觉识别金额匹配
4. 扫描件通过 Tesseract OCR 提取发票号匹配（三阶段 Pipeline：快速→深搜→消元）
5. 文档完整性检查（缺项警告不阻止）
6. 生成带嵌入式图片的 Excel

## 技术点

| 环节 | 策略 | 准确率 |
|------|------|--------|
| PDF 匹配 | 文件名规则直匹（高铁票号、dzfp 发票号、滴滴日期） | ~100% |
| 付款截图 | AI 视觉识别金额 + ±2 元容差 | ~95% |
| 扫描件 | Tesseract OCR + 三阶段 Pipeline（psm 3-7, 反色, 2x） | ~80% |
| 消元补位 | 行需求驱动分配，不依赖 OCR 关键词分类 | ~100% |

## 约束

- 图片不压缩像素，仅控制 Excel 显示尺寸
- 不改源模板；`build.py` 固定不变，只更新 JSON 映射
- 不自动删文件；缺项/多余不阻止生成
- 月度复用优先从上期复制映射文件

## 许可

MIT
