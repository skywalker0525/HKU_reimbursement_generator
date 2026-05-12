# Improvement Issues / 改进 Issue 清单

These are recommended follow-up issues for the project.

以下是建议后续拆分跟进的改进事项。

## 1. Improve Receipt Field Extraction

Current extraction is rule-based. Add a review-friendly confidence score for invoice number, date, description, currency, and amount.

目前票据信息识别主要基于规则。建议为发票号、日期、描述、币种和金额增加置信度，方便用户优先检查不确定项。

## 2. Add App Packaging

Package the Electron app into a Windows installer or portable executable so users do not need to run `npm install`.

将 Electron 应用打包成 Windows 安装包或绿色版可执行文件，避免用户手动执行 `npm install`。

## 3. Add Configurable Claimant Profiles

Allow users to save common claimant profiles, including name, department, student/staff number, and telephone.

支持保存常用报销人信息，包括姓名、部门、学号/职员号和电话。

## 4. Improve Word Layout Controls

Let users choose whether each receipt item starts on a new page and whether exchange-rate evidence appears before or after the receipt image.

增加 Word 排版选项，例如每个事项是否单独分页、汇率截图放在票据图片前还是后。

## 5. Add Automated Regression Tests

Add test fixtures for sample PDFs and validate extracted amount, invoice date, exchange rate, Excel total formula, and Word page breaks.

增加自动化回归测试，用样例 PDF 校验金额、开票日期、汇率、Excel 总额公式和 Word 分页。

## 6. Support More Currency Types

Currently RMB and HKD are the main supported currencies. Add USD and other common currencies with clear exchange-rate sources.

目前主要支持人民币和港币。后续可增加美元等常见币种，并明确对应汇率来源。
