# HKU_reimbursement_generator

Windows desktop app for generating HKU reimbursement materials.

用于生成港大报销材料的 Windows 桌面程序。

The app UI is Chinese + English. Generated Excel, Word, filenames, and receipt descriptions are English.

软件界面为中文 + 英文；生成的 Excel、Word、文件名和票据描述均为英文。

## Install / 安装

```powershell
git clone https://github.com/skywalker0525/HKU_reimbursement_generator.git
cd HKU_reimbursement_generator
npm install
```

## Run / 运行

```powershell
npm start
```

Or double-click / 或双击：

```text
Start_Reimbursement_Generator.bat
```

## Use / 使用

1. Fill in claimant information / 填写报销人信息：
   - Claimant / Payee Name
   - Department
   - Staff/Student No.
   - Telephone No.
2. Click `选择票据 / Select Receipts`.
3. Select receipt PDFs or images / 选择 PDF 或图片票据。
4. Click `选择输出位置 / Select Output Folder`.
5. Click `识别票据 / Analyze Receipts`.
6. Review and edit the preview table / 检查并修改预览表。
7. Click `生成材料 / Generate Package`.

## Output / 输出

The app creates / 程序会生成：

```text
Reimbursement_DATE_CLAIMANT_NAME/
```

Main output files / 主要输出文件：

- `receipts/`: renamed receipt copies / 重命名后的票据副本。
- `receipt_screenshots/`: receipt images inserted into Word / 插入 Word 的票据图片。
- `exchange_rates/`: HKAB exchange-rate screenshots / HKAB 汇率截图。
- `615_404 Reimbursement form_CLAIMANT_NAME.xlsx`
- `Reimbursement_CLAIMANT_NAME.docx`

## Notes / 说明

- The HKU Excel template is included in `templates/`.
- HKU Excel 模板已包含在 `templates/` 文件夹中。
- RMB receipts use HKAB CNY Selling by invoice date.
- 人民币票据按发票日期使用 HKAB CNY Selling 汇率。
- Excel total is formula-based and moves down automatically if there are many items.
- Excel 总额使用公式；如果报销事项较多，Total 行会自动下移。
- Original receipt files are copied, not overwritten.
- 原始票据只会复制，不会被覆盖。
