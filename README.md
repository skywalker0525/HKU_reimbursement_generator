# HKU Reimbursement Generator

Windows desktop app for preparing HKU reimbursement packages.

用于生成香港大学报销材料的 Windows 桌面程序。

The app UI is Chinese and English. Generated Excel, Word documents, filenames, and receipt descriptions are prepared in English for reimbursement submission.

软件界面为中英双语。生成的 Excel、Word、文件名和票据英文描述会按英文报销材料准备。

## Features / 功能

- Analyze PDF, JPG, JPEG, and PNG receipt files.
- 识别 PDF、JPG、JPEG、PNG 票据文件。
- Convert PDF receipts directly into full PDF page images, without capturing the PDF viewer toolbar, browser interface, or only part of the page.
- PDF 票据会直接渲染成完整页面图片，不截取 PDF 查看器工具栏、浏览器界面，也不会只截取页面的一部分。
- Generate the HKU reimbursement Excel form from the included template.
- 使用内置模板生成港大报销 Excel 表格。
- Generate a Word document with receipt images and exchange-rate evidence.
- 生成包含票据图片和汇率证明的 Word 文档。
- Fetch HKAB exchange rates automatically for supported currencies.
- 自动查询 HKAB 支持币种的汇率。
- Stop generation when required exchange rates are missing, so the user can fill the rate and generate again.
- 如果缺少必要汇率，程序会中断生成，用户填好汇率后再重新生成。
- Save and reuse claimant profiles.
- 保存并复用常用报销人信息。
- Configure Word layout options, including page breaks and exchange-rate evidence position.
- 可设置 Word 排版选项，包括是否每张票据单独分页，以及汇率证明放在票据图片之前或之后。

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

Or double-click:

也可以双击：

```text
Start_Reimbursement_Generator.bat
```

## Use / 使用

1. Fill in claimant information, or load a saved claimant profile.
   填写报销人信息，或选择已保存的常用报销人。
2. Select the Word layout options.
   选择 Word 排版选项。
3. Click `选择票据 / Select Receipts`.
   点击 `选择票据 / Select Receipts`。
4. Select receipt PDFs or images.
   选择 PDF 或图片票据。
5. Click `选择输出位置 / Select Output Folder`.
   点击 `选择输出位置 / Select Output Folder`。
6. Click `识别票据 / Analyze Receipts`.
   点击 `识别票据 / Analyze Receipts`。
7. Review and edit the preview table.
   检查并修改预览表格。
8. Click `生成材料 / Generate Package`.
   点击 `生成材料 / Generate Package`。

If the app reports missing exchange rates, fill `HKAB Rate` or `HKD Amount` in the highlighted rows, then click `Generate Package` again.

如果程序提示缺少汇率，请在高亮行填写 `HKAB Rate` 或 `HKD Amount`，然后再次点击生成。

## Exchange Rates / 汇率

- `HKD` receipts do not need exchange rates.
- `HKD` 票据不需要汇率。
- `RMB/CNY`, `USD`, `EUR`, `GBP`, and `JPY` can use automatic HKAB lookup when available.
- `RMB/CNY`、`USD`、`EUR`、`GBP`、`JPY` 可在 HKAB 可查询时自动获取汇率。
- HKAB rates are Hong Kong dollars per 100 units of foreign currency.
- HKAB 汇率为每 100 单位外币对应的港币金额。
- `OTHER` requires a manual `HKAB Rate` or `HKD Amount`.
- `OTHER` 币种需要手动填写 `HKAB Rate` 或 `HKD Amount`。
- Generation stops before writing the output package if required rates are missing.
- 如果必要汇率缺失，程序会在写入输出包之前停止。

## Output / 输出

The app creates an output folder like:

程序会创建类似下面的输出文件夹：

```text
Reimbursement_YYYYMMDD_CLAIMANT_NAME/
```

Main output files:

主要输出文件：

- `receipts/`: renamed receipt copies.
- `receipts/`：重命名后的票据副本。
- `receipt_screenshots/`: receipt images inserted into Word.
- `receipt_screenshots/`：插入 Word 的票据图片。
- `exchange_rates/`: HKAB exchange-rate evidence images and HTML evidence files.
- `exchange_rates/`：HKAB 汇率证明图片和 HTML 证明文件。
- `615_404 Reimbursement form_CLAIMANT_NAME.xlsx`: completed HKU Excel form.
- `615_404 Reimbursement form_CLAIMANT_NAME.xlsx`：生成后的港大报销 Excel 表格。
- `Reimbursement_CLAIMANT_NAME.docx`: supporting Word document.
- `Reimbursement_CLAIMANT_NAME.docx`：报销支持材料 Word 文档。
- `receipt_preview.json`: generated receipt metadata for review.
- `receipt_preview.json`：生成时使用的票据信息记录。

Original receipt files are copied, not overwritten.

原始票据只会被复制，不会被覆盖。

## Test / 测试

```powershell
npm test
```

On Windows PowerShell, if script execution policy blocks `npm`, use:

如果 Windows PowerShell 的脚本执行策略阻止 `npm`，请使用：

```powershell
npm.cmd test
```

## Package For Windows / 打包 Windows 应用

Create a portable Windows build:

生成 Windows 便携版：

```powershell
npm run package:win
```

Create Windows distribution targets:

生成 Windows 发布包：

```powershell
npm run dist:win
```

The packaged files are written to `dist/`.

打包后的文件会输出到 `dist/`。

## Project Notes / 项目说明

- The HKU Excel template is included in `templates/`.
- 港大 Excel 模板位于 `templates/`。
- Claimant profiles are stored in Electron user data, not in the project folder.
- 常用报销人信息保存在 Electron 用户数据目录，不保存在项目目录。
- Regression tests are in `test/`.
- 回归测试位于 `test/`。
