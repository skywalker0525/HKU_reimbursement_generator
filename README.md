# HKU_reimbursement_generator

A local desktop app for preparing HKU reimbursement materials.

The app interface is Chinese + English. The generated reimbursement package is English only.

## What It Generates

For selected PDF/JPG/PNG receipts, the app creates a folder such as:

```text
Reimbursement_20260512_CLAIMANT_NAME/
```

Inside the folder:

- `receipts/`: copied source receipts renamed in English.
- `receipt_screenshots/`: full-page receipt images used in the Word document.
- `exchange_rates/`: HKAB exchange-rate evidence screenshots.
- `615_404 Reimbursement form_CLAIMANT_NAME.xlsx`: completed HKU Excel form.
- `Reimbursement_CLAIMANT_NAME.docx`: supporting Word document.
- `receipt_preview.json`: reviewed data used for generation.

## Features

- Bilingual desktop UI.
- Multi-select PDF/JPG/PNG receipts.
- Editable preview table before final generation.
- PDF receipts are converted from the rendered PDF page canvas into full-page PNG images without cropping.
- RMB receipts use HKAB CNY Selling rate by invoice date.
- HKAB rate source/API:

```text
https://www.hkab.org.hk/api/member/public/getExrate/YYYY-MM-DD
```

- RMB conversion formula:

```text
HKD = RMB amount * HKAB CNY Selling / 100
```

- Excel only fills:
  - Claimant / Payee Name
  - Department
  - Staff/Student No.
  - Telephone No.
  - Description
  - Voucher
  - Amount
  - Total

- Excel total cell after `Total :` uses:

```text
=SUM(K29:K54)
```

## Requirements

- Windows
- Node.js 20 or newer
- Internet connection for HKAB exchange-rate lookup

## First-Time Setup

Clone the repository:

```powershell
git clone https://github.com/skywalker0525/HKU_reimbursement_generator.git
cd HKU_reimbursement_generator
```

Install dependencies:

```powershell
npm install
```

## Run the App

Start from PowerShell:

```powershell
npm start
```

Or double-click:

```text
Start_Reimbursement_Generator.bat
```

## How To Use

1. Fill in:
   - `Claimant / Payee Name`
   - `Department`
   - `Staff/Student No.`
   - `Telephone No.`
2. Click `选择票据 / Select Receipts`.
3. Select PDF/JPG/PNG receipts.
4. Click `选择输出位置 / Select Output Folder`.
5. Click `识别票据 / Analyze Receipts`.
6. Review and edit the preview table.
7. Click `生成材料 / Generate Package`.

## Required Local Template

The app expects this blank HKU reimbursement Excel form to exist relative to the app folder:

```text
../报销/615_404 Reimbursement form.xlsx
```

This matches the original local workspace layout:

```text
新建文件夹/
  BilingualReimbursementGenerator/
  报销/
    615_404 Reimbursement form.xlsx
```

If you move the app, keep the `报销/615_404 Reimbursement form.xlsx` template one folder above the app directory.

## Notes

- Original receipt files are never overwritten.
- If a receipt field is uncertain, edit it in the preview table before generation.
- If HKAB cannot return a rate for a date, the preview marks the rate for manual review instead of silently guessing.
- Final generated filenames and reimbursement document content are English only.
