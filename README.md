# Bilingual Reimbursement Generator

This is a local Chinese/English desktop app for generating English HKU reimbursement materials.

## How to Run

Double-click `Start_Reimbursement_Generator.bat`, or run:

```powershell
npm.cmd start
```

## Workflow

1. Fill `Claimant / Payee Name`, `Department`, `Staff/Student No.`, and `Telephone No.`.
2. Click `选择票据 / Select Receipts` and select PDF/JPG/PNG receipt files.
3. Click `选择输出位置 / Select Output Folder`.
4. Click `识别票据 / Analyze Receipts`.
5. Review and edit the bilingual preview table. Final generated text should be English.
6. Click `生成材料 / Generate Package` and watch the progress bar.

## Output

The app creates an English-named folder based on the claimant name, such as `Reimbursement_20260512_CLAIMANT_NAME`. If the claimant is blank, it uses `CLAIMANT`.

- `piao/`: copied source receipts renamed in English.
- `exchange_rates/`: HKAB exchange-rate screenshots.
- `receipt_screenshots/`: screenshots used inside the Word document.
- `615_404 Reimbursement form_CLAIMANT_NAME.xlsx`: completed Excel form.
- `Reimbursement_CLAIMANT_NAME.docx`: English supporting document.
- `receipt_preview.json`: audit copy of the reviewed data.

## Notes

- The Excel template is copied from the blank form `..\报销\615_404 Reimbursement form.xlsx`.
- The app only fills: Claimant / Payee Name, Department, Staff/Student No., Telephone No., Description, Voucher, and Amount.
- Claimant fields are blank by default; enter the correct person before generating.
- RMB conversion uses `HKD = RMB amount × HKAB CNY Selling rate / 100`.
- The app first uses HKAB's official source/API endpoint `https://www.hkab.org.hk/api/member/public/getExrate/YYYY-MM-DD` for invoice-date CNY Selling rates, then saves an evidence screenshot of the returned CNY row.
- PDF receipts are converted from the rendered PDF page canvas into full-page PNG images without cropping.
- The Excel total cell after `Total :` uses the formula `SUM(K29:K54)`.
- Original receipts are never overwritten.
