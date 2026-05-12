# HKU_reimbursement_generator

A Windows desktop app for generating HKU reimbursement materials.

The app interface is Chinese + English. The generated Excel, Word, filenames, and receipt descriptions are English.

## Install

```powershell
git clone https://github.com/skywalker0525/HKU_reimbursement_generator.git
cd HKU_reimbursement_generator
npm install
```

## Run

```powershell
npm start
```

Or double-click:

```text
Start_Reimbursement_Generator.bat
```

## Use

1. Fill in claimant information:
   - Claimant / Payee Name
   - Department
   - Staff/Student No.
   - Telephone No.
2. Click `选择票据 / Select Receipts`.
3. Select receipt PDFs or images.
4. Click `选择输出位置 / Select Output Folder`.
5. Click `识别票据 / Analyze Receipts`.
6. Review and edit the preview table.
7. Click `生成材料 / Generate Package`.

## Output

The app creates:

```text
Reimbursement_DATE_CLAIMANT_NAME/
```

Main output files:

- `receipts/`: renamed receipt copies.
- `receipt_screenshots/`: receipt images inserted into Word.
- `exchange_rates/`: HKAB exchange-rate screenshots.
- `615_404 Reimbursement form_CLAIMANT_NAME.xlsx`
- `Reimbursement_CLAIMANT_NAME.docx`

## Notes

- The HKU Excel template is included in `templates/`.
- RMB receipts use HKAB CNY Selling by invoice date.
- Excel total is formula-based and moves down automatically if there are many reimbursement items.
- Original receipt files are copied, not overwritten.
