# AP Workload & Performance Dashboard Generator

## Run locally
1. Install Python 3.10 or newer.
2. Extract this ZIP.
3. In the extracted folder run `python -m pip install -r requirements.txt`.
4. Run `streamlit run app.py`.
5. Upload an invoice export with the 19 columns required by the state AP source format.
6. Select the reporting date (the newest source *Date Received* is suggested, but you should choose your official report as-of date).
7. Click **Generate Excel dashboard**, then download the result.

## Streamlit Community Cloud
Put all files including `assets/reference_dashboard.xlsx` in a private GitHub repository, connect the repository to Streamlit Community Cloud, and select `app.py` as the entry point. The Excel template contains operational invoice information from your reference workbook; **do not publish this repo publicly** without first removing sensitive/vendor information and replacing it with an approved sanitized template.

## What is preserved
All ten Excel worksheets, the original drawing/chart objects, cell formatting, frozen panes and chart positions are preserved from the reference template. The Master Data tab is replaced with uploaded data (up to 20,000 records), while prepared row calculation formulas are retained. Master Data AE2 holds the as-of date. The app writes current monthly caches and updates the series references for the relevant Dashboard / Monthly Summary charts using internal workbook references. All original charts remain packaged.

## Payment cycle logic
Include a cycle only when the invoice status is exactly `Paid`, both Date Received and PS Entry Date are valid, and `PS Entry Date >= Date Received`. A zero-day cycle is valid; a missing or negative cycle is excluded. Average and median use the same filtered list. The app reports exclusions and writes current verified cycle summaries into Excel caches; Master Data continues to contain linked formulas.

## Important limits and validation status
- Max invoices: 20,000 due to fixed reference formulas.
- Monthly Summary: existing reference layout supports December 2025 through November 2032 (84 monthly rows); only periods through the selected date are charted.
- Charts use correct live *internal* references for the reporting period at **generation time**. Uploading a new month's file regenerates and extends chart series automatically. Editing Master Data directly in Excel without regenerating the report will recalculate worksheet formulas but **will not extend the chart series ranges by itself**; that would require an Excel Table / named-range chart compatibility change.
- Some historical snapshot comparisons (Reports 1–3) remain fixed according to the original reference file; latest Report 4 values depend on Master Data. Historical baseline datasets cannot be recreated from a single current invoice export.
- The original file is not entirely worksheet-independent: certain original formulas still reference other summary sheets. Those historical dependencies are preserved so the reference look is unchanged; a fully independent rebuild needs a separate audit.
- This is an XML-preserving workbook generator: it validates ZIP/XML integrity but does not execute Microsoft Excel's calculation engine. Test opening the output in desktop Excel and check chart rendering before official circulation.
- Incoming Excel input date cells must be genuine Excel dates or consistent date strings. Duplicate invoice-number/vendor combinations are not automatically deleted.

## Data dictionary
A:S corresponds to ID, Status, Invoice Number, Vendor, Invoice Date, Date Received, Amount, Aging Received, Urgent, Category, Fiscal Year, PO Number, Department, Fund, Invoice Approver, Invoice Approver 2, AP Payer, PS Entry Date, Voucher Number. All additional Master Data columns are derived helper fields.

## Troubleshooting
If charts display empty in Excel, open desktop Excel, enable automatic recalculation and save the generated report. When possible, use Excel 365. If Excel presents a repair log, capture it; do not use that report for production until investigated.
