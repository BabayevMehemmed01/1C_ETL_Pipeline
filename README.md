# 1C to Power BI ETL Pipeline

Operational reports exported from 1C ERP are not tables. They are outlined worksheets: a department total sits above a client total, the client total sits above invoice lines, and several of those labels share merged cells. Loading that file into Power BI either double-counts the money or drops the group names that give each invoice its context.

This project is an ETL pipeline that turns one of those exports into a single flat fact table. Group totals stay out of the result. The names those groups carried are written onto every invoice row, so a later sum in Power BI equals the total printed on the original report.

## Project overview

The extractor reads the worksheet as 1C wrote it, including cells that sit inside a collapsed outline. The transform step classifies each row as a report header, a group total, an invoice line, or the final total. Group totals are not loaded. Their labels are carried down onto the invoice lines beneath them: division, then client or supplier, then the document itself. Merged ranges are read from the anchor cell, so a payer name that visually spans four columns becomes one field.

The loader writes a header row in the same typeface and color as the values under it, with bold as the only distinction, plus a frozen header and an autofilter. Dates and amounts are real Excel types, not text that happens to look like a date.

## Data cleaning and wrangling

A grouped receivables export typically arrives in three stacked bands:

| What 1C prints | What the pipeline keeps |
| --- | --- |
| Title, filter, and a two-row header with merged amount groups | One header row. `Задолженность` over `Сумма` becomes `Задолженность Сумма` |
| A division row and a client row that only hold subtotals | The names, repeated on each invoice. The subtotals are dropped |
| An invoice line with the document, currency, due date, and amounts | One fact row |
| A closing `Итого` row | Used as a check, then omitted from the output |

The same pass covers the other layouts this export family uses: a flat receivables list, a payables hierarchy (division, supplier, document), and a trial-balance sheet where the account code is a group row and each counterparty is a line. On the trial balance, the printed total is the net of debit and credit, so the reconciliation compares `sum(debit) - sum(credit)` with that net instead of summing a single column.

After the write, the pipeline adds the detail amounts and compares them with the report total. A mismatch exits with a non-zero status rather than silently shipping a table that will not tie out.

## Tech stack

Python, pandas, and openpyxl. pandas holds the cleaned fact table. openpyxl writes the workbook and the header formatting. `python-calamine` reads the source worksheet, including the outlined rows a normal preview hides.

## Layout

```
data/input/dummy_1c_raw.xlsx       grouped 1C-style export
data/output/dummy_cleaned_data.xlsx  flat table for Power BI
temizle_excel.py
```

`temizle_excel.py` reads only `data/input/dummy_1c_raw.xlsx` and writes `data/output/dummy_cleaned_data.xlsx`.

## Run

```bash
pip install -r requirements.txt
python temizle_excel.py
```

## Note

The workbook in this repository is synthetic. Client names, invoice numbers, and amounts were invented so the transformation can be reviewed without exposing operating data. The original receivables, payables, and trial-balance extracts are confidential and are not part of this repository.
