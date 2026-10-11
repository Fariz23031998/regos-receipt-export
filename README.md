# REGOS receipt export

Two scripts for a REGOS Firebird database:

1. `export_receipt.py` reads one receipt and saves its lines to Excel.
2. `build_qr_pdf.py` reads that Excel file and builds a PDF. Each page has the item name and a QR code of its marking.

## What you need

- Python 3.11 or newer
- [Firebird 3](https://firebirdsql.org/) installed, with `isql` available and the server running
- The shop database file `Regos.FDB` in this folder (the same folder as the scripts)

The export script connects as `SYSDBA` / `masterkey` to `localhost` and the `Regos.FDB` file next to the script.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Export a receipt to Excel

Use the receipt code shown on the register, including the prefix:

```powershell
python export_receipt.py 1-0751361
```

This writes `check-1-0751361.xlsx` in the current folder. Columns:

| Column | Meaning |
| --- | --- |
| Code | Item code |
| Name | Item name |
| Quantity | Quantity on that line |
| Marking | Marking code scanned for that line |

To choose the file path:

```powershell
python export_receipt.py 1-0751361 -o D:\receipt.xlsx
```

An open receipt is read from the current sale lines. A posted receipt is read from the saved operations. If the code does not exist, or the receipt has no lines, the script stops with an error.

## Build the QR PDF

```powershell
python build_qr_pdf.py check-1-0751361.xlsx
```

This writes `check-1-0751361.pdf` next to the Excel file. One page per row that has a marking: the product name at the top, a QR code of the marking, and the marking text under the code.

To choose the PDF path:

```powershell
python build_qr_pdf.py check-1-0751361.xlsx -o D:\labels.pdf
```

The workbook must have `Name` and `Marking` columns. Rows with an empty marking are skipped. If every marking is empty, the script stops and does not create a PDF.

## Typical run

```powershell
python export_receipt.py 1-0751361
python build_qr_pdf.py check-1-0751361.xlsx
```
