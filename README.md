# Statement to Excel

A small Windows desktop tool that turns HABIBMETRO bank SMS alerts into a formatted Excel statement.

Copy the SMS alerts into a Notepad `.txt` file, pick the file in the app, and you get an `.xlsx` with:

- **Transactions** sheet: date, time, account, description, category, debit and credit, plus a running balance (Excel formulas)
- **Summary** sheet: number of transactions and totals for each category, and the net amount

## Usage

**Option 1: the .exe (no Python needed)**
Download `Statement-to-Excel.exe` from the [Releases](../../releases) page and double-click it.

**Option 2: run from source**
```
pip install -r requirements.txt
python statement_to_excel.pyw
```

Then click **Select TXT File** and choose your file. The Excel file is saved next to it with the same name.

Try it with [`sample.txt`](sample.txt), which contains made-up data.

## Input format

One alert per line, as the bank sends it:
```
HABIBMETRO - Alert: 01-01-26 10:20 00000-**00** Amt 75.00 Debited /SOC - GSM Subscription Charges [Monthly]...
```
If a line can't be read, the app skips it and tells you how many lines were skipped.

## Build the .exe yourself
```
pip install pyinstaller
pyinstaller --onefile --windowed --name "Statement to Excel" statement_to_excel.pyw
```

## Privacy
The app runs completely offline and sends no data anywhere. Don't commit real statements to this repo. The `.gitignore` blocks `.txt`, `.xlsx` and image files, except `sample.txt`.
