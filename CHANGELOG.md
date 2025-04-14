# Changelog

### App Updates

---

**_Updates to V1.1:_**  
- Removed Account Analyses  
- Added functionality to `extract_data_from_excel` function  
- Dropped NaN columns  
- Ordered the columns differently with formatting  
- Checking if assets and PO are None  
- Fixed the PO Number heading requirements  
- Renamed "Transaction Description" to "Cluster"  
- Added a "balances" sheet

---

**_Updates to V1.2:_**  
- Bug fixes regarding BCA files with no positive commitments  
- No "balances" sheet if there is no assets file  
- Added Obligation grouping round

---

**_Updates to V1.3:_**  
- Format changes  
- Added `Account_number_changes.xlsx`

---

**_Updates to V1.6:_**  
- Added multi-file processing  
- Error catching for files during automated process

---

**_Updates to V1.7:_**  
- Changed all **Balance Type** entries for incomes from `"Expenditure Income"` to `"Income"`  
- Reversed previous logic that treated negative expenditures as income — only `"adjustments receipts"` and `"adjustment: budget increase"` are treated as income  
- Updated descriptions in the “Balances” tab (Column A) to:  
  - Period opening balances (from reports):  
  - Period closing balances (available budget):  
  - Commitments during period (from reports):  
  - Obligations during period (from reports):  
  - Expenses during period (calculated):  
  - Total consumption during period (from reports - calculated total income):  
  - Total income during period (calculated):  
- Set column widths in “Balances” sheet:  
  - Column A: `1150`  
  - Columns B, C, and D: `330`  
- Corrected formulas in “Balances” tab:  
  - **C5**: `=IF(ISNUMBER('BCA Assets'!AC23),'BCA Assets'!AC23,0)`  
  - **C6**: `=-(C3-C2-C8-C4-C5)`  
- Fixed **B8** and **C8** to correctly calculate the sum of all income entries

