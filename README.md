# 📊 Excel VBA Expenses Dashboard v2 (FY 2026)

A fully automated, one-click **Expenses Dashboard** built entirely with Excel VBA. Run a single macro and watch it generate PivotTables, charts, KPI cards, slicers, and a professional dashboard — live, with status bar animation.

---

## ✨ Features

- 🟢 **7 PivotTables** — Department, Category, Monthly Trend, Approved Status, KPI Totals, Payment Method, Budget vs Actual
- 📈 **5 Dashboard Charts** — Clustered Column, Doughnut, Line with Markers, Budget vs Actual, Horizontal Bar
- 💳 **6 Live KPI Cards** — Total Expenses, Budget Allocated, Budget Remaining, Transactions, Approval Rate, Avg Expense
- 🔘 **2 Interactive Slicers** — Month (4 columns) & Department (6 columns), connected to all 7 PivotTables
- ⚙️ **Live Progress Animation** — Status bar updates at every build step with DoEvents
- 🎨 **Professional Green Theme** — Gradient banner, styled KPI cards, formatted charts, Aptos font throughout
- 🔗 **Formula-linked KPI values** — Cards update dynamically when slicers are used

---

## 🗂️ Required Data Sheet Structure

Your workbook must have a sheet named **`Data`** with these column headers in **Row 1**:

| Column | Description |
|---|---|
| Expense ID | Unique transaction ID |
| Date | Transaction date |
| Department | Spending department |
| Category | Expense category |
| Amount (INR) | Expense amount |
| Payment Method | Cash / Card / UPI etc. |
| Approved | Yes / No |
| Budget Allocated | Total budget for department |
| Budget Remaining | Remaining budget |

> A `Month` column is auto-added by the macro if missing.

---

## 🚀 How to Use

1. Open your Excel workbook with the `Data` sheet ready
2. Press `Alt + F11` to open the VBA Editor
3. Go to **Insert → Module**
4. Paste the full VBA code from `ExpensesDashboard.bas`
5. Close the VBA Editor
6. Press `Alt + F8` → select `BuildExpensesDashboard` → click **Run**
7. Watch the status bar for live progress updates
8. ✅ A success message confirms the dashboard is ready!

---

## 🛠️ Tech Stack

- Microsoft Excel (Office 2016 or later)
- VBA (Visual Basic for Applications)
- PivotTables & PivotCache
- Excel Slicers
- Chart Objects & Series Formatting
- Shape & TextBox Automation

---

## 📌 Notes

- `ScreenUpdating = True` is kept **ON** intentionally for live animation
- Row/column headings remain **visible** (DisplayHeadings = True)
- Dashboard zoom is set to **70%**
- Helper columns `AF:AG` are auto-hidden after build
- Page setup is configured for **1-page print fit**

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

⭐ If you found this useful, give the repo a star!
