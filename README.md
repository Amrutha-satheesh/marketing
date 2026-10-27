# Marketing Agency - Multi-Currency Financial Tracker (Base: INR ₹)

A dynamic financial tracker built for a marketing company with **4 initial employees** and built-in scalability to add and analyze future candidates/hires.

* **Primary Base Accounting Currency**: **INR (₹)**
* **Supported Currency Conversions**: **USD ($)** and **AED (UAE Dirham)**

---

## 💱 Currency Conversion Model (Base: INR ₹)

The system treats **Indian Rupee (INR ₹)** as the company's accounting base. All figures convert dynamically:

| Target Currency | Conversion Rule | Example Calculation (Salary: ₹45,000) |
|---|---|---|
| **INR (₹)** | Base (1.0) | `₹45,000.00` |
| **USD ($)** | $\text{INR} \div 94.94$ | `₹45,000 ÷ 94.94 = $473.98` |
| **AED (Dirham)** | $\text{INR} \div 25.85$ | `₹45,000 ÷ 25.85 = AED 1,740.81` |

*Exchange rates can be edited directly in the Web App via the **"Edit Rates"** button.*

---

## 📁 Included Files

1. **`index.html`**: Interactive Web Application dashboard:
   - **Defaults immediately to INR (₹)** as base currency.
   - Real-time conversion of table inputs, calculated cells, overhead inputs, and summary KPI cards.
   - Seamless two-way editing (edit values in INR, USD, or AED).
   - Changeable Overhead Manager (Office Rent, Software Subscriptions, Utilities):
     - **➕ Add to Current Cost**: Adds the distributed overhead share directly on top of each employee's existing cost (e.g. Current ₹6,000 + ₹9,000 = ₹15,000).
     - **🔄 Set Equal Share**: Sets each employee's cost to the exact equal overhead share (e.g. ₹9,000).
   - Candidate hiring modal and one-click Excel (.xlsx) export.

2. **`Marketing_Company_Financial_Tracker.xlsx`**: Excel spreadsheet featuring 4 organized tabs:
   - **Tab 1: `INR (₹) - Base Tracker`** *(Primary base sheet with native formulas)*
   - **Tab 2: `USD ($) - Converted`** *(Live converted sheet)*
   - **Tab 3: `AED - Converted`** *(Live converted sheet)*
   - **Tab 4: `Rates & Overheads`** *(Exchange rate details and expense breakdown)*

3. **`generate_tracker_excel.py`**: Python script to regenerate the spreadsheet anytime.

---

## 📊 Default Base Team & Projects (INR ₹)

| Employee Name | Monthly Salary | Changeable Cost | Working Hours | Per Hour Value | Project Name | Project Time | Project Value | Profit Margin |
|---|:---:|:---:|:---:|:---:|:---|:---:|:---:|:---:|
| **Alex Morgan** *(SEO Lead)* | ₹45,000 | ₹6,000 | 160 hrs | **₹281.25 / hr** | Tech Corp SEO & Organic Growth | 80 hrs | ₹55,000 | **+₹32,500.00 (59.1%)** |
| **Sarah Jenkins** *(Paid Ads)* | ₹40,000 | ₹6,000 | 160 hrs | **₹250.00 / hr** | Q4 E-Commerce Meta & Google Ads | 90 hrs | ₹52,000 | **+₹29,500.00 (56.7%)** |
| **David Kim** *(Content Lead)* | ₹35,000 | ₹5,000 | 160 hrs | **₹218.75 / hr** | SaaS Lead Gen Content Hub | 70 hrs | ₹40,000 | **+₹24,687.50 (61.7%)** |
| **Elena Rostova** *(Creative Lead)* | ₹38,000 | ₹5,000 | 160 hrs | **₹237.50 / hr** | Fintech Brand Refresh & Social | 75 hrs | ₹45,000 | **+₹27,187.50 (60.4%)** |
| **TOTAL / SUMMARY** | **₹1,58,000** | **₹22,000** | **640 hrs** | **₹246.88 / hr avg** | **4 Projects** | **315 hrs** | **₹1,92,000** | **+₹1,14,312.50 (59.5%)** |
