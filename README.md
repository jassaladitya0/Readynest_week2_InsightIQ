# InsightIQ — Customer & Sales Analytics
### Designed & Developed by Aditya Jassal

---

## Project Structure

```
InsightIQ/
├── backend/
│   └── app.py          ← Python Flask backend (all analytics here)
├── frontend/
│   └── index.html      ← Complete frontend (open in browser)
└── README.md
```

---

## How to Run

### Step 1 — Install Python dependencies
```bash
pip install flask flask-cors pandas numpy openpyxl xlrd reportlab matplotlib
```

### Step 2 — Start the Python backend
```bash
cd backend
python app.py
```
You should see: `InsightIQ Backend running on http://localhost:5050`

### Step 3 — Start the Frontend server
```bash
cd frontend
python -m http.server 3000
```

### Step 4 — Open the Frontend URL in your browser
Open [http://localhost:3000](http://localhost:3000) in your browser.

*(Alternatively, you can still open `frontend/index.html` directly in your browser).*

---

## Features

### Upload
- Supports **CSV, XLSX, XLS, JSON** files
- Auto-detects customer, sales, product, region, date, quantity columns
- Sample dataset included (click "Load sample dataset")

### Overview Dashboard
- 6 KPIs: Revenue, Customers, Avg Order, High Value Count, Products, Top Region
- Monthly Sales Trend (line chart)
- Region-wise Revenue (bar chart)
- Customer Segmentation pie (High/Medium/Low Value)
- New vs Returning customers
- Filters: Region, Segment, Customer Type

### Customer Analysis
- Segment cards (High / Medium / Low Value) with revenue stats
- Full customer table with search (rank, name, spend, orders, region, type, segment)

### Product Performance
- Top 10 Products by revenue (horizontal bar chart)
- Category Breakdown (bar chart)
- Full product table with search

### Business Recommendations
- Python-generated rule-based suggestions (10 strategies)
- Key insights derived from your data
- Data quality report (detected columns, null counts)
- Priority: High / Medium / Low

### Reports & Downloads
- **PDF Report** (Python ReportLab): Cover page, KPIs, all charts embedded, segmentation tables, top customers, top products, key insights, business recommendations, data quality section. Dark-themed, professional.
- **Cleaned Dataset (CSV)**: Whitespace trimmed, numeric coercion, date normalized, duplicates removed, nulls filled, `customer_segment` column added
- **Cleaned Dataset (XLSX)**: Same as above in Excel format
- Raw data preview (first 10 rows)

---

## API Endpoints

| Method | Endpoint   | Description                        |
|--------|------------|------------------------------------|
| GET    | /health    | Check if backend is running        |
| POST   | /analyze   | Full analysis — returns JSON       |
| POST   | /report    | Generate & download PDF report     |
| POST   | /clean     | Download cleaned CSV or XLSX       |

---

## Dataset Format
Your file can have any column names — the backend auto-detects:
- `customer_id` / `customer` / `cust_id`
- `sales` / `revenue` / `amount` / `total`
- `product` / `item` / `product_name`
- `category` / `cat`
- `region` / `location` / `area` / `city`
- `order_date` / `date` / `purchase_date`
- `quantity` / `qty` / `units`
- `customer_type` / `cust_type`

---

## Tech Stack
- **Backend**: Python 3, Flask, Pandas, NumPy, ReportLab, Matplotlib
- **Frontend**: HTML5, CSS3, Vanilla JavaScript, Chart.js
- **No AI / external APIs** — all analysis is pure Python
