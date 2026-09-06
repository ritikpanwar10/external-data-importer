
# Import Data from External Sources in Excel

A step-by-step technical reference for connecting, importing, combining, and automating data refreshes from external sources into Microsoft Excel using Power Query and native functions.

---

## 1. Import from Text / CSV File

Used for processing transactional logs, ERP reports, or external system exports delivered in `.csv` or `.txt` format.

### Step-by-Step Instructions:
1. Open Excel and navigate to the **Data** tab.
2. Select **Get Data** > **From File** > **From Text/CSV**.
3. Browse and select your target file (e.g., `orders_data.csv`).
4. In the preview window:
   * **File Origin:** `65001: Unicode (UTF-8)`
   * **Delimiter:** `Comma`
   * **Data Type Detection:** `Based on first 200 rows`
5. Click **Transform Data** to clean dates, filter records, or cast column types, or click **Load** to import directly into a table.

### Example Dataset (`orders_data.csv`):
```csv
Order_ID,Customer_Name,Product,Amount,Order_Date
1001,Amit Sharma,Keyboard,45.00,2026-08-15
1002,Sarah Jenkins,Monitor,220.50,2026-08-16
1003,Priya Patel,Mouse,15.75,2026-08-16

```

### Result:

* Creates a dynamic structured Excel table (`Table_OrdersData`).
* Preserves data types (Dates as native Excel dates, Amounts as currency/decimals).

---

## 2. Import from a Live Web Page (Web Scraping)

Used for scraping HTML tables from websites, such as live exchange rates, commodity prices, or public indicators.

### Step-by-Step Instructions:

1. Go to the **Data** tab.
2. Select **Get Data** > **From Other Sources** > **From Web**.
3. Enter the target URL:
```text
[https://en.wikipedia.org/wiki/List_of_countries_by_GDP_(nominal](https://en.wikipedia.org/wiki/List_of_countries_by_GDP_(nominal))

```


4. Click **OK**.
5. In the **Navigator** dialog:
* Expand the URL tree.
* Select the table from the list (e.g., `Table 1` or `By country/territory`).
* Verify the preview matches the web table.


6. Click **Transform Data** to remove header rows or invalid symbols, then click **Close & Load**.

### Refresh Mechanism:

* Right-click anywhere in the imported table and select **Refresh**.
* Excel connects to the URL, checks for layout changes, and pulls updated numbers automatically.

---

## 3. Import from a Relational Database (SQL Server / PostgreSQL / MySQL)

Used for extracting direct, filtered views from enterprise database servers without exporting flat files.

### Step-by-Step Instructions:

1. Go to **Data** > **Get Data** > **From Database** > select your database (e.g., **From SQL Server Database** or **From PostgreSQL Database**).
2. Enter connection details:
* **Server:** `localhost:5432` or `sql-server.company.internal`
* **Database:** `analytics_db`


3. Expand **Advanced options** and input a SQL statement to filter records on the database engine prior to loading:

```sql
SELECT 
    order_id,
    customer_id,
    order_date,
    SUM(quantity * unit_price) AS total_order_value
FROM orders
WHERE order_date >= '2026-01-01'
GROUP BY order_id, customer_id, order_date;

```

4. Authenticate using Windows or Database credentials.
5. Click **Load** to stream the dataset into an Excel Data Model or Worksheet Table.

---

## 4. Consolidate Multiple Files from a Folder

Used for combining identical periodic files (e.g., `Jan_Sales.xlsx`, `Feb_Sales.xlsx`, `Mar_Sales.xlsx`) into a single master dataset automatically.

### Step-by-Step Instructions:

1. Place all source files with identical schema into a dedicated folder:
```text
C:\DataProjects\MonthlySales\

```


2. In Excel: **Data** > **Get Data** > **From File** > **From Folder**.
3. Select the folder path and click **Open**.
4. In the preview window, select **Combine** > **Combine & Transform Data**.
5. Choose the sample sheet (e.g., `Sheet1`) as the template for all files.
6. Click **OK**.
7. In the Power Query editor, remove the auto-generated `Source.Name` column if not needed, then click **Close & Load**.

### Automation Benefit:

* Adding `Apr_Sales.xlsx` to the folder requires zero manual copy-pasting.
* Clicking **Data** > **Refresh All** appends the new file's rows into the master table automatically.

---

## 5. Live Web API Data Using Native Formulas (`WEBSERVICE` + `FILTERXML`)

Used for pulling single-cell live data points (such as currency rates, weather, or inventory counts) via REST APIs directly inside formulas without opening Power Query.

### Functions:

* `=WEBSERVICE(url)`: Fetches raw text, JSON, or XML data from a web endpoint.
* `=FILTERXML(xml, xpath)`: Parses an XML string and returns the target node using an XPath query.

### Implementation Example:

Cell `A1` contains the API endpoint:

```text
[https://api.worldbank.org/v2/country/IND/indicator/SP.POP.TOTL?date=2024](https://api.worldbank.org/v2/country/IND/indicator/SP.POP.TOTL?date=2024)

```

Cell `B1` formula to extract the indicator value:

```excel
=FILTERXML(WEBSERVICE(A1), "//wb:value")

```

### Result:

The cell immediately evaluates to the numerical value returned by the endpoint.

```

---

**Recommended Repository Structure**

```text
excel-external-data-connectors/
│
├── README.md
├── sample-data/
│   ├── sample_orders.csv
│   └── monthly-reports/
│       ├── Jan_Sales.xlsx
│       └── Feb_Sales.xlsx
├── sql-scripts/
│   └── query_extract_orders.sql
└── workbooks/
    ├── CSV_Import_Demo.xlsx
    ├── Folder_Consolidation_Demo.xlsx
    └── Web_Scraping_Demo.xlsx

```
