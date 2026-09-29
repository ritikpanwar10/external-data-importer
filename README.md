
# Import Data from External Sources in Excel

A step-by-step technical reference for connecting, importing, combining, and automating data refreshes from external sources into Microsoft Excel using Power Query and native functions.

---

## 1. Import from Text / CSV File

Used for processing transactional logs, ERP reports, or external system exports delivered in `.csv` or `.txt` format.

### Step-by-Step Instructions:
1. Open Excel and navigate to the **Data** tab.
2. Select **Get Data** > **From File** > **From Text/CSV**.


   <img width="955" height="656" alt="1" src="https://github.com/user-attachments/assets/b7dc15fa-eaf3-4fa4-b81f-0ffdc761007e" />

4. Browse and select your target file (e.g., `orders_data.csv`).


   <img width="930" height="583" alt="2" src="https://github.com/user-attachments/assets/46cf5ad7-c13f-43b4-abe7-7dcab04073f8" />

   
5. In the preview window:
   * **File Origin:** `65001: Unicode (UTF-8)`
   * **Delimiter:** `Comma`
   * **Data Type Detection:** `Based on first 200 rows`
6. Click **Transform Data** to clean dates, filter records, or cast column types, or click **Load** to import directly into a table.


<img width="1082" height="812" alt="3" src="https://github.com/user-attachments/assets/61aeb343-1b49-4c6a-b893-334d5ffe5b4f" />

### Example Dataset (`orders_data.csv`):

<img width="617" height="270" alt="4" src="https://github.com/user-attachments/assets/9940403e-e900-4e71-b432-098b804ff499" />


### Result:

* Creates a dynamic structured Excel table (`Table_OrdersData`).
* Preserves data types (Dates as native Excel dates, Amounts as currency/decimals).

---

## 2. Import from a Live Web Page (Web Scraping)

Used for scraping HTML tables from websites, such as live exchange rates, commodity prices, or public indicators.

### Step-by-Step Instructions:

1. Go to the **Data** tab.
2. Select **Get Data** > **From Other Sources** > **From Web**.


<img width="837" height="863" alt="5" src="https://github.com/user-attachments/assets/f2354d7d-45e7-4766-a46c-65af82b32503" />
 
3. Enter the target URL:
[https://en.wikipedia.org/wiki/List_of_countries_by_GDP_(nominal)](https://en.wikipedia.org/wiki/List_of_countries_by_GDP_(nominal))

4. Click **OK**.


<img width="871" height="267" alt="6" src="https://github.com/user-attachments/assets/6fc8a1b2-d715-4961-8c01-80a6563a4a6c" />

  
5. In the **Navigator** dialog:
* Expand the URL tree.
* Select the table from the list (e.g., GDP forecast or estimate (million US$) by country).
* Verify the preview matches the web table.


<img width="1093" height="871" alt="7" src="https://github.com/user-attachments/assets/39a39952-f497-4f74-9e91-d04d1ce63fac" />

6. Click **Transform Data** to remove header rows or invalid symbols, then click **Close & Load**.


<img width="1726" height="835" alt="8" src="https://github.com/user-attachments/assets/a3bb0968-0457-4066-bb13-1eebfdb9f15c" />


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
