# 📊 Excel Data Exploration & Analysis – Product Dataset

## 📌 Project Overview

This project is part of my journey toward becoming a **Data Analyst**. It focuses on performing basic data exploration and analysis using **Microsoft Excel**.

The project uses a product dataset containing information such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

The objective of this project is to build foundational data analysis skills by using Excel formulas and functions to summarize data, apply conditional logic, perform conditional calculations, and extract useful information from text fields.

---

## 🎯 Project Objectives

The main objectives of this project are to:

* Explore and summarize product data.
* Calculate basic statistical measures such as total, count, average, minimum, and maximum.
* Categorize products based on their price.
* Perform conditional calculations using `SUMIF` and `COUNTIF`.
* Extract information from Product IDs using text functions.
* Develop practical Excel skills required for entry-level data analysis.

---

## 📂 Dataset

### Dataset: Product Dataset

The dataset contains the following attributes:

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Number of units available          |
| Category     | Product category                   |
| Price        | Price of the product               |

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Formulas & Functions
* GitHub

### Excel Functions Used

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`

---

# 🔍 Analysis Tasks

## 1. Basic Data Exploration

The first step was to calculate key summary statistics for the product prices.

### Total Price

The `SUM` function was used to calculate the total price of all products.

```excel
=SUM(F2:F101)
```

### Number of Products

The `COUNT` function was used to determine the number of products in the dataset.

```excel
=COUNT(F2:F101)
```

### Average Price

The `AVERAGE` function was used to calculate the average product price.

```excel
=AVERAGE(F2:F101)
```

> **Note:** The cell range should be adjusted according to the actual size of the dataset.

---

## 2. Minimum and Maximum Price

The `MIN` and `MAX` functions were used to identify the lowest and highest product prices.

### Minimum Price

```excel
=MIN(F2:F101)
```

### Maximum Price

```excel
=MAX(F2:F101)
```

These calculations help provide a quick understanding of the price range within the dataset.

---

## 3. Logical Function – IF

A new column called **Price Range** was created to categorize products according to their price.

### Business Rule

* Price **≥ $500** → `High Price`
* Price **< $500** → `Standard Price`

### Formula

```excel
=IF(F2>=500,"High Price","Standard Price")
```

The formula was then applied to all product records.

This demonstrates how Excel's `IF` function can be used to create categories based on business rules.

---

## 4. Conditional Functions – SUMIF and COUNTIF

### Total Price of Electronics Products

The `SUMIF` function was used to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(E2:E101,"Electronics",F2:F101)
```

Where:

* `E2:E101` = Category column
* `"Electronics"` = Condition
* `F2:F101` = Price column

---

### Number of Products Below $100

The `COUNTIF` function was used to determine how many products have a price below $100.

```excel
=COUNTIF(F2:F101,"<100")
```

This provides a quick way to identify the number of lower-priced products in the dataset.

---

# 🔤 5. Text Extraction Using LEFT, RIGHT and MID

The Product ID contains information that can be extracted using Excel text functions.

Three new columns were created:

* **Day**
* **Country Code**
* **Month**

---

### Day

The first two characters of the Product ID were extracted using the `LEFT` function.

```excel
=LEFT(A2,2)
```

---

### Country Code

The last two characters of the Product ID were extracted using the `RIGHT` function.

```excel
=RIGHT(A2,2)
```

---

### Month

Characters 4 to 6 of the Product ID were extracted using the `MID` function.

```excel
=MID(A2,4,3)
```

Where:

* `A2` = Product ID
* `4` = Starting position
* `3` = Number of characters to extract

---

# 📈 Key Skills Demonstrated

Through this project, I practiced the following data analysis skills:

### Data Exploration

* Understanding dataset structure
* Summarizing numerical data
* Identifying minimum and maximum values

### Excel Functions

* Aggregation using `SUM`, `COUNT`, and `AVERAGE`
* Finding minimum and maximum values using `MIN` and `MAX`
* Conditional logic using `IF`
* Conditional aggregation using `SUMIF`
* Conditional counting using `COUNTIF`

### Text Analysis

* Extracting characters using `LEFT`
* Extracting characters from the end using `RIGHT`
* Extracting characters from specific positions using `MID`

### Data Preparation

* Creating calculated columns
* Categorizing records
* Extracting structured information from identifiers

---

# 📁 Project Structure

```text
Excel-Data-Exploration/
│
├── README.md
│
├── Excel Assignment 1 - Data Exploration.xlsx
│
└── screenshots/
    └── excel-analysis.png
```

---

# 📊 Project Outcome

This project helped me develop a practical understanding of how **Excel can be used for basic data exploration, data transformation, and analysis**.

The analysis demonstrates how simple Excel functions can turn raw product data into meaningful information that can support further analysis and reporting.

This project is also the **first step in my Data Analyst portfolio**, where I will continue building projects using Excel, SQL, Power BI, Python, and other data analytics tools.

---

# 🚀 Future Improvements

As I continue developing my data analytics skills, I plan to extend this project by:

* Creating Excel PivotTables
* Building interactive dashboards
* Adding charts and visualizations
* Performing category-level analysis
* Analyzing quantity and revenue-related metrics
* Identifying pricing trends
* Recreating the analysis using SQL
* Building a Power BI dashboard

---

# 👨‍💻 About Me

I am an aspiring **Data Analyst** currently building my skills through hands-on projects and practical datasets.

My learning journey focuses on developing skills in:

* 📊 Excel
* 🗄️ SQL
* 📈 Power BI
* 🐍 Python
* 📉 Data Visualization
* 🔍 Data Analysis

This repository documents my progress and projects as I work toward becoming a professional Data Analyst.

---

## ⭐ Portfolio

More data analytics projects will be added to this repository as I continue learning and developing my analytical skills.

**Thank you for visiting my project!**
