# Excel Minor Project – Sales Operations Analytics

## Project Overview

This project is an **Excel-based Sales Operations Analytics project** designed to analyze sales data for a fast-growing multi-category retailer operating across **6 Indian cities**.

The project uses multiple interconnected datasets such as:

* Orders
* Customers
* Products
* Sales Representatives
* Monthly Sales Targets

The main objective is to use **Microsoft Excel data-analysis techniques** to clean, transform, analyze, and extract meaningful business insights from raw sales data.

---

## Project Scenario

The sales analytics team receives raw data from different business areas and needs to answer important questions before a leadership review.

The project focuses on:

* Data lookup and joining
* Data cleaning
* Conditional aggregation
* Pivot table analysis
* Sales representative performance
* Revenue analysis
* Target analysis
* Advanced Excel functions

---

## Dataset Structure

The workbook contains the following sheets:

| Sheet            | Description                                        |
| ---------------- | -------------------------------------------------- |
| `Orders`         | Contains order-level sales transactions            |
| `Customers`      | Contains customer information                      |
| `Products`       | Contains product and category information          |
| `Sales_Reps`     | Contains sales representative details              |
| `Targets`        | Contains monthly sales targets for representatives |
| `Q1_Lookups`     | Questions related to Excel lookup functions        |
| `Q2_Cleaning`    | Data cleaning and quality analysis                 |
| `Q3_Aggregation` | Conditional aggregation questions                  |
| `Q4_Pivot`       | Pivot table analysis                               |
| `Q5_Advanced`    | Advanced Excel analysis                            |

---

## Tools & Technologies

* **Microsoft Excel**
* VLOOKUP
* INDEX & MATCH
* SUMIFS
* COUNTIFS
* AVERAGEIFS
* IF
* RANK.EQ
* LARGE
* INDEX/MATCH
* TEXTJOIN
* Pivot Tables
* Data Cleaning
* Conditional Aggregation
* Data Validation
* Business Analytics

---

# Project Tasks

## Q1 – Lookup Functions

The first section focuses on retrieving information from different datasets.

Techniques used:

* VLOOKUP
* INDEX/MATCH
* Two-way lookup
* Nested/chain lookup
* Approximate lookup
* Self-join concepts

### Example

Finding the customer associated with a particular order:

`Order ID → Customer ID`

Finding a product name:

`Order ID → Product ID → Product Name`

---

## Q2 – Data Cleaning & Quality Diagnostics

This section focuses on identifying and handling data-quality problems.

The analysis includes:

* Duplicate orders
* Leading/trailing whitespace
* Inconsistent city capitalization
* Broken customer references
* Discontinued products
* Anomalous data patterns

### Techniques Used

* COUNTIF
* COUNTIFS
* LEN
* TRIM
* PROPER
* UPPER
* MATCH
* ISNA
* ISERROR

---

## Q3 – Conditional Aggregation

This section analyzes sales using multiple conditions.

Business questions include:

* Revenue generated during specific quarters
* Sales from specific cities
* Orders with high discounts
* Completed payments
* Revenue by product category
* Revenue by customer segment
* Top-performing sales representatives

### Functions Used

```text
SUMIFS
COUNTIFS
AVERAGEIFS
IF
SUMPRODUCT
LARGE
```

---

## Q4 – Pivot Table Analysis

Pivot Tables are used to summarize large amounts of sales data.

The analysis includes:

* Sales region vs. product category
* Monthly revenue by sales representative
* Percentage contribution by category
* Quarterly revenue
* Cross-sheet analysis

### Pivot Table Features

* Row fields
* Column fields
* Value fields
* Filters
* Percentage of row total
* Date grouping
* Calculated analysis

---

## Q5 – Advanced Excel Analysis

The final section uses advanced Excel techniques to answer business questions.

Topics include:

* Sales representative ranking
* Cumulative revenue
* Top-N analysis
* Revenue-based product ranking
* TEXTJOIN
* LARGE
* INDEX/MATCH

### Example

Sales representatives can be ranked according to their total revenue:

```text
Highest Revenue → Rank 1
Second Highest → Rank 2
...
```

The project also identifies top-performing representatives and products based on revenue.

---

# Data Analysis Workflow

```text
Raw Data
   ↓
Data Inspection
   ↓
Data Cleaning
   ↓
Lookup & Data Joining
   ↓
Helper Columns
   ↓
Conditional Aggregation
   ↓
Pivot Table Analysis
   ↓
Advanced Analysis
   ↓
Business Insights
```

---

# Key Learning Outcomes

Through this project, the following skills were developed:

1. Understanding structured business datasets.
2. Performing data lookups across multiple Excel sheets.
3. Identifying and diagnosing data-quality problems.
4. Cleaning inconsistent data.
5. Using conditional aggregation functions.
6. Creating and interpreting Pivot Tables.
7. Ranking sales representatives based on revenue.
8. Performing Top-N analysis.
9. Working with multiple related datasets.
10. Applying Excel to real-world business analytics problems.

---

# Business Use Cases

The analysis can help management answer questions such as:

* Which sales representatives generate the most revenue?
* Which products contribute the most to sales?
* Which cities generate higher revenue?
* How much revenue is generated by each category?
* Which representatives are performing strongly?
* What percentage of regional revenue comes from a particular category?
* What are the quarterly revenue trends?
* Where are data-quality issues affecting analysis?

---

# Project Files

```text
Excel-Minor-Project/
│
├── Excel_Minor_Project_5.xlsx
│
└── README.md
```

---

# How to Use

1. Download the Excel workbook.
2. Open `Excel_Minor_Project_5.xlsx` using Microsoft Excel.
3. Review the raw datasets:

   * Orders
   * Customers
   * Products
   * Sales_Reps
   * Targets
4. Explore the question sheets:

   * Q1_Lookups
   * Q2_Cleaning
   * Q3_Aggregation
   * Q4_Pivot
   * Q5_Advanced
5. Apply the required Excel formulas and analysis techniques.
6. Use the results to derive business insights.

---

# Project Objective

The overall objective is to demonstrate how **Excel can be used as a complete data-analysis tool** for transforming raw sales data into useful business information.

The project combines **data cleaning, lookup functions, aggregation, Pivot Tables, ranking, and advanced Excel functions** to support data-driven sales decisions.

---

## Author

**Sri Sai Nandini**

B.Tech – Artificial Intelligence & Data Science

---

## Project Type

**Excel Minor Project – Sales Operations Analytics**

**Domain:** Data Analytics / Business Intelligence

**Tool:** Microsoft Excel
