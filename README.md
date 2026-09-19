Markdown# 📊 Excel Data Exploration & Analysis — Product Dataset

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-Data%20Analytics-blue?style=for-the-badge)
![Level](https://img.shields.io/badge/Skill%20Level-Beginner%20to%20Intermediate-orange?style=for-the-badge)

---

## 📌 Project Overview
This project is part of my **Data Analytics Portfolio** under **Module 1: Excel Data Exploration**. 

As a Data Analyst, basic data exploration, conditional aggregation, and string manipulation are essential foundational steps before performing advanced statistical analysis or visualization. This project focuses on analyzing an e-commerce Product Dataset to derive summary statistics, apply business logic, perform conditional calculations, and extract metadata from product identifiers using Microsoft Excel formulas.

---

## 🎯 Key Objectives
- **Data Summarization:** Calculate key descriptive metrics including total spend, product counts, and average item prices.
- **Data Range & Statistical Analysis:** Identify price extremities (minimum and maximum price points).
- **Logical Segmentation:** Apply conditional logic (`IF`) to categorize inventory based on pricing tiers.
- **Conditional Aggregation:** Filter and summarize metrics by specific attributes using `SUMIF` and `COUNTIF`.
- **Text Parsing & Transformation:** Extract structured temporal and geographic attributes from alphanumeric `Product ID` fields using `LEFT`, `RIGHT`, and `MID`.

---

## 📂 Repository Structure

├── 📁 Screenshots/                  # PDF/Image deliverables showing formulas and outputs│   ├── basic_calculations.png│   ├── logical_if_columns.png│   └── text_parsing_functions.png├── 📄 Excel_Assignment_1_Data_Exploration.xlsx  # Complete Excel Workbook with live formulas└── 📄 README.md                     # Portfolio Documentation
---

## 📋 Dataset Architecture

The raw dataset contains **34 product items** across various categories (Electronics, Fashion, Kitchen, Outdoor) with the following original attributes:

| Column Name | Data Type | Description | Example |
| :--- | :--- | :--- | :--- |
| `Product ID` | Text | Unique alphanumeric code formatted as `DD-MMM-CC` | `28-JAN-US` |
| `Product Name` | Text | Name of the consumer product | `Laptop` |
| `Brand Name` | Text | Manufacturer / Brand | `Dell` |
| `Price ($)` | Currency | Unit price in USD ($) | `$1,000` |
| `Quantity` | Integer | Stock inventory quantity | `30` |
| `Category` | Text | Product classification category | `Electronics` |

---

## 🛠️ Tasks, Formulas & Key Findings

### 1. Basic Data Exploration & Descriptive Statistics
To understand overall inventory pricing, core summary metrics were calculated across the dataset range (`D2:D35`).

* **Total Catalog Price:** `$10,100`  
  * Formula: `=SUM(D2:D35)`
* **Total Products Count:** `34`  
  * Formula: `=COUNTA(B2:B35)`
* **Average Product Price:** `$297.06`  
  * Formula: `=AVERAGE(D2:D35)`

---

### 2. Minimum and Maximum Range Analysis
Identifying price bounds to understand baseline and peak product costs.

* **Minimum Product Price:** `$30` (T-shirt / Adidas)  
  * Formula: `=MIN(D2:D35)`
* **Maximum Product Price:** `$1,000` (Laptop / Dell)  
  * Formula: `=MAX(D2:D35)`

---

### 3. Logical Classification (`IF`)
Categorized products into pricing tiers based on a `$500` threshold to assist pricing strategy segmentation.

* **Logic Rule:**
  * If $\text{Price} \ge \$500 \rightarrow$ **"HIGH PRICE"**
  * If $\text{Price} < \$500 \rightarrow$ **"STANDARD PRICE"**
* **Excel Formula (Cell `G2`):**
  ```excel
  =IF(D2>=500, "HIGH PRICE", "STANDARD PRICE")
4. Conditional Aggregation (SUMIF & COUNTIF)Performed targeted aggregation for business reporting.Total Price of Electronics Category: $8,050Formula: =SUMIF(F2:F35, "Electronics", D2:D35)Count of Products Under $100: 11 productsFormula: =COUNTIF(D2:D35, "<100")5. Text Extraction & Feature EngineeringExtracted embedded metadata from the Product ID string (e.g., 28-JAN-US) to engineer new analytical features:New FeatureExtraction GoalExcel FormulaOutput Example (28-JAN-US)DayFirst 2 characters from the left=LEFT(A2, 2)28Country CodeLast 2 characters from the right=RIGHT(A2, 2)USMonth3 characters starting at 4th position=MID(A2, 4, 3)JAN
