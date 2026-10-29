
# 📊 Microsoft Excel – Data Exploration and Analysis

## 📌 Project Overview

This project focuses on performing basic data exploration and analysis using **Microsoft Excel**. The dataset contains product information, including Product ID, Product Name, Brand Name, Price, Quantity, and Category.

The objective is to develop foundational Excel skills by applying formulas and functions to perform calculations, categorize products, summarize data, and extract meaningful information from text fields.

## 🎯 Problem Statement

As a Data Analyst, the task is to analyze a product dataset using Microsoft Excel. The assignment focuses on data summarization, logical functions, conditional aggregation, and text manipulation.

The analysis helps transform raw product data into meaningful insights that can support data-driven decision-making.

## 📂 Dataset Description

The dataset contains the following columns:

| Column Name | Description |
|---|---|
| Product ID | Unique identifier for each product |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Price ($) | Price of the product in dollars |
| Quantity | Available quantity of the product |
| Category | Product category |

### Sample Dataset

| Product ID | Product Name | Brand Name | Price ($) | Quantity | Category |
|---|---|---|---:|---:|---|
| 28-JAN-US | Laptop | Dell | 1000 | 30 | Electronics |
| 15-FEB-US | Sneakers | Nike | 80 | 15 | Fashion |
| 03-MAR-US | Coffee Maker | Keurig | 130 | 40 | Kitchen |
| 11-APR-US | Smartphone | Samsung | 900 | 25 | Electronics |
| 22-MAY-US | Backpack | North Face | 70 | 20 | Outdoor |

## 🛠️ Tools and Technologies

- **Microsoft Excel**
- Excel Formulas and Functions
- Conditional Aggregation
- Logical Functions
- Text Manipulation
- Data Exploration and Analysis

## ✅ Assignment Objectives

### 1. Basic Data Exploration
- Understand the dataset structure.
- Identify product details and categories.
- Explore product prices and quantities.
- Summarize key information from the dataset.

### 2. Logical Functions
Use the `IF()` function to categorize products based on price.

**Condition:**
- If Price >= $500 → High Price
- If Price < $500 → Standard Price

**Excel Formula:**

```excel
=IF(D2>=500,"High Price","Standard Price")
```

### 3. Data Summarization
Use Excel functions to calculate key metrics, such as:

- Total quantity of products
- Average product price
- Highest product price
- Lowest product price
- Number of products in the dataset

**Example Formulas:**

```excel
=SUM(E2:E6)
=AVERAGE(D2:D6)
=MAX(D2:D6)
=MIN(D2:D6)
=COUNTA(A2:A6)
```

### 4. Conditional Aggregation
Use conditional functions such as `SUMIF()`, `COUNTIF()`, and `AVERAGEIF()` to analyze products based on specific criteria.

**Example Formulas:**

Calculate the total quantity of Electronics products:

```excel
=SUMIF(F2:F6,"Electronics",E2:E6)
```

Count the number of Electronics products:

```excel
=COUNTIF(F2:F6,"Electronics")
```

Calculate the average price of Electronics products:

```excel
=AVERAGEIF(F2:F6,"Electronics",D2:D6)
```

### 5. Text Manipulation
Use Excel text functions to extract useful information from the Product ID.

Examples of relevant functions:

- `LEFT()` – Extract characters from the beginning of a text string.
- `MID()` – Extract characters from a specified position.
- `RIGHT()` – Extract characters from the end of a text string.
- `FIND()` – Find the position of a character or substring.
- `LEN()` – Count the number of characters in a text string.

**Example Formulas:**

Extract the first two characters from the Product ID:

```excel
=LEFT(A2,2)
```

Extract the month code from the Product ID:

```excel
=MID(A2,4,3)
```

Extract the country code from the Product ID:

```excel
=RIGHT(A2,2)
```

## 📈 Expected Outcomes

By completing this assignment, the following skills can be developed:

- Understanding and exploring structured datasets.
- Applying Excel formulas to calculate summary statistics.
- Using logical functions to categorize products.
- Performing conditional calculations and aggregations.
- Extracting information from text fields.
- Improving data analysis and spreadsheet management skills.
- Preparing data for further analysis and reporting.



## 🚀 How to Use This Project

1. Clone or download this repository.
2. Open `Product_Data.xlsx` in Microsoft Excel.
3. Review the product dataset.
4. Apply the formulas and functions described in this README.
5. Verify the calculated results.
6. Save the completed Excel workbook.
7. Upload the workbook and any relevant screenshots to the repository.

## 💡 Key Learnings

This assignment provides practical experience with:

- Data exploration in Microsoft Excel.
- Basic statistical calculations.
- Logical decision-making using `IF()`.
- Conditional analysis using `SUMIF()`, `COUNTIF()`, and `AVERAGEIF()`.
- Text extraction using `LEFT()`, `MID()`, `RIGHT()`, `FIND()`, and `LEN()`.
- Organizing and preparing product data for analysis.

## 👩‍💻 Author

**Gomathy P**

Aspiring Data Analyst | Microsoft Excel | Data Exploration and Analysis

---

⭐ If you find this project useful, feel free to explore the repository and review the Excel formulas used in the analysis.

