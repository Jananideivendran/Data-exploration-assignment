
# 📊 Product Data Exploration Using Excel
---
# Excel Assignment 1: Data Exploration & Analysis

## 📌 Project Overview
As a Data Analyst, this project requires performing basic data exploration and analysis on a **Product Dataset** in Microsoft Excel. The dataset includes information about different products, specifically their Product IDs, Product Name, Brand Name, Quantity, Categories, and Prices. 

The goal of this assignment is to build foundational skills in data summarization, statistical evaluation, logical functions, conditional aggregation, and text manipulation.


## 📂 Dataset Description

The dataset contains information about various products and their attributes.

### Dataset Attributes

| Attribute | Description |
|---|---|
| Product ID | A unique identifier assigned to each product |
| Product Name | The name of the product |
| Brand Name | The brand associated with the product |
| Quantity | The available quantity of each product |
| Category | The category to which the product belongs |
| Price | The price of each product |

---

## 🎯 Project Objectives

- Explore and understand the product dataset.
- Summarize product quantities and prices.
- Apply logical functions to categorize products based on conditions.
- Perform conditional calculations using Excel functions.
- Extract meaningful information from text fields.
- Develop foundational data exploration and analysis skills.
- Prepare data for further analysis and reporting.

---

## 🛠️ Tools Used

- **Microsoft Excel**
- Excel Formulas and Functions

---

## 🛠️ Assignment Tasks & Implemented Formulas

### 1. Basic Data Exploration (Sum, Count, and Average)
Perform the following calculations using Excel built-in functions:
* **Total Price:** Calculate the total price of all products in the dataset.
  * *Formula:* `=SUM(Price_Range)`
* **Product Count:** Determine how many products exist in the dataset.
  * *Formula:* `=COUNTA(Product_Name_Range)`
* **Average Price:** Calculate the average price of the products.
  * *Formula:* `=AVERAGE(Price_Range)`

### 2. Minimum and Maximum Values
Using Excel functions:
* **Minimum Price:** Identify the minimum price among all products.
  * *Formula:* `=MIN(Price_Range)`
* **Maximum Price:** Identify the maximum price among all products.
  * *Formula:* `=MAX(Price_Range)`

### 3. Logical Function – `IF`
* Create a new column named `Price Range`.
* Use the `IF` function to categorize products based on their price:
  * If Price $\ge$ $500$ $\rightarrow$ classify as **High Price**
  * If Price $<$ $500$ $\rightarrow$ classify as **Standard Price**
* *Formula Example:* `=IF(Price_Cell >= 500, "High Price", "Standard Price")`

### 4. Conditional Functions – `SUMIF` and `COUNTIF`
Perform the following conditional calculations:
* **Electronics Category Total:** Use the `SUMIF` function to calculate the total price of products in the Electronics category.
  * *Formula:* `=SUMIF(Category_Range, "Electronics", Price_Range)`
* **Products Under $100 Count:** Use the `COUNTIF` function to determine the number of products with a price less than $100.
  * *Formula:* `=COUNTIF(Price_Range, "<100")`

### 5. Text Formatting – `LEFT`, `RIGHT`, `MID`
Use text functions to extract information from the `Product ID` column. Create the following new columns:
* **Day:** Create a new column named `Day` with the first 2 characters of each `Product ID` using the `LEFT` function.
  * *Formula:* `=LEFT(ProductID_Cell, 2)`
* **Country Code:** Create a new column named `Country Code` by extracting the last 2 characters from the `Product ID` column using the `RIGHT` function.
  * *Formula:* `=RIGHT(ProductID_Cell, 2)`
* **Month:** Create a new column named `Month` by extracting the 4th to 6th characters from the `Product ID` column using the `MID` function.
  * *Formula:* `=MID(ProductID_Cell, 4, 3)`

---

## 📦 Deliverables Checklist
Submit a project folder containing:
1. **Excel File (.xlsx):**
   * Completed calculations using formulas
   * Newly created columns (`Price Range`, `Day`, `Country Code`, `Month`)
   * Correct application of Excel functions and clearly labeled columns/results
2. **PDF Document:**
   * Include screenshots of calculated results for each function, newly created columns, and applied formulas in the formula bar.
---

## 🔍 Key Analysis Areas

The assignment focuses on the following areas:

- Calculating total and average product prices.
- Summarizing product quantities.
- Categorizing products using logical conditions.
- Counting products belonging to specific categories or brands.
- Calculating totals based on product categories or brands.
- Extracting useful information from product names and IDs.
- Improving understanding of Excel-based data exploration.

---

## 📚 Key Learnings

Through this project, I practiced:

- Working with structured datasets in Excel.
- Applying mathematical and statistical functions.
- Using logical functions to categorize data.
- Performing conditional calculations.
- Manipulating text fields to extract information.
- Developing analytical thinking and problem-solving skills.
- Building a foundation for more advanced data analysis projects.

---

## 🚀 Conclusion

This project helped me strengthen my foundational Excel skills and understand
how formulas and functions can be used to explore structured product data.

By practicing data summarization, logical functions, conditional aggregation,
and text manipulation, I developed essential skills that can be applied to
real-world data analysis tasks.

This project represents one of my initial steps toward building a portfolio
for an entry-level Data Analyst role.

---

## 👩‍💻 Author

 Janani Aspiring Data Analyst | Excel | SQL | Power BI | Python
