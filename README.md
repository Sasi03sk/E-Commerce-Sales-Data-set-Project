# 🛒 E-Commerce Sales Analysis using Microsoft Excel

## 📌 Project Overview

This project is an **E-Commerce Sales Analysis project developed using Microsoft Excel** as part of my Data Analytics learning journey.

The objective of this project is to transform raw e-commerce data into meaningful business insights by performing **data cleaning, data preparation, Excel formula analysis, XLOOKUP/VLOOKUP, descriptive statistics, Pivot Table analysis, PivotCharts, and interactive dashboard development**.

This project demonstrates my practical understanding of the **end-to-end data analysis process using Excel** and is one of the projects in my growing Data Analyst portfolio.

---

## 🎯 Project Objectives

The main objectives of this project are:

- Import and organize an E-Commerce dataset in Excel
- Clean and prepare raw data for analysis
- Identify and remove duplicate records
- Handle missing and inconsistent values
- Correct formatting errors
- Validate the dataset
- Remove irrelevant columns where applicable
- Apply Excel formulas for calculations and analysis
- Use **XLOOKUP/VLOOKUP** to combine related tables
- Create Pivot Tables for business analysis
- Calculate descriptive statistics
- Create different types of charts
- Build an interactive Excel dashboard
- Create PivotCharts for visual analysis
- Generate business insights using:
  - Descriptive Analysis
  - Diagnostic Analysis
  - Predictive Analysis
  - Prescriptive Analysis

---

## 📂 Dataset

The dataset contains E-Commerce sales and customer-related information that can be analyzed across different dimensions such as:

- Product
- Category
- Sales
- Quantity
- Region
- Store Type
- Age Group
- Gender
- Payment Type

The raw dataset was imported into Excel and prepared for further analysis.

---

## 🔄 Data Analysis Workflow

The project follows a structured Data Analytics workflow:

```text
Raw Dataset
     ↓
Data Import
     ↓
Data Cleaning
     ↓
Data Validation
     ↓
Data Preparation
     ↓
Excel Formulas
     ↓
XLOOKUP / VLOOKUP
     ↓
Descriptive Statistics
     ↓
Pivot Tables
     ↓
PivotCharts
     ↓
Dashboard
     ↓
Business Insights
```

---

# 🧹 1. Data Cleaning & Preparation

The first stage of the project focused on preparing the raw dataset for analysis.

### Data Cleaning Activities

- Checked for duplicate records
- Removed duplicate data where required
- Identified missing values
- Checked inconsistent values
- Corrected formatting issues
- Standardized data formats
- Validated important columns
- Checked numerical and text fields
- Removed irrelevant columns where applicable

The purpose of this stage was to ensure that the dataset was reliable and ready for analysis.

---

# 🧮 2. Excel Formulas & Functions

Several Excel formulas and functions were used to perform calculations, transformations, and data analysis.

### Examples of Excel Functions Used

- `SUM`
- `AVERAGE`
- `MIN`
- `MAX`
- `COUNT`
- `COUNTA`
- `MEDIAN`
- `MODE`
- `STDEV`
- `IF`
- `AND`
- `CONCAT`
- `LOWER`
- `TRIM`
- `LOG10`
- `PI`
- `NOW`
- `POWER`
- `ROUNDDOWN`

These functions were used for numerical calculations, logical operations, text manipulation, data cleaning, and statistical analysis.

---

# 🔎 3. XLOOKUP / VLOOKUP

Lookup functions were used to combine information from different tables.

### XLOOKUP

Example structure:

```excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

### VLOOKUP

Example structure:

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```

These functions helped connect related information between tables and create a more complete dataset for analysis.

---

# 📊 4. Pivot Table Analysis

Pivot Tables were created to summarize the E-Commerce dataset and identify important business patterns.

### Key Pivot Table Analyses

1. **Category vs Sales**
2. **Region vs Store Type**
3. **Age Group vs Gender**
4. **Payment Type Analysis**
5. **Additional sales and quantity comparisons**

Pivot Tables were used to analyze:

- Total Sales
- Quantity Sold
- Customer Segments
- Regions
- Store Types
- Payment Methods
- Product Categories

---

# 📈 5. Descriptive Statistics

Descriptive statistics were calculated to understand the distribution and characteristics of the dataset.

### Statistics Used

- Mean
- Median
- Mode
- Minimum
- Maximum
- Standard Deviation
- Count

These calculations helped understand the **central tendency and spread of the sales data**.

---

# 📉 6. Data Visualization

Different visualization techniques were used to communicate the analysis clearly.

### Charts Created

- Column Chart
- Bar Chart
- Pie/Donut Chart
- PivotCharts

Charts were used to compare sales, quantities, categories, regions, customer segments, and payment methods.

---

# 📊 7. Interactive Dashboard

An interactive Excel dashboard was developed to provide a quick overview of the E-Commerce business performance.

### Dashboard Components

- KPI / Sales Summary
- Category Analysis
- Regional Analysis
- Store Type Analysis
- Customer Demographics
- Payment Method Analysis
- PivotCharts
- Interactive Slicers

The dashboard allows users to filter and explore the data interactively.

---

# 💡 8. Business Insights

The analysis generated several important business insights.

## Category vs Sales

### Descriptive

- **Accessories** generated the highest sales at approximately **₹3,66,401.82**.
- **Grooming** generated the lowest sales at approximately **₹21,748.59**.
- Accessories also had the highest quantity sold with **837 units**.

### Diagnostic

Higher sales in Accessories are associated with the higher quantity sold.

### Predictive

If the existing trend continues, Accessories may continue to remain a high-performing category.

### Prescriptive

- Maintain sufficient stock for Accessories.
- Promote lower-performing categories such as Grooming.
- Investigate the reasons behind low category performance.

---

## Region vs Store Type

### Descriptive

- **North** recorded the highest regional sales at approximately **₹6,43,578.69**.
- **West** recorded the lowest regional sales at approximately **₹2,70,609.03**.
- **Online** generated the highest sales among store types at approximately **₹9,14,963.12**.

### Diagnostic

- North performed strongly across Flagship and Outlet stores.
- South showed strong Online sales.
- West had comparatively lower overall sales.

### Predictive

North and South may continue to perform well if current trends remain consistent.

### Prescriptive

- Improve sales performance in the West region.
- Continue focusing on Online sales.
- Identify the best-performing store type within each region.

---

## Age Group vs Gender

### Descriptive

- **Adults** generated the highest sales at approximately **₹15,65,705.27**.
- Female customers generated approximately **₹12,31,593.68** in sales.
- Male customers generated approximately **₹9,41,481.95** in sales.

### Diagnostic

Adults generated the highest sales because they also represented the highest number of orders.

Female customers generated more sales than male customers.

### Predictive

Adults may continue to represent the primary customer segment.

### Prescriptive

- Focus marketing campaigns on Adult customers.
- Create targeted offers for Youth and Senior customers.
- Analyze product preferences across male and female customers.

---

## Payment Type Analysis

### Descriptive

- **PayPal** recorded the highest sales at approximately **₹7,55,679.41**.
- Credit Card sales were approximately **₹7,15,582.14**.
- COD sales were approximately **₹7,01,814.08**.

### Diagnostic

PayPal had the highest sales and number of orders among the payment methods analyzed.

All three payment methods were regularly used.

### Predictive

PayPal may continue to remain an important payment method if the current pattern continues.

### Prescriptive

- Continue providing all major payment options.
- Promote preferred payment methods.
- Analyze whether payment method influences customer ordering behavior.

---

# 🛠️ Tools & Technologies

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Data cleaning, analysis & visualization |
| **Excel Tables** | Data organization |
| **Excel Formulas** | Calculations & transformations |
| **XLOOKUP / VLOOKUP** | Combining tables |
| **Pivot Tables** | Data summarization |
| **PivotCharts** | Data visualization |
| **Slicers** | Interactive filtering |
| **Dashboard** | Business reporting |
| **Descriptive Statistics** | Statistical analysis |

---

# 📁 Project Structure

```text
E-Commerce-Sales-Analysis/
│
├── README.md
│
├── Dataset/
│   └── Ecommerce_Sales_Dataset.xlsx
│
├── Excel/
│   └── Ecommerce_Sales_Analysis.xlsx
│
├── Dashboard/
│   └── Ecommerce_Dashboard.png
│
├── Insights/
│   └── E-Commerce_Project_Insights.pdf
│
└── Screenshots/
    ├── Data_Cleaning.png
    ├── XLOOKUP.png
    ├── Pivot_Tables.png
    ├── Charts.png
    └── Dashboard.png
```

---

# 📌 Key Skills Demonstrated

Through this project, I demonstrated practical skills in:

- Data Cleaning
- Data Preparation
- Data Validation
- Excel Data Analysis
- Excel Functions
- XLOOKUP
- VLOOKUP
- Pivot Tables
- PivotCharts
- Descriptive Statistics
- Data Visualization
- Dashboard Development
- Business Intelligence
- Business Insights
- Analytical Thinking

---

# 🚀 What I Learned

This project helped me understand how a Data Analyst approaches a dataset from beginning to end.

I learned how to:

1. Clean and prepare raw data.
2. Validate data before analysis.
3. Use Excel formulas to transform and analyze data.
4. Combine information using XLOOKUP/VLOOKUP.
5. Summarize large datasets using Pivot Tables.
6. Apply descriptive statistics.
7. Build meaningful visualizations.
8. Create an interactive dashboard.
9. Convert data into actionable business insights.
10. Present analytical findings using a **Descriptive → Diagnostic → Predictive → Prescriptive** approach.

---

# 🎯 Future Improvements

As I continue developing my Data Analyst portfolio, I plan to extend this project by:

- Recreating the analysis using **SQL**
- Building an interactive **Power BI dashboard**
- Performing deeper statistical analysis using **Python**
- Adding automated data-cleaning workflows
- Developing advanced business KPIs
- Performing customer segmentation
- Performing sales trend and forecasting analysis

---

# 👨‍💻 About Me

I am an aspiring **Data Analyst** building my portfolio through practical projects and continuously developing my skills in:

**Excel | SQL | Python | Power BI | Data Visualization | Statistics | AI Tools**

I am interested in transforming raw data into meaningful insights that can support better business decisions.

---

## ⭐ Portfolio

This project is part of my Data Analyst portfolio and demonstrates my practical application of **Microsoft Excel for data cleaning, analysis, visualization, dashboard development, and business insights**.

If you find this project useful, feel free to ⭐ **star this repository**.

---
