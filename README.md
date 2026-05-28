# excel-data-cleaning-project
Data cleaning and preparation project using Microsoft Excel. 



## Project Overview

This project involves auditing, cleaning, and transforming a raw, 14-column sales dataset manually using Microsoft Excel. The primary objective was to trace data inconsistencies, handle missing values intelligently, check for duplicates, incorrect data, resolve structural formatting issues, and ensure transactional integrity across all records so the dataset is reliable.

## Dataset Features

The dataset consists of 14 columns capturing end-to-end transaction details, including:
 ⁠Order ID⁠ & ⁠Customer ID⁠ (Unique identifiers)
 ⁠Date⁠ (Purchase timeline)
 ⁠Product⁠, ⁠Payment Method⁠, & ⁠Order Status⁠ (Categorical details)
 ⁠Tracking Number⁠ & ⁠Shipping Address⁠ (Logistics data)
 ⁠Referral Source⁠ (Marketing channel tracking)
 ⁠Items in Cart⁠, ⁠Quantity⁠, ⁠Unit Price⁠, & ⁠Total Price⁠ (Financial and quantitative metrics)
 ⁠Coupon Code⁠ (Promotional data)

## Manual Cleaning Workflow & Methodology

Step 1: Integrity Check & Duplicate Audit

Conducted a thorough audit across all 14 columns to check for duplicate records using Excel’s Conditional Formatting (Highlight Cells Rules > Duplicate Values) and the Remove Duplicates tool.
Result: Confirmed that the data maintained transactional integrity with 0 duplicate rows found, validating that each ⁠Order ID⁠ represented a unique sale.

Step 2: Temporal & Financial Format Standardization

 Date Column: Reviewed and manually adjusted all entries uniformly into a structured format using Excel's Format Cells > Date settings to ensure seamless chronological sorting.

 Financial Metrics: Standardized the currency and number formatting for ⁠Unit Price⁠ and ⁠Total Price⁠ columns to ensure numerical consistency.

Step 3: Categorical Text Standardization

 Product Column: Audited using Filters and standardized varying entries manually to ensure identical products grouped together perfectly.

 Payment Method & Order Status: Cleaned up erratic text casing by applying the ⁠=PROPER()⁠ formula in columns to convert text strings into standard proper casing. This eliminated issues where variations like "Debit Card", "debit card" would be treated as separate categories. The formulas were then pasted back as Values to clean the original columns.

Step 4: Intelligent Imputation for Missing Data

 Identified missing values within the Coupon Code column.
• Methodology Note: Since Coupon Code is a text-based (categorical attribute rather than numerical, standard statistical imputations like mean, median, or mode could not be applied.
• To preserve data integrity without introducing bias, Excel's Go To Special > Blanks feature was used to isolate the missing rows, and they were systematically filled with the string "No Coupon", explicitly denoting organic, non-discounted transactions.

Step 5: Structural & Aesthetic Polish

Formatted the header row manually with uniform background fills, distinct typography, and consistent casing to improve scannability and structural professionalism.


## Excel Functions & Features Applied
Text Transformation: =PROPER formula.

Data Auditing: Filters, Sorting, and Conditional Formatting to profile data anomalies.

Cell Isolation: Go to Special(Blanks) for efficient bulk data imputation.

Structural Formatting: Cell styles, alignment tools, and date serialization.


BEFORE VS. AFTER SUMMARY

|METRIC / ATTRIBUTE   | BEFORE CLEANING  |  AFTER CLEANING|
|:--- | :--- | :---|
|**Duplicates** | Checked 14 columns  | 0 Duplicates |
|**Date Format**| Inconsistent/Mixed  | Uniformly |
|**Text Casing** |Mixed |Proper Case |
|**Missing Data** | Blank Coupon Codes | Filled as "No Coupon" |
|**Header Layout** | Non-uniform styles | Clean and Consistent |


## Outcome

The dataset was successfully cleaned, structured, and standardized. It is now ready for exploratory data analysis  (EDA) and visualization to generate meaningful business insights.
