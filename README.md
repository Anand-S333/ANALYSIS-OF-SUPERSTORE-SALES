Project: Sales Dataset Analysis 
Objective: 
Analyze the dataset containing sales transaction information from a retail business. The goal 
is to perform comprehensive data analysis, extract meaningful insights, and create interactive 
visualizations using Microsoft Excel. Additionally, leverage advanced Excel features such as 
What-If Analysis, Goal Seek, Macros, Power Query, and Power Pivot to deepen your 
analysis and automate tasks. 
Instructions: 
1. Format the Worksheet: 
Format the columns for better readability. 
Use Date formatting for the "Order Date" and "Ship Date" columns to ensure 
proper sorting and filtering. 
Apply bold headers and freeze the top row to keep it visible as you scroll 
through the data. 
2. Clean the Data: 
Check for and remove any duplicate rows or blank entries in the dataset. 
Use data validation to ensure that all data entries (like sales numbers and dates) 
are consistent and free of errors. For example, ensure Sales and Discount are 
never negative. 
Use IF or IFERROR to flag and handle any negative or erroneous sales data that 
might result from data entry mistakes. 
Identify and resolve any outliers or incorrect values, especially in Sales and 
Discount. 
3. Analyze Sales Performance: 
Calculate the total sales for each transaction by applying discount on the sales. 
Create a new column titled "Adjusted Sales" to reflect this. Before doing this step, 
validate for any negative values. 
Determine the total revenue, average order value, and total discounts across all 
transactions.
Segment the data by product categories and calculate key metrics (e.g., total 
adjusted sales, total discounts) for each segment. 
4. Use Formulas to Extract Insights: 
Apply COUNTIF/COUNTIFS to count the number of orders for each Product 
Category or Customer Segment (if available) and group by these categories. 
Calculate the average discount per order using AVERAGEIF based on 
Customer Segment. 
5. Time-Based Analysis: 
Group the data by Date (daily, monthly, quarterly) using PivotTables to assess 
sales trends over time. 
Identify peak sales periods by analyzing the Adjusted Sales over different time 
intervals. Look for patterns that could indicate seasonal variations. 
Use Date Functions like MONTH(), YEAR(), WEEKDAY(), and EOMONTH() to 
categorize and analyze sales performance over different periods. 
6. Sales Channel and Product Type Analysis: 
Take Sales Channel related column like "Ship Mode", create a PivotTable to 
calculate the total Sales, average Sales, and total Discounts for each Sales 
Channel. 
Identify the best-selling products or product categories based on Quantity and 
Sales. 
7. What-If Analysis & Goal Seek: 
Use Goal Seek to determine the sales target needed to achieve a specific revenue 
or profit goal. For example, set a goal for Adjusted Sales or Profit and use Goal 
Seek to find out how many units of a product need to be sold to reach that target. 
Use What-If Analysis to simulate different scenarios such as: 
• How would a 10% increase in Sales impact total revenue? 
• What would happen if the Discount rate decreased by 5%? 
• Create a Scenario Manager to compare various sales and discount 
scenarios and their effect on Adjusted Sales or Profit. 
Advanced Excel | © KGiSL MicroCollege | Internal version 1.0 
8. Pivot Tables and Pivot Charts: 
Create PivotTables to aggregate data by different dimensions such as Product 
Category, Customer Segment, or Ship Mode. 
Use PivotCharts to visualize sales trends and compare performance across 
different categories (e.g., bar charts for sales by product category, line charts for 
sales over time). 
Add Slicers to your PivotTables and charts to allow interactive filtering by 
categories such as Product Category, Ship Mode, or Customer Segment. 
9. Create a Sales Dashboard: 
• Design a summary dashboard with key performance indicators (KPIs) such as: 
Total sales (Sales or Adjusted Sales) 
Average sales per order 
Total discounts 
Top-performing products or product categories 
• Use data validation to allow the user to select specific date ranges or product 
categories to dynamically update the dashboard. 
• Include interactive features such as slicers for Product Category, Ship Mode, or 
Customer Segment to allow users to filter and view data by specific segments. 
11. Automate Tasks Using Macros: 
• Record and run macros to automate repetitive tasks, such as: 
Formatting the worksheet (e.g., currency formatting, bolding headers, freezing 
panes). 
Generating summary tables or charts for Adjusted Sales or Sales by Category. 
Running calculations and formatting pivot tables automatically when new data is 
added. 
• Assign the macros to buttons on the worksheet to make it easy to run them. 
Advanced Excel | © KGiSL MicroCollege | Internal version 1.0 
12. Power Pivot & Data Modeling: 
• Create Two Separate Tables: 
• Sheet 2: Products Table 
Create a table with the following columns: 
Product ID 
Product Name 
Category 
Sub-Category 
Sales 
Quantity 
Discount 
Profit 
• Sheet 3: Customers Table 
Create a table with the following columns: 
Customer ID 
Customer Name 
Segment 
Country 
City 
State 
Postal Code 
Region 
• Use Power Pivot to create a data model by establishing a relationship between the two 
tables: 
Advanced Excel | © KGiSL MicroCollege | Internal version 1.0 
Connect the Product ID column in the Products Table with the Product ID in 
the Sales transactions table (or any relevant data that has product-related 
sales). 
Connect the Customer ID column in the Customers Table with the Customer 
ID in the Sales transactions table (or any relevant data). 
• Create Calculated Columns in Power Pivot: 
• Once the tables are connected, create calculated columns in Power Pivot to perform 
complex calculations, such as Profit Margin (calculated as Profit/Sales). 
• Create a Power Pivot-based PivotTable to summarize sales performance using your 
data model. 
13. Insights and Recommendations: 
Summarize the key insights from your analysis, such as: 
• Which Sales Channels (e.g., Ship Mode) are driving the highest Sales. 
• Which Product Categories are performing best and worst based on Adjusted 
Sales and Quantity. 
• The effect of Discounts on Net Sales and Adjusted Sales. 
• Any actionable recommendations for improving Sales Performance, reducing 
Returns (if applicable), or optimizing Discount strategies. 
