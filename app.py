import pandas as pd
import plotly.io as pio

from dash_bootstrap_templates import load_figure_template
from dash import Dash, dcc, html
import dash_bootstrap_components as dbc
import numpy as np # linear algebra
import pandas as pd # data processing, CSV file I/O (e.g. pd.read_csv)
import plotly.express as px
import seaborn as sns
import plotly.graph_objects as go
from scipy.stats import pearsonr
sales = pd.read_excel('AdventureWorks.xlsx', sheet_name = 'FactInternetSales')
product = pd.read_excel('AdventureWorks.xlsx', sheet_name = 'DimProduct')
category = pd.read_excel('AdventureWorks.xlsx', sheet_name = 'DimProductCategory')
sub_category = pd.read_excel('AdventureWorks.xlsx', sheet_name = 'DimProductSubcategory')
#JOIN PRODUCT CATEGORY
product = product.merge(sub_category, on='ProductSubcategoryKey')
product = product.merge(category, left_on = 'ProductCategoryKey', right_on = 'ProductCategoryKey')
sales = sales.merge(product, on='ProductKey')

total_sales = sales.SalesAmount.sum()
# KPI Card for Total Sales
fig_kpi = go.Figure(go.Indicator(
    mode="number",
    value=total_sales,
    title={"text": "Total Sales"},
))

# Bar Chart for Total Sales
fig_bar = go.Figure(data=[go.Bar(x=['Total Sales'], y=[total_sales], marker_color='blue')])
fig_bar.update_layout(title='Total Sales', yaxis_title='Sales Amount')

# Pie Chart for Sales by Product Category
category_sales = sales.groupby('ProductCategory')['SalesAmount'].sum().reset_index()
fig_pie = px.pie(category_sales, names='ProductCategory', values='SalesAmount', title='Sales by Product Category')

# Convert OrderDate to datetime
sales['OrderDate'] = pd.to_datetime(sales['OrderDate'])

# Identify the major category and the other two smaller categories
major_category = category_sales.loc[category_sales['SalesAmount'].idxmax(), 'ProductCategory']
smaller_categories = category_sales[category_sales['ProductCategory'] != major_category]['ProductCategory'].tolist()

# Monthly Sales Trends with Category Breakdown
sales['YearMonth'] = sales['OrderDate'].dt.to_period('M').astype(str)
monthly_sales = sales.groupby(['YearMonth', 'ProductCategory'])['SalesAmount'].sum().reset_index()

# Separate data for the major category and smaller categories
monthly_sales_major = monthly_sales[monthly_sales['ProductCategory'] == major_category]
monthly_sales_smaller = monthly_sales[monthly_sales['ProductCategory'].isin(smaller_categories)]

# Line Chart for Monthly Sales Trends - Major Category
fig_line_monthly_major = px.line(monthly_sales_major, x='YearMonth', y='SalesAmount', color='ProductCategory', title=f'Monthly Sales Trends - {major_category}')
fig_line_monthly_sales = px.line(monthly_sales, x='YearMonth', y='SalesAmount')
# Line Chart for Monthly Sales Trends - Smaller Categories
fig_line_monthly_smaller = px.line(monthly_sales_smaller, x='YearMonth', y='SalesAmount', color='ProductCategory', title='Monthly Sales Trends - Smaller Categories')

# Yearly Sales Trends with Category Breakdown
sales['Year'] = sales['OrderDate'].dt.year
yearly_sales = sales.groupby(['Year', 'ProductCategory'])['SalesAmount'].sum().reset_index()

# Separate data for the major category and smaller categories
yearly_sales_major = yearly_sales[yearly_sales['ProductCategory'] == major_category]
yearly_sales_smaller = yearly_sales[yearly_sales['ProductCategory'].isin(smaller_categories)]

# Line Chart for Yearly Sales Trends - Major Category
fig_line_yearly_major = px.line(yearly_sales_major, x='Year', y='SalesAmount', color='ProductCategory', title=f'Yearly Sales Trends - {major_category}')

# Line Chart for Yearly Sales Trends - Smaller Categories
fig_line_yearly_smaller = px.line(yearly_sales_smaller, x='Year', y='SalesAmount', color='ProductCategory', title='Yearly Sales Trends - Smaller Categories')
fig_line_yearly_sales = px.line(yearly_sales, x='Year', y='SalesAmount')


# Line Chart for Sales Trends Over Time
sales['OrderDate'] = pd.to_datetime(sales['OrderDate'])
sales['YearMonth'] = sales['OrderDate'].dt.to_period('M').astype(str)  # Convert Period to string
monthly_sales = sales.groupby('YearMonth')['SalesAmount'].sum().reset_index()
fig_line_monthly = px.line(monthly_sales, x='YearMonth', y='SalesAmount', title='Monthly Sales Trends')

sales['Year'] = sales['OrderDate'].dt.year
yearly_sales = sales.groupby('Year')['SalesAmount'].sum().reset_index()
fig_line_yearly = px.line(yearly_sales, x='Year', y='SalesAmount', title='Yearly Sales Trends')


# Q2
customer = pd.read_excel('AdventureWorks.xlsx', sheet_name = 'DimCustomer')
sales = sales.merge(customer, on='CustomerKey')
sales['Age'] = sales['BirthDate'].apply(lambda x: pd.Timestamp.now().year - pd.to_datetime(x).year)
geography = pd.read_excel('AdventureWorks.xlsx', sheet_name = 'DimGeography')
sales = sales.merge(geography, on='GeographyKey')
# Analyze customer demographics
# Example: Age distribution
age_distribution = px.histogram(sales, x='Age', title='Customer Age Distribution')

# Example: Gender distribution
gender_distribution = px.pie(sales, names='Gender', title='Customer Gender Distribution')
# Example: Geographic distribution
geographic_distribution = px.histogram(sales, x='EnglishCountryRegionName', title='Customer Geographic Distribution')
#geographic_distribution = px.scatter_geo(sales, locations='EnglishCountryRegionName', color='EnglishCountryRegionName', title='Customer Geographic Distribution')
# Correlate demographics with sales data (example correlation plot)
age_sales_correlation = px.scatter(sales, x='Age', y='SalesAmount', trendline='ols',
                               title='Correlation between Customer Age and Sales Amount by Gender')


# NO GENDER CORRELATION
income_sales_correlation = px.scatter(sales, trendline='ols', trendline_color_override="red", x='YearlyIncome', y='SalesAmount', title="Correlation between Income and Sales Amount")
# Group by country and calculate total sales amount
country_sales = sales.groupby('EnglishCountryRegionName')['SalesAmount'].sum().reset_index()
fig_country_sales = px.bar(country_sales, x='EnglishCountryRegionName', y='SalesAmount', title='Total Sales Amount by Country')

# Q3
# Calculate total sales amount by product
product_sales = sales.groupby('ProductName')['SalesAmount'].sum().reset_index()

# Identify top-performing products by sales
top_products = product_sales.sort_values(by='SalesAmount', ascending=False).head(10)

# Visualize top-performing products
fig_top_products = px.bar(top_products, x='ProductName', y='SalesAmount',
                          title='Top-Performing Products by Sales',
                          labels={'SalesAmount': 'Total Sales Amount', 'ProductName': 'Product Name'})
fig_top_products.show()

# Calculate total sales amount by product
product_sales = sales.groupby('ProductName')['SalesAmount'].sum().reset_index()

# Identify top-performing products by sales
top_products = product_sales.sort_values(by='SalesAmount', ascending=False).head(10)

# Visualize top-performing products
fig_top_products = px.bar(top_products, x='ProductName', y='SalesAmount',
                          title='Top-Performing Products by Sales',
                          labels={'SalesAmount': 'Total Sales Amount', 'ProductName': 'Product Name'})

monthly_sales_top_product = sales[sales['ProductName'] == top_products['ProductName'].tolist()[0]].groupby((['YearMonth']))['SalesAmount'].sum().reset_index()
fig_line_monthly_top_product = px.line(monthly_sales, x='YearMonth', y='SalesAmount', title='Monthly Sales Trend of BestSelling Product')

load_figure_template('FLATLY')
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])

# Define the layout of the app
app.layout = html.Div([
    html.H1("Sales Performance Dashboard"),
    html.Div([
        dcc.Graph(figure=fig_kpi)
    ], style={'width': '15%', 'height' : '20%', 'display': 'inline-block'}),

    
    html.Div([
        dcc.Graph(figure=fig_pie)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
    html.Div([
        dcc.Graph(figure=fig_line_monthly_major)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
    html.Div([
        dcc.Graph(figure=fig_line_monthly_smaller)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
    html.Div([
        dcc.Graph(figure=fig_line_yearly_major)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
    html.Div([
        dcc.Graph(figure=fig_line_yearly_smaller)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
   html.Div([
        dcc.Graph(figure=age_distribution)
    ], style={'width': '30%', 'display': 'inline-block'}),
    

    html.Div([
        dcc.Graph(figure=income_sales_correlation)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
    html.Div([
        dcc.Graph(figure=geographic_distribution)
    ], style={'width': '30%', 'display': 'inline-block'}),
    
    
    
    html.Div([
        dcc.Graph(figure=fig_country_sales)
    ], style={'width': '30%', 'display': 'inline-block'}), 
    html.Div([
        dcc.Graph(figure=fig_top_products)
    ], style={'width': '30%', 'display': 'inline-block'}), 
    html.Div([
        dcc.Graph(figure=fig_line_monthly_top_product)
    ], style={'width': '30%', 'display': 'inline-block'}), 
])

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)