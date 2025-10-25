# Meesho Sales Data Analysis

A comprehensive data analysis framework for analyzing Meesho sales data, return rates, and product performance using real CSV data with NLP-based product categorization and professional visualizations.

##  Overview

This project analyzes Meesho sales data to understand return patterns, product performance, and business insights using advanced NLP techniques and data visualization.

##  Project Structure

```
d:\meesho\
├── real_data_analysis.py              
├── return_percentage_plots.py         
├── meesho ForwardReports.csv          
├── meesho Orders Aug.csv              
├── meesho_analysis_results.xlsx      
├── real_meesho_data_analysis.png    
├── return_percentage_analysis.png     
├── requirements.txt                  
└── venv/                            
```

##  Quick Start

### 1. Setup Environment
```bash
# Activate virtual environment
cd d:\meesho
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Run Complete Analysis
```bash
# Run main analysis (generates Excel + PNG)
python real_data_analysis.py

# Run return percentage analysis (generates return-focused charts)
python return_percentage_plots.py
```

##  Analysis Features

### **Data Processing**
- **Data Merging:** Combines ForwardReports.csv (138 records) + Orders Aug.csv (208 records) → 133 merged records
- **Data Cleaning:** Handles missing values, converts data types, standardizes formats
- **Return Flagging:** Identifies returns using order status ('Return', 'rto')

### **NLP Product Categorization**
- **6 Categories:** Ethnic Wear, Western Wear, Beauty & Grooming, Accessories, Home & Living, Electronics
- **Keyword Matching:** 100+ keywords across categories for accurate classification
- **Text Preprocessing:** Lowercase conversion, special character removal, stop word filtering
- **Smart Classification:** Products assigned to categories with highest keyword match scores

### **Statistical Analysis**
- **Overall Return Rate:** Total returns / Total orders
- **Category Analysis:** Return rates and percentages by product category
- **Price Range Analysis:** Return patterns across different price segments
- **Percentage Calculations:** Both within-category and total-order percentages

##  Output Files

### **Excel Report (`meesho_analysis_results.xlsx`)**
-  **Category_Analysis** - Return rates and percentages by category
- **Price_Range_Analysis** - Return analysis by price ranges
- **Summary** - Overall statistics and metrics
- **Product_Categories** - All products with their assigned categories

### **Visualizations**
- **`real_meesho_data_analysis.png`** - Main dashboard with category and price range analysis
- **`return_percentage_analysis.png`** - Return-focused charts showing percentage of total returns

##  Data Requirements

### **Input Files**
1. **`meesho ForwardReports.csv`** 
   - Columns: order_date, sub_order_num, order_status, state, pin, gst_amount, meesho_price, shipping_charges_total, price

2. **`meesho Orders Aug.csv`**
   - Columns: "Reason for Credit Entry", "Sub Order No", "Order Date", "Customer State", "Product Name", "SKU", "Size", "Quantity", "Supplier Listed Price", "Supplier Discounted Price"

### **Output Files**
- **Excel Analysis:** Comprehensive data breakdown 
- **PNG Visualizations:** Charts with Meesho branding

## 🔧 Technical Workflow

### **Step 1: Data Loading & Merging**
- Load CSV files and merge on `sub_order_num` and `Sub Order No`
- Result: 133 merged records with 19 columns

### **Step 2: Data Preprocessing**
- Convert dates, clean product names, create return flags
- Handle missing values and standardize data formats

### **Step 3: NLP Product Categorization**
- Apply keyword matching across 6 product categories
- Use 100+ keywords for accurate classification
- Assign products to categories with highest match scores

### **Step 4: Statistical Analysis**
- Calculate return rates, percentages, and distributions
- Generate category-wise and price range analysis
- Compute both within-category and total-order metrics

### **Step 5: Visualization & Export**
- Create professional charts with Meesho brand colors
- Export comprehensive Excel report with 4 sheets
- Generate clean PNG visualizations without grid lines

##  Dependencies

```
pandas >= 1.5.0
numpy >= 1.21.0
matplotlib >= 3.5.0
seaborn >= 0.11.0
plotly >= 5.0.0
openpyxl >= 3.0.0
```

##  Usage Examples

### **Run Complete Analysis**
```bash
python real_data_analysis.py
```
Generates: Excel report + main visualization

### **Run Return Percentage Analysis**
```bash
python return_percentage_plots.py
```
