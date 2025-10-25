# Meesho Sales Data Analysis

Real data analysis of Meesho sales with NLP product categorization and return rate analysis.

##  Quick Start

```bash
# Activate environment
venv\Scripts\activate

# Run main analysis
python real_data_analysis.py

# Run return percentage analysis
python return_percentage_plots.py

# Run stacked bar analysis
python stacked_bar_analysis.py
```

##  How to Run

### **Prerequisites**
- Python 3.8+ installed
- Virtual environment activated

### **Step 1: Setup Environment**
```bash
# Navigate to project directory
cd d:\meesho

# Activate virtual environment
venv\Scripts\activate

# Install dependencies (if not already installed)
pip install -r requirements.txt
```

### **Step 2: Run Analysis Scripts**

#### **Main Analysis (Complete)**
```bash
python real_data_analysis.py
```
**Output:**
- `meesho_analysis_results.xlsx` - Excel report with 4 sheets
- `real_meesho_data_analysis.png` - Main dashboard
- Console output with detailed statistics

#### **Return Percentage Analysis**
```bash
python return_percentage_plots.py
```
**Output:**
- `return_percentage_analysis.png` - Return percentage charts
- Console output with return analysis

#### **Stacked Bar Analysis**
```bash
python stacked_bar_analysis.py
```
**Output:**
- `stacked_bar_analysis.png` - Stacked bar charts
- `return_percentage_analysis.png` - Percentage analysis
- Console output with detailed breakdown

### **Step 3: View Results**
- **Excel File:** Open `meesho_analysis_results.xlsx` for comprehensive data
- **Charts:** View PNG files for visualizations
- **Console:** Check terminal for detailed statistics and insights

### **Expected Runtime**
- Main analysis: ~30-60 seconds
- Return percentage: ~10-20 seconds  
- Stacked bar analysis: ~15-30 seconds

##  Files

### **Core Scripts**
- `real_data_analysis.py` - Main analysis with NLP categorization
- `return_percentage_plots.py` - Return percentage charts
- `stacked_bar_analysis.py` - Stacked bar charts (orders vs returns)

### **Data**
- `meesho ForwardReports.csv` - Forward reports (138 records)
- `meesho Orders Aug.csv` - Orders data (208 records)
- `meesho_analysis_results.xlsx` - Excel output (4 sheets)

### **Output**
- `real_meesho_data_analysis.png` - Main dashboard
- `return_percentage_analysis.png` - Return percentage charts
- `stacked_bar_analysis.png` - Stacked bar charts

##  Analysis Features

### **Data Processing**
- Merges 2 CSV files → 133 records
- NLP product categorization (6 categories)
- Return rate calculations

### **Visualizations**
- **Stacked Bars:** Total orders vs returns
- **Return Percentages:** % of total returns by category/price range
- **Meesho Branding:** Purple (#580b48) & Yellow (#FFA500)

### **Excel Output**
- Category Analysis
- Price Range Analysis  
- Summary Statistics
- Product Categories (all products with categories)

##  Key Insights

- Return rates by product category
- Return patterns by price range
- Percentage of total returns
- Product categorization accuracy

##  Requirements

```
pandas
matplotlib
seaborn
plotly
openpyxl
```

##  Technical Workflow

1. **Load & Merge Data** - CSV files → 133 merged records
2. **NLP Categorization** - 6 categories using keyword matching
3. **Statistical Analysis** - Return rates and percentages
4. **Visualization** - Stacked bars and percentage charts
5. **Excel Export** - Comprehensive 4-sheet report

##  Output Files

- **Excel:** `meesho_analysis_results.xlsx` (4 sheets)
- **Charts:** PNG files with Meesho branding
- **Console:** Detailed statistics and insights
