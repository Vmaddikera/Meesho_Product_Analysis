"""
Return Percentage Analysis - Based on Excel Results
Plots percentage of returns for each category and price range
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Meesho Brand Colors
JAMUNI = '#580b48'  
AAM = '#FFA500'     

def load_excel_data():
    """Load data from Excel file"""
    print("Loading data from Excel file...")
    
    # Read category analysis
    category_data = pd.read_excel('meesho_analysis_results.xlsx', sheet_name='Category_Analysis', index_col=0)
    print("Category Analysis Data:")
    print(category_data)
    
    # Read price range analysis
    price_data = pd.read_excel('meesho_analysis_results.xlsx', sheet_name='Price_Range_Analysis', index_col=0)
    print("\nPrice Range Analysis Data:")
    print(price_data)
    
    # Calculate percentage of total returns for each price range
    total_returns = price_data['returns'].sum()
    price_data['percentage_of_total_returns'] = (price_data['returns'] / total_returns) * 100
    
    print("\nPrice Range Data with Return Percentages:")
    print(price_data)
    
    return category_data, price_data

def create_return_percentage_plots(category_data, price_data):
    """Create plots showing percentage of returns for each category and price range"""
    print("Creating return percentage plots...")
    
    # Create the plots
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
    fig.suptitle('Meesho Return Analysis - Percentage of Total Returns', 
                 fontsize=20, fontweight='bold', color=JAMUNI, y=0.98)
    
    # Plot 1: Percentage of Total Returns by Category
    categories = category_data.index
    return_percentages = category_data['percentage_of_total_returns'].values
    
    # Use alternating Meesho brand colors
    colors = [JAMUNI, AAM, JAMUNI, AAM, JAMUNI, AAM][:len(categories)]
    
    bars1 = ax1.bar(categories, return_percentages, 
                   color=colors, edgecolor='white', linewidth=2, alpha=0.9)
    
    ax1.set_title('Percentage of Total Returns by Category', 
                 fontsize=16, fontweight='bold', 
                 color=JAMUNI, pad=20)
    ax1.set_xlabel('Product Categories', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Percentage of Total Returns (%)', fontsize=14, fontweight='bold')
    ax1.tick_params(axis='x', rotation=45, labelsize=12)
    ax1.tick_params(axis='y', labelsize=12)
    ax1.grid(False)
    ax1.set_ylim(0, max(return_percentages) * 1.2)
    
    # Add value labels and return counts
    for i, (bar, percentage, returns, total_returns) in enumerate(zip(bars1, return_percentages, 
                                                                   category_data['returns'], 
                                                                   category_data['total_orders'])):
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                f'{percentage:.1f}%\n({returns} returns)', ha='center', va='bottom', 
                fontweight='bold', fontsize=10)
    
    # Plot 2: Percentage of Total Returns by Price Range
    price_ranges = price_data.index
    price_return_percentages = price_data['percentage_of_total_returns'].values
    
    bars2 = ax2.bar(price_ranges, price_return_percentages,
                   color=AAM, edgecolor='white', linewidth=2, alpha=0.9)
    
    ax2.set_title('Percentage of Total Returns by Price Range', 
                 fontsize=16, fontweight='bold', 
                 color=JAMUNI, pad=20)
    ax2.set_xlabel('Price Ranges (₹)', fontsize=14, fontweight='bold')
    ax2.set_ylabel('Percentage of Total Returns (%)', fontsize=14, fontweight='bold')
    ax2.tick_params(axis='x', rotation=45, labelsize=12)
    ax2.tick_params(axis='y', labelsize=12)
    ax2.grid(False)
    ax2.set_ylim(0, max(price_return_percentages) * 1.2)
    
    # Add value labels and return counts
    for bar, percentage, returns in zip(bars2, price_return_percentages, price_data['returns']):
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                f'{percentage:.1f}%\n({returns} returns)', ha='center', va='bottom', 
                fontweight='bold', fontsize=10)
    
    plt.tight_layout()
    plt.savefig('return_percentage_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return fig

def print_return_analysis(category_data, price_data):
    """Print detailed return analysis"""
    print("\n" + "="*60)
    print("RETURN PERCENTAGE ANALYSIS")
    print("="*60)
    
    print("\nReturn Distribution by Category:")
    for category, data in category_data.iterrows():
        print(f"  {category}:")
        print(f"    - Percentage of total returns: {data['percentage_of_total_returns']:.2f}%")
        print(f"    - Number of returns: {data['returns']}")
        print(f"    - Total orders in category: {data['total_orders']}")
        print(f"    - Return rate within category: {data['return_rate_within_category']:.2f}%")
        print()
    
    print("\nReturn Distribution by Price Range:")
    for price_range, data in price_data.iterrows():
        print(f"  {price_range}:")
        print(f"    - Percentage of total returns: {data['percentage_of_total_returns']:.2f}%")
        print(f"    - Number of returns: {data['returns']}")
        print(f"    - Total orders in range: {data['total_orders']}")
        print(f"    - Return rate within range: {data['return_rate']:.2f}%")
        print()

def main():
    """Main function to run the return percentage analysis"""
    print("Starting Return Percentage Analysis...")
    print("=" * 50)
    
    # Load data from Excel
    category_data, price_data = load_excel_data()
    
    # Create plots
    fig = create_return_percentage_plots(category_data, price_data)
    
    # Print analysis
    print_return_analysis(category_data, price_data)
    
    print("\n" + "="*60)
    print("="*60)
    print(" Plotted percentage of returns for each category")
    print(" Plotted return rates for each price range")
    print(" Used Meesho brand colors (Purple & Yellow/Orange)")
    print(" Generated: return_percentage_analysis.png")
    print(" No grid lines - clean professional look")

if __name__ == "__main__":
    main()
