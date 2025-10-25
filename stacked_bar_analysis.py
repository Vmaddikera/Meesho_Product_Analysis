"""
Stacked Bar Chart Analysis for Meesho Data
Shows total orders vs returns with return percentages
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Meesho Brand Colors
JAMUNI = '#580b48'  # Deep Purple
AAM = '#FFA500'     # Bright Yellow/Orange

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

def create_stacked_bar_charts(category_data, price_data):
    """Create stacked bar charts showing total orders vs returns"""
    print("Creating stacked bar charts...")
    
    # Create the plots
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
    fig.suptitle('Meesho Orders vs Returns Analysis - Stacked Bar Charts', 
                 fontsize=20, fontweight='bold', color=JAMUNI, y=0.98)
    
    # Plot 1: Stacked Bar Chart by Category
    categories = category_data.index
    total_orders = category_data['total_orders'].values
    returns = category_data['returns'].values
    non_returns = total_orders - returns
    
    # Create stacked bars
    bars1 = ax1.bar(categories, non_returns, 
                   color=JAMUNI, alpha=0.7, label='Delivered Orders', edgecolor='white', linewidth=1)
    bars2 = ax1.bar(categories, returns, bottom=non_returns,
                   color=AAM, alpha=0.9, label='Returned Orders', edgecolor='white', linewidth=1)
    
    ax1.set_title('Orders vs Returns by Category', 
                 fontsize=16, fontweight='bold', 
                 color=JAMUNI, pad=20)
    ax1.set_xlabel('Product Categories', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Number of Orders', fontsize=14, fontweight='bold')
    ax1.tick_params(axis='x', rotation=45, labelsize=12)
    ax1.tick_params(axis='y', labelsize=12)
    ax1.grid(False)
    ax1.legend(fontsize=12, loc='upper right')
    
    # Add value labels on top of each stack
    for i, (category, total, ret, non_ret) in enumerate(zip(categories, total_orders, returns, non_returns)):
        # Label for total orders
        ax1.text(i, total + 0.5, f'Total: {total}', ha='center', va='bottom', 
                fontweight='bold', fontsize=10, color=JAMUNI)
        # Label for returns
        if ret > 0:
            ax1.text(i, non_ret + ret/2, f'Returns: {ret}', ha='center', va='center', 
                    fontweight='bold', fontsize=10, color='white')
    
    # Plot 2: Stacked Bar Chart by Price Range
    price_ranges = price_data.index
    total_orders_price = price_data['total_orders'].values
    returns_price = price_data['returns'].values
    non_returns_price = total_orders_price - returns_price
    
    # Create stacked bars
    bars3 = ax2.bar(price_ranges, non_returns_price, 
                   color=JAMUNI, alpha=0.7, label='Delivered Orders', edgecolor='white', linewidth=1)
    bars4 = ax2.bar(price_ranges, returns_price, bottom=non_returns_price,
                   color=AAM, alpha=0.9, label='Returned Orders', edgecolor='white', linewidth=1)
    
    ax2.set_title('Orders vs Returns by Price Range', 
                 fontsize=16, fontweight='bold', 
                 color=JAMUNI, pad=20)
    ax2.set_xlabel('Price Ranges (₹)', fontsize=14, fontweight='bold')
    ax2.set_ylabel('Number of Orders', fontsize=14, fontweight='bold')
    ax2.tick_params(axis='x', rotation=45, labelsize=12)
    ax2.tick_params(axis='y', labelsize=12)
    ax2.grid(False)
    ax2.legend(fontsize=12, loc='upper right')
    
    # Add value labels on top of each stack
    for i, (price_range, total, ret, non_ret) in enumerate(zip(price_ranges, total_orders_price, returns_price, non_returns_price)):
        # Label for total orders
        ax2.text(i, total + 0.5, f'Total: {total}', ha='center', va='bottom', 
                fontweight='bold', fontsize=10, color=JAMUNI)
        # Label for returns
        if ret > 0:
            ax2.text(i, non_ret + ret/2, f'Returns: {ret}', ha='center', va='center', 
                    fontweight='bold', fontsize=10, color='white')
    
    plt.tight_layout()
    plt.savefig('stacked_bar_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return fig

def create_percentage_analysis(category_data, price_data):
    """Create percentage analysis showing return percentages out of total returns"""
    print("Creating percentage analysis...")
    
    # Create the plots
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
    fig.suptitle('Return Percentage Analysis - Percentage of Total Returns', 
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
    
    # Add value labels
    for i, (bar, percentage, returns) in enumerate(zip(bars1, return_percentages, category_data['returns'])):
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
    
    # Add value labels
    for bar, percentage, returns in zip(bars2, price_return_percentages, price_data['returns']):
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                f'{percentage:.1f}%\n({returns} returns)', ha='center', va='bottom', 
                fontweight='bold', fontsize=10)
    
    plt.tight_layout()
    plt.savefig('return_percentage_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    return fig

def print_analysis_summary(category_data, price_data):
    """Print detailed analysis summary"""
    print("\n" + "="*70)
    print("STACKED BAR CHART ANALYSIS SUMMARY")
    print("="*70)
    
    print("\n📊 STACKED BAR CHARTS:")
    print("   - Bottom Stack (Purple): Delivered Orders")
    print("   - Top Stack (Yellow/Orange): Returned Orders")
    print("   - Shows total orders vs returns for each category/price range")
    
    print("\n📈 PERCENTAGE ANALYSIS:")
    print("   - Shows percentage of total returns for each category/price range")
    print("   - Percentages calculated as: (Category Returns / Total Returns) × 100")
    
    print("\n📋 CATEGORY ANALYSIS:")
    for category, data in category_data.iterrows():
        print(f"   {category}:")
        print(f"     - Total Orders: {data['total_orders']}")
        print(f"     - Returns: {data['returns']}")
        print(f"     - Delivered: {data['total_orders'] - data['returns']}")
        print(f"     - Return Rate: {data['return_rate_within_category']:.2f}%")
        print(f"     - % of Total Returns: {data['percentage_of_total_returns']:.2f}%")
        print()
    
    print("\n💰 PRICE RANGE ANALYSIS:")
    for price_range, data in price_data.iterrows():
        print(f"   {price_range}:")
        print(f"     - Total Orders: {data['total_orders']}")
        print(f"     - Returns: {data['returns']}")
        print(f"     - Delivered: {data['total_orders'] - data['returns']}")
        print(f"     - Return Rate: {data['return_rate']:.2f}%")
        print(f"     - % of Total Returns: {data['percentage_of_total_returns']:.2f}%")
        print()

def main():
    """Main function to run the stacked bar analysis"""
    print("Starting Stacked Bar Chart Analysis...")
    print("=" * 50)
    
    # Load data from Excel
    category_data, price_data = load_excel_data()
    
    # Create stacked bar charts
    fig1 = create_stacked_bar_charts(category_data, price_data)
    
    # Create percentage analysis
    fig2 = create_percentage_analysis(category_data, price_data)
    
    # Print analysis summary
    print_analysis_summary(category_data, price_data)
    
    print("\n" + "="*70)
    print("STACKED BAR ANALYSIS COMPLETED!")
    print("="*70)
    print("✓ Created stacked bar charts showing orders vs returns")
    print("✓ Created percentage analysis of total returns")
    print("✓ Used Meesho brand colors (Purple & Yellow/Orange)")
    print("✓ Generated: stacked_bar_analysis.png")
    print("✓ Generated: return_percentage_analysis.png")
    print("✓ No grid lines - clean professional look")

if __name__ == "__main__":
    main()
