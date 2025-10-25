"""
Real Meesho Data Analysis with NLP Product Categorization
Using actual CSV data to calculate real return rates
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import re
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# Meesho Brand Colors
JAMUNI = '#580b48'  # Deep Purple
AAM = '#FFA500'     # Bright Yellow/Orange

class RealMeeshoAnalysis:
    def __init__(self):
        # Simple stop words list
        self.stop_words = {'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'are', 'was', 'were', 'be', 'been', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may', 'might', 'must', 'can', 'this', 'that', 'these', 'those', 'i', 'you', 'he', 'she', 'it', 'we', 'they', 'me', 'him', 'her', 'us', 'them'}
        
    def load_real_data(self):
        """Load the actual CSV files"""
        print("Loading real Meesho data...")
        
        # Load forward reports
        self.forward_reports = pd.read_csv('meesho ForwardReports.csv')
        print(f"Forward Reports: {self.forward_reports.shape}")
        
        # Load orders data
        self.orders_data = pd.read_csv('meesho Orders Aug.csv')
        print(f"Orders Data: {self.orders_data.shape}")
        
        # Merge datasets
        self.merged_data = pd.merge(
            self.forward_reports, 
            self.orders_data, 
            left_on='sub_order_num', 
            right_on='Sub Order No', 
            how='inner'
        )
        print(f"Merged Data: {self.merged_data.shape}")
        
        return self.merged_data
    
    def preprocess_data(self):
        """Clean and preprocess the data"""
        print("Preprocessing data...")
        
        # Convert date columns
        self.merged_data['order_date'] = pd.to_datetime(self.merged_data['order_date'])
        
        # Clean product names
        self.merged_data['Product Name'] = self.merged_data['Product Name'].astype(str)
        self.merged_data['Product Name'] = self.merged_data['Product Name'].str.strip()
        
        # Create return flag
        self.merged_data['is_return'] = self.merged_data['order_status'].isin(['Return', 'rto'])
        self.merged_data['is_delivered'] = self.merged_data['order_status'] == 'Delivered'
        self.merged_data['is_cancelled'] = self.merged_data['order_status'] == 'Cancelled'
        
        # Clean price data
        self.merged_data['meesho_price_clean'] = pd.to_numeric(self.merged_data['meesho_price'], errors='coerce')
        
        print("Data preprocessing completed!")
        return self.merged_data
    
    def extract_product_features(self, text):
        """Extract features from product names using simple tokenization"""
        if pd.isna(text):
            return []
        
        # Convert to lowercase and remove special characters
        text = re.sub(r'[^a-zA-Z\s]', '', str(text).lower())
        
        # Simple tokenization without NLTK
        tokens = text.split()
        tokens = [token for token in tokens 
                 if token not in self.stop_words and len(token) > 2]
        
        return tokens
    
    def categorize_products_nlp(self):
        """Categorize products using advanced NLP analysis"""
        print("Performing NLP analysis for product categorization...")
        
        # Define comprehensive product categories with keywords
        categories = {
            'Ethnic Wear': [
                'saree', 'lehenga', 'choli', 'kurta', 'dupatta', 'suit', 'ethnic', 'traditional',
                'salwar', 'kameez', 'anarkali', 'ghagra', 'churidar', 'palazzo', 'sharara',
                'embroidered', 'sequence', 'work', 'designer', 'party', 'wear', 'heavy',
                'georgette', 'crepe', 'net', 'organza', 'velvet', 'silk', 'cotton'
            ],
            'Western Wear': [
                'gown', 'dress', 'top', 'bottom', 'shirt', 'western', 'casual', 'jeans',
                'trouser', 'pant', 'skirt', 'blouse', 'tank', 'crop', 'cami', 'tunic',
                'maxi', 'mini', 'midi', 'bodycon', 'a-line', 'wrap', 'shift'
            ],
            'Beauty & Grooming': [
                'hair', 'straightener', 'beauty', 'grooming', 'cosmetic', 'makeup',
                'skincare', 'haircare', 'styling', 'tools', 'brush', 'comb', 'mirror'
            ],
            'Accessories': [
                'belt', 'jewelry', 'bag', 'accessory', 'jewellery', 'necklace', 'earring',
                'bracelet', 'ring', 'watch', 'scarf', 'shawl', 'handbag', 'purse'
            ],
            'Home & Living': [
                'home', 'living', 'decor', 'furniture', 'kitchen', 'bedding', 'curtain',
                'cushion', 'pillow', 'blanket', 'towel', 'carpet', 'rug'
            ],
            'Electronics': [
                'electronic', 'gadget', 'device', 'machine', 'phone', 'mobile', 'charger',
                'cable', 'headphone', 'speaker', 'camera', 'laptop', 'tablet'
            ]
        }
        
        def assign_category_nlp(product_name):
            """Advanced NLP-based category assignment"""
            product_lower = str(product_name).lower()
            
            # Extract features
            features = self.extract_product_features(product_name)
            
            # Score each category
            category_scores = {}
            for category, keywords in categories.items():
                score = 0
                for keyword in keywords:
                    if keyword in product_lower:
                        score += 1
                    # Also check in extracted features
                    if keyword in features:
                        score += 0.5
                category_scores[category] = score
            
            # Return category with highest score
            if max(category_scores.values()) > 0:
                return max(category_scores, key=category_scores.get)
            else:
                return 'Other'
        
        # Apply NLP categorization
        self.merged_data['product_category'] = self.merged_data['Product Name'].apply(assign_category_nlp)
        
        # Print category distribution
        category_counts = self.merged_data['product_category'].value_counts()
        print("\nProduct Category Distribution:")
        for category, count in category_counts.items():
            print(f"  {category}: {count} products ({count/len(self.merged_data)*100:.1f}%)")
        
        return self.merged_data
    
    def calculate_real_return_rates(self):
        """Calculate real return rates from actual data"""
        print("Calculating real return rates...")
        
        # Overall return rate
        total_orders = len(self.merged_data)
        total_returns = self.merged_data['is_return'].sum()
        overall_return_rate = (total_returns / total_orders) * 100
        
        print(f"\nOverall Statistics:")
        print(f"  Total Orders: {total_orders:,}")
        print(f"  Total Returns: {total_returns:,}")
        print(f"  Overall Return Rate: {overall_return_rate:.2f}%")
        
        # Return rate by category - percentages relative to TOTAL orders
        category_returns = self.merged_data.groupby('product_category').agg({
            'is_return': ['sum', 'count'],
            'meesho_price_clean': 'mean'
        }).round(2)
        
        category_returns.columns = ['returns', 'total_orders', 'avg_price']
        # Calculate return rate within category
        category_returns['return_rate_within_category'] = (category_returns['returns'] / category_returns['total_orders']) * 100
        # Calculate percentage of total orders
        category_returns['percentage_of_total_orders'] = (category_returns['total_orders'] / total_orders) * 100
        # Calculate percentage of total returns
        category_returns['percentage_of_total_returns'] = (category_returns['returns'] / total_returns) * 100
        category_returns = category_returns.sort_values('return_rate_within_category', ascending=False)
        
        print(f"\nReturn Rates by Category:")
        for category, data in category_returns.iterrows():
            print(f"  {category}:")
            print(f"    - Return Rate within category: {data['return_rate_within_category']:.2f}%")
            print(f"    - Percentage of total orders: {data['percentage_of_total_orders']:.2f}%")
            print(f"    - Percentage of total returns: {data['percentage_of_total_returns']:.2f}%")
            print(f"    - Orders: {data['total_orders']}, Returns: {data['returns']}")
            print()
        
        # Return rate by price range
        self.merged_data['price_range'] = pd.cut(
            self.merged_data['meesho_price_clean'], 
            bins=[0, 500, 1000, 1500, 2000, float('inf')], 
            labels=['0-500', '500-1000', '1000-1500', '1500-2000', '2000+']
        )
        
        price_returns = self.merged_data.groupby('price_range').agg({
            'is_return': ['sum', 'count'],
            'meesho_price_clean': 'mean'
        }).round(2)
        
        price_returns.columns = ['returns', 'total_orders', 'avg_price']
        price_returns['return_rate'] = (price_returns['returns'] / price_returns['total_orders']) * 100
        
        print(f"\nReturn Rates by Price Range:")
        for price_range, data in price_returns.iterrows():
            print(f"  {price_range}: {data['return_rate']:.2f}% ({data['returns']} returns out of {data['total_orders']} orders)")
        
        return category_returns, price_returns, overall_return_rate
    
    def save_data_to_excel(self, category_returns, price_returns, overall_rate):
        """Save the analysis data to Excel file"""
        print("Saving data to Excel...")
        
        with pd.ExcelWriter('meesho_analysis_results.xlsx', engine='openpyxl') as writer:
            # Save category analysis
            category_returns.to_excel(writer, sheet_name='Category_Analysis', index=True)
            
            # Save price range analysis
            price_returns.to_excel(writer, sheet_name='Price_Range_Analysis', index=True)
            
            # Save overall summary
            summary_data = {
                'Metric': ['Total Orders', 'Total Returns', 'Overall Return Rate (%)'],
                'Value': [len(self.merged_data), self.merged_data['is_return'].sum(), overall_rate]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Save product categories data
            product_categories_data = self.merged_data[['Product Name', 'product_category', 'is_return', 'meesho_price_clean', 'order_status']].copy()
            product_categories_data.to_excel(writer, sheet_name='Product_Categories', index=False)
        
        print("✓ Data saved to: meesho_analysis_results.xlsx")
        print("  - Sheet 1: Category_Analysis")
        print("  - Sheet 2: Price_Range_Analysis") 
        print("  - Sheet 3: Summary")
        print("  - Sheet 4: Product_Categories (all products with their categories)")
    
    def create_real_data_plots(self, category_returns, price_returns, overall_rate):
        """Create plots with real data"""
        print("Creating plots with real data...")
        
        # Create the plots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
        fig.suptitle(f'Meesho Real Data Analysis - Return Rate Analysis (Overall: {overall_rate:.1f}%)', 
                     fontsize=20, fontweight='bold', color=JAMUNI, y=0.98)
        
        # Plot 1: Return Rate by Category (Real Data) - showing percentage of total orders
        categories = category_returns.index
        category_rates = category_returns['percentage_of_total_orders'].values
        
        # Use only Meesho brand colors
        colors = [JAMUNI, AAM, JAMUNI, AAM, JAMUNI, AAM][:len(categories)]
        
        bars1 = ax1.bar(categories, category_rates, 
                       color=colors, edgecolor='white', linewidth=2, alpha=0.9)
        
        ax1.set_title('Percentage of Total Orders by Category', 
                     fontsize=16, fontweight='bold', 
                     color=JAMUNI, pad=20)
        ax1.set_xlabel('Product Categories', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Percentage of Total Orders (%)', fontsize=14, fontweight='bold')
        ax1.tick_params(axis='x', rotation=45, labelsize=12)
        ax1.tick_params(axis='y', labelsize=12)
        ax1.grid(False)
        ax1.set_ylim(0, max(category_rates) * 1.2)
        
        # Add value labels and order counts
        for i, (bar, rate, orders, returns) in enumerate(zip(bars1, category_rates, category_returns['total_orders'], category_returns['returns'])):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{rate:.1f}%\n({orders} orders, {returns} returns)', ha='center', va='bottom', 
                    fontweight='bold', fontsize=10)
        
        # Plot 2: Return Rate by Price Range (Real Data)
        price_ranges = price_returns.index
        price_rates = price_returns['return_rate'].values
        
        bars2 = ax2.bar(price_ranges, price_rates,
                       color=AAM, edgecolor='white', linewidth=2, alpha=0.9)
        
        ax2.set_title('Return Rate by Price Range (Real Data)', 
                     fontsize=16, fontweight='bold', 
                     color=JAMUNI, pad=20)
        ax2.set_xlabel('Price Ranges (₹)', fontsize=14, fontweight='bold')
        ax2.set_ylabel('Return Rate (%)', fontsize=14, fontweight='bold')
        ax2.tick_params(axis='x', rotation=45, labelsize=12)
        ax2.tick_params(axis='y', labelsize=12)
        ax2.grid(False)
        ax2.set_ylim(0, max(price_rates) * 1.2)
        
        # Add value labels and order counts
        for bar, rate, orders in zip(bars2, price_rates, price_returns['total_orders']):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{rate:.1f}%\n({orders} orders)', ha='center', va='bottom', 
                    fontweight='bold', fontsize=10)
        
        plt.tight_layout()
        plt.savefig('real_meesho_data_analysis.png', dpi=300, bbox_inches='tight')
        plt.show()
        
        return fig
    
    def analyze_product_names(self):
        """Analyze product names to understand patterns"""
        print("Analyzing product names...")
        
        # Get all product names
        all_products = ' '.join(self.merged_data['Product Name'].astype(str))
        
        # Extract features
        features = self.extract_product_features(all_products)
        
        # Get most common words
        word_freq = Counter(features)
        most_common = word_freq.most_common(20)
        
        print("\nMost Common Words in Product Names:")
        for word, count in most_common:
            print(f"  {word}: {count}")
        
        return most_common
    
    def run_complete_analysis(self):
        """Run the complete real data analysis"""
        print("Starting Real Meesho Data Analysis...")
        print("=" * 60)
        
        # Load and preprocess data
        self.load_real_data()
        self.preprocess_data()
        
        # Categorize products using NLP
        self.categorize_products_nlp()
        
        # Calculate real return rates
        category_returns, price_returns, overall_rate = self.calculate_real_return_rates()
        
        # Save data to Excel
        self.save_data_to_excel(category_returns, price_returns, overall_rate)
        
        # Create plots with real data
        self.create_real_data_plots(category_returns, price_returns, overall_rate)
        
        # Analyze product names
        self.analyze_product_names()
        
        print("\n" + "="*60)
        print("="*60)
        print(" Used CSV data from your files")
        print(" Applied NLP analysis for product categorization")
        print(" Calculated real return rates from actual data")
        print(" Created plots with Meesho brand colors")
        print(" Generated: real_meesho_data_analysis.png")
        
        return category_returns, price_returns, overall_rate

# Run the analysis
if __name__ == "__main__":
    analyzer = RealMeeshoAnalysis()
    category_returns, price_returns, overall_rate = analyzer.run_complete_analysis()
