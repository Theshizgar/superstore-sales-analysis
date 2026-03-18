import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

#-------Settings

sns.set_theme(style="whitegrid")
plt.rcParams['figure.figsize'] = (12,6)

#------Load Data

df= pd.read_csv('Sample - Superstore.csv', encoding='windows-1252')
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Year'] = df['Order Date'].dt.year
print("Data Loaded!")


#---Chart 1: Sales by Category

fig, ax = plt.subplots()
category_sales = df.groupby ('Category')['Sales'].sum().sort_values()
colors = ['#2ecc71', '#3498db', '#e74c3c']
category_sales.plot(kind='barh', ax=ax, color=colors)
ax.set_title('Total Sales by Category', fontsize=16, fontweight='bold')
ax.set_xlabel('Total Sales (in USD)')
ax.set_ylabel('Category')
for i, v in enumerate (category_sales):
    ax.text(v + 500, i, f'${v:,.0f}', va='center', fontweight='bold')
plt.tight_layout()
plt.savefig('chart1_sales_by_category.png', dpi=150)
plt.close()
print("Chart 1 has been saved")

#------Chart 2: Monthly Sales Trend
fig, ax = plt.subplots()
df['Month'] = df['Order Date'].dt.to_period('M')
monthly_sales = df.groupby('Month')['Sales'].sum()
monthly_sales.plot(kind='line', ax=ax, color='#3498db', linewidth=2)
ax.set_title('Monthly Sales trent (2014-2017)', fontsize=16, fontweight='bold')
ax.set_xlabel('Month')
ax.set_ylabel('Total Sales (in USD)')
plt.tight_layout()
plt.savefig('Chart2_monthly_trend.png', dpi=150)
plt.close()
print("Chart 2 has been saved!")

#------Chart 3: Profit by Region

fig, ax = plt.subplots()
region_profit = df.groupby('Region')['Profit'].sum().sort_values()
colors = ['#e74c3c' if x < 0 else '#2ecc71' for x in region_profit]
region_profit.plot(kind='barh', ax=ax, color=colors)
ax.set_title('Total profit by region', fontsize=16, fontweight='bold')
ax.set_xlabel('Total profit (in USD)')
ax.set_ylabel('Region')
for i, v in enumerate(region_profit):
    ax.text(v + 100, i, f'${v:,.0f}', va = 'center', fontweight ='bold')
plt.tight_layout()
plt.savefig('Chart3_Profit_by_region', dpi=150)
plt.close()
print("Chart 3 saved!")

#-------Chart 4: Top 10 Sub Categories by Sales

fig, ax = plt.subplots()
sub_sales = df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=False).head(10)
sns.barplot(x=sub_sales.values, y=sub_sales.index, ax=ax, palette='Blues_r')
ax.set_title('Top 10 Sub Categories by Sales', fontsize=16, fontweight='bold')
ax.set_xlabel('Total Sales (in USD)')
ax.set_ylabel('Sub Category')
plt.tight_layout()
plt.savefig('Chart4_Top_SC.png', dpi=150)
plt.close()
print("Chart 4 saved!")


#------Chart 5: Discount vs profit (scatter)

fig, ax =plt.subplots()
colors=['#e74c3c' if x < 0 else '#2ecc71' for x in df['Profit']]
ax.scatter(df['Discount'], df['Profit'], alpha=0.4, c=colors, s=10)
ax.set_title('Discount vs Profit Analysis', fontsize=16, fontweight='bold')
ax.set_xlabel('Discount Rate')
ax.set_ylabel('Profit (in USD)')
ax.axhline(y=0, color='black', linestyle='--', linewidth=1)
plt.tight_layout()
plt.savefig('Chart5_discount_vs_proft.png', dpi=150)
plt.close()
print("Chart 5 saved!")

#-------Summary stats

print("\n Business Summary")
print("x"*40)
print(f"Total Revenue:               ${df['Sales'].sum():>12,.2f}")
print(f"Total Profit:                ${df['Profit'].sum():>12,.2f}")
print(f"Profit Margin:               {(df['Profit'].sum()/df['Sales'].sum()*100):>11.1f}%")
print(f"Total Orders:                {df['Order ID'].nunique():>12,}")
print(f"Total Customers:             {df['Customer ID'].nunique():>12,}")
print(f"Best Region:                 {df.groupby('Region')['Profit'].sum().idxmax():>12}")
print(f"Best Category:               {df.groupby('Category')['Profit'].sum().idxmax():>12}")
print(f"x"*40)
print(f"\nAll Charts Saved!")
# print(f"")
# print(f"")
# print(f"")


# print(f"")
