
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from utils import create_financial_template


# Create the template if needed:
# create_financial_template()


# Load the Excel file
file_path = "financial_data.xlsx"  # Change to your file path
df = pd.read_excel(file_path, engine="openpyxl")

# Extract gross profit values from Profit & Loss Statement
PL_gross_profit = [int(value) if not pd.isna(value) else 0 for value in df.iloc[3, 1:16]]
# Extract operating profit values from Profit & Loss Statement
PL_operating_profit = [int(value) if not pd.isna(value) else 0 for value in df.iloc[7, 1:16]]
# Extract gross profit values from Profit & Loss Statement
PL_profit_before_tax = [int(value) if not pd.isna(value) else 0 for value in df.iloc[10, 1:16]]

# Extract Ending cash values from Cash Flow Statement
CF_ending_cash = [int(value) if not pd.isna(value) else 0 for value in df.iloc[21, 1:16]]

# define a list for the months
months = list(range(1, 13)) + [12, 24, 36]

# Print the extracted values if needed:
# print(PL_gross_profit)
# print(PL_operating_profit)
# print(PL_profit_before_tax)
# print(CF_ending_cash)

## Create the Profit & Loss statement graph

plt.figure(figsize=(14, 7))

# Plot each dataset
plt.plot(months[:12] + months[12+1:], PL_gross_profit[:12] + PL_gross_profit[12+1:]  , marker=' ', linestyle='dotted', linewidth=2, alpha=1,color='green', label="Gross Profit")
plt.plot(months[:12] + months[12+1:], PL_operating_profit[:12] + PL_operating_profit[12+1:], marker=' ', linestyle='dotted', linewidth=2, alpha=1,color='red',  label="Operating Profit")
plt.plot(months[:12] + months[12+1:], PL_profit_before_tax[:12] + PL_profit_before_tax[12+1:], marker='o', linestyle='-',linewidth =2.5, markersize = 6, color="#1B4965", label="Profit Before Tax")

# Labels and Title
plt.xlabel("Time (months)", fontsize=24, fontweight='bold', labelpad=30)
plt.ylabel("Profit (USD)", fontsize=24, fontweight='bold', labelpad=60, rotation=0)
# plt.title("Profit & Loss Statement")
plt.xlim(1, 36)  # Focus mainly on 1-12 (leave space for 24 & 36)

# Apply Logarithmic Scale for better visualization
plt.yscale("log")

# Format Y-axis labels in dollar notation ($10K, $100K, $1M, $10M)
plt.yticks(
    [10_000, 100_000, 1_000_000, 5_000_000],  # Scale levels
    ["0", "100K", "1M", "5M"]  # Labels
)
# Format Y-axis labels in dollar notation ($10K, $100K, $1M, $10M)
plt.xticks(
    [1, 3, 6, 12, 24, 36],  # Scale levels
    ["1", "3", "6", "12", "24", "36"]  # Labels
)

ax = plt.gca()  # Get the current axis
ax.spines["top"].set_visible(False)   # Hide the top border
ax.spines["right"].set_visible(False)  # Hide the left border

plt.xticks(fontsize=24, fontweight = 'bold')  # Increase size of X-axis tick labels
plt.yticks(fontsize=24, fontweight = 'bold')  # Increase size of Y-axis tick labels
# Show legend
plt.legend()

# Show grid with minor lines for better readability
plt.grid(True, which='both', linestyle='--', alpha=0.0)
plt.legend(fontsize=20, loc='lower right')

# Save the plot with transparent background
plt.savefig('profit_loss_plot.png', transparent=True, dpi=300, bbox_inches='tight')

# Display the graph
plt.show()


## Create the Cash Flow statement graph

# Create figure
plt.figure(figsize=(14, 7))

# Plot main data
plt.plot(months, CF_ending_cash, marker='o', linestyle='-', linewidth=2.5, markersize=6, color='#1B4965', label="Monthly Cash Flow" )

# Find zero crossing point
for i in range(len(CF_ending_cash)-1):
    if (CF_ending_cash[i] < 0 and CF_ending_cash[i+1] > 0) or (CF_ending_cash[i] > 0 and CF_ending_cash[i+1] < 0):
        # Interpolate to find exact crossing point
        x_cross = months[i] + (months[i+1] - months[i]) * (-CF_ending_cash[i])/(CF_ending_cash[i+1]-CF_ending_cash[i])

        # Add vertical line and annotation
        plt.axvline(x=x_cross, color='#62B6CB', linestyle='--', alpha=0.5)
        plt.annotate(f'Break-even at month {x_cross:.1f}',
                    xy=(x_cross, 0),
                    fontsize = 20, fontweight='bold',
                    xytext=(30, 30), textcoords='offset points',
                    ha='left', va='bottom',
                    bbox=dict(boxstyle='round,pad=0.5', fc='#CAE9FF', alpha=0.5),
                    arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'))

# Labels and Title
plt.xlabel("Time (months)", fontsize=24, fontweight='bold', labelpad=30)
plt.ylabel("Cash (USD)", fontsize=24, fontweight='bold', labelpad=60, rotation=0)
# plt.title("Cash Flow Statement")
plt.xlim(1, 36)

# Option 1: Use a symmetric log scale (good for data crossing zero)
plt.yscale('symlog', linthresh=100000)  # Linear scaling between -100k and 100k, log scaling beyond

# Add more ticks with focused range
plt.yticks(
    [-300_000, 0, 100_000, 1_000_000, 5_000_000],
    ["-300K", "0", "100K", "1M", "5M"]
)

plt.xticks(
    [1, 3, 6, 12, 24, 36],  # Scale levels
    ["1", "3", "6", "12", "24", "36"]  # Labels
)

# Add horizontal line at y=0
plt.axhline(y=0, color='#62B6CB', linestyle='--', alpha=0.3)

ax = plt.gca()  # Get the current axis
ax.spines["top"].set_visible(False)   # Hide the top border
ax.spines["right"].set_visible(False)  # Hide the left border

plt.xticks(fontsize=24, fontweight = 'bold')  # Increase size of X-axis tick labels
plt.yticks(fontsize=24, fontweight = 'bold')  # Increase size of Y-axis tick labels
# Show legend
plt.legend()

# Show grid with minor lines for better readability
plt.grid(True, which='both', linestyle='--', alpha=0.0)

plt.legend(fontsize=20, loc='lower right')


# Save with transparent background
plt.savefig('cash_flow_analysis.png', transparent=True, dpi=300, bbox_inches='tight')

# Display the graph
plt.show()

