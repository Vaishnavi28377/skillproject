import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Load the Data
file_path = "D:\storesdataset.xlsx"  # Replace with your file
try:
    df = pd.read_excel(file_path)  # Or pd.read_excel(file_path) if Excel
    print("Data loaded successfully!")
except FileNotFoundError:
    print(f"Error: File not found at {file_path}")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    exit()

# 2. Initial Exploration
print("\nFirst 5 rows:")
print(df.head())

print("\nData Info:")
print(df.info())

print("\nDescriptive Statistics:")
print(df.describe())

# 3. Handle Missing Values
print("\nMissing Values:")
print(df.isnull().sum())

# Imputation or removal based on context
# Example: Filling missing values with the mean
# df['Sales'] = df['Sales'].fillna(df['Sales'].mean())

# 4. Data Type Conversion (if needed)
#Example - already inferred properly
# df['Postal Code'] = df['Postal Code'].astype(str)  # Treat as categorical if needed
# df['Order Date'] = pd.to_datetime(df['Order Date'])

# 5. Univariate Analysis

# Numerical Variables
numerical_cols = df.select_dtypes(include=['number']).columns

for col in numerical_cols:
    plt.figure(figsize=(8, 6))
    sns.histplot(df[col], kde=True) #or distplot, but distplot is depreciated
    plt.title(f'Distribution of {col}')
    plt.show()

# Categorical Variables
categorical_cols = df.select_dtypes(include=['object']).columns

for col in categorical_cols:
    plt.figure(figsize=(8, 6))
    sns.countplot(x=col, data=df)
    plt.title(f'Count of {col}')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

# 6. Bivariate Analysis

# Numerical vs. Numerical (Correlation)
if len(numerical_cols) > 1:
    correlation_matrix = df[numerical_cols].corr()
    plt.figure(figsize=(10, 8))
    sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm')
    plt.title('Correlation Matrix')
    plt.show()

# Numerical vs. Categorical (Boxplots)
for cat_col in categorical_cols:
    for num_col in numerical_cols:
        plt.figure(figsize=(10, 6))
        sns.boxplot(x=cat_col, y=num_col, data=df)
        plt.title(f'{num_col} by {cat_col}')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.show()

# Categorical vs. Categorical (Crosstabulations)
for i in range(len(categorical_cols)):
    for j in range(i + 1, len(categorical_cols)):
        col1 = categorical_cols[i]
        col2 = categorical_cols[j]
        cross_tab = pd.crosstab(df[col1], df[col2])
        print(f"\nCrosstabulation of {col1} and {col2}:")
        print(cross_tab)

        plt.figure(figsize=(8, 6))
        sns.heatmap(cross_tab, annot=True, cmap='YlGnBu', fmt='d')
        plt.title(f'Heatmap of {col1} and {col2}')
        plt.show()

# 7. Further Analysis & Feature Engineering

# Profit Margin
df['Profit_Margin'] = df['Profit'] / df['Sales'] #creates column
print("\nDescriptive Statistics (with Profit Margin):")
print(df.describe())

#Sales by Region plot
plt.figure(figsize=(10, 6))
sns.barplot(x='Region', y='Sales', data=df)
plt.title('Sales by Region')
plt.show()

# 8. Outlier Detection (Boxplots)
for col in numerical_cols:
    plt.figure(figsize=(8, 6))
    sns.boxplot(y=df[col])
    plt.title(f'Boxplot of {col}')
    plt.show()

# 9. Save Processed Data (Optional)
# df.to_csv('sales_data_processed.csv', index=False)
# print("Processed data saved to sales_data_processed.csv")