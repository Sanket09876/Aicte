import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the data from the Excel file
file_path = r"C:\Users\SANKET DAS\Desktop\hotel_bookings.xlsx"
df = pd.read_excel(file_path)

# Print basic information about the DataFrame
print(df.info())
print(df.describe())
print(df.head())

# Check if required columns are present
required_columns = [
    'hotel', 'arrival_date_year', 'arrival_date_month', 'arrival_date_day_of_month',
    'adr', 'stays_in_week_nights', 'total_of_special_requests', 'lead_time'
]

# Ensure all required columns are in the DataFrame
for col in required_columns:
    if col not in df.columns:
        raise ValueError(f"Missing column: {col}")

# Plot number of bookings per hotel type
plt.figure(figsize=(10, 6))
sns.countplot(data=df, x='hotel')
plt.title('Number of Bookings per Hotel Type')
plt.xlabel('Hotel Type')
plt.ylabel('Number of Bookings')
plt.show()

# Create a new column for arrival_date by combining year, month, and day
df['arrival_date'] = pd.to_datetime(
    df[['arrival_date_year', 'arrival_date_month', 'arrival_date_day_of_month']]
)

# Plot average daily rate (ADR) over time
plt.figure(figsize=(14, 7))
df.groupby('arrival_date')['adr'].mean().plot()
plt.title('Average Daily Rate Over Time')
plt.xlabel('Arrival Date')
plt.ylabel('Average Daily Rate (ADR)')
plt.show()

# Plot ADR vs. number of week nights stayed
plt.figure(figsize=(12, 6))
sns.boxplot(data=df, x='stays_in_week_nights', y='adr')
plt.title('ADR vs. Number of Week Nights Stayed')
plt.xlabel('Number of Week Nights Stayed')
plt.ylabel('Average Daily Rate (ADR)')
plt.show()

# Plot number of special requests
plt.figure(figsize=(10, 6))
sns.countplot(data=df, x='total_of_special_requests')
plt.title('Number of Special Requests')
plt.xlabel('Total Number of Special Requests')
plt.ylabel('Number of Bookings')
plt.show()

# Plot lead time vs. number of special requests
plt.figure(figsize=(10, 6))
sns.scatterplot(data=df, x='lead_time', y='total_of_special_requests')
plt.title('Lead Time vs. Number of Special Requests')
plt.xlabel('Lead Time (Days)')
plt.ylabel('Number of Special Requests')
plt.show()

# Save the processed DataFrame to a new Excel file
output_path = r"C:\Users\SANKET DAS\Desktop\processed_hotel_bookings.xlsx"
df.to_excel(output_path, index=False)
