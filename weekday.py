import os

import matplotlib
matplotlib.use('TkAgg')  # Set the backend before importing pyplot

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from tkinter import filedialog, Tk

plt.rcParams['toolbar'] = 'None'
# Create a Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window

# Open a file dialog for selecting the Excel file
file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

if not file_path:
    print("No file selected. Exiting...")
    exit()

# Read the Excel file, ignoring the first three rows
df = pd.read_excel(file_path, header=3, parse_dates=['Datum'])

# Preprocess the data
df['Year'] = df['Datum'].dt.year  # Extract the year from the date
df['Weekday'] = df['Datum'].dt.day_name()  # Extract the weekday name
weekday_energy = df.groupby(['Year', 'Weekday'])['Gesamtverbrauch (Wh)'].sum() / 1000  # Group by year and weekday, then sum energy consumption and convert to kW/h

# Calculate the percentage of energy consumption per weekday relative to the total consumption for each year
total_energy_per_year = df.groupby('Year')['Gesamtverbrauch (Wh)'].sum() / 1000  # Total energy consumption per year in kW/h
weekday_energy_percent = weekday_energy.div(total_energy_per_year, level='Year') * 100  # Calculate percentage

# Get unique years in the dataset
years = df['Year'].unique()

# Define the desired order of weekdays
weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

# Reorder the data based on the desired order of weekdays
weekday_energy_percent = weekday_energy_percent.reindex(weekday_order, level='Weekday')

# Get the file name from the file path
file_name = os.path.basename(file_path)

# Plot the graph
plt.figure(figsize=(12, 6))

# Define the width of each bar
bar_width = 0.2

# Iterate over each year and plot energy consumption percentage per weekday with a different color and x-offset
for i, year in enumerate(years):
    x_offset = i * bar_width
    year_data = weekday_energy_percent.loc[year]
    x_values = np.arange(len(year_data.index)) + x_offset
    plt.bar(x_values, year_data.values, width=bar_width, label=str(year), alpha=0.7)
    
    # Add text labels on top of each bar, covering it
    for x, y in zip(x_values, year_data.values):
        plt.text(x, y + 0.5, f'{y:.1f}%', ha='center', va='bottom', rotation=90)

plt.title('Percentage of Energy Consumption by Weekday (Comparison Across Years) - {}'.format(file_name))
plt.xlabel('Weekday')
plt.ylabel('Percentage of Energy Consumption (%)')
# Set the tick labels on the x-axis using the specified order of weekdays
plt.xticks(np.arange(7) + bar_width * (len(years) - 1) / 2, weekday_order)
plt.legend(title='Year')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.ylim(top=plt.ylim()[1] * 1.1)
plt.tight_layout()
plt.show(block=True)
