import os

import matplotlib
matplotlib.use('MacOSX')  # Set the backend before importing pyplot

import pandas as pd
import matplotlib.pyplot as plt
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
df['Month'] = df['Datum'].dt.month  # Extract the month from the date
monthly_energy = df.groupby(['Year', 'Month'])['Gesamtverbrauch (Wh)'].sum() / 1000  # Group by year and month, then sum energy consumption and convert to kW/h

# Get the file name from the file path
file_name = os.path.basename(file_path)

# Plot the graph
plt.figure(figsize=(12, 6))

# Define the width of each bar
bar_width = 0.2

# Iterate over each year and plot monthly energy consumption with a different color and x-offset
for i, (year, data) in enumerate(monthly_energy.groupby(level=0)):
    x_offset = i * bar_width
    plt.bar(data.index.get_level_values('Month') + x_offset, data, width=bar_width, label=str(year), alpha=0.7)

plt.title('Monthly Energy Consumption Comparison - {}'.format(file_name))
plt.xlabel('Month')
plt.ylabel('Energy Consumption (kW/h)')
plt.xticks(range(1, 13), ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
plt.legend(title='Year')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show(block=True)
