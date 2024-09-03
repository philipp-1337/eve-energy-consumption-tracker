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
df.set_index('Datum', inplace=True)
weekly_energy = df.resample('W').sum() / 1000  # Convert watt hours to kW/h and resample weekly

# Calculate the moving average
window_size = 8  # Choose the window size for the moving average
moving_average = weekly_energy['Gesamtverbrauch (Wh)'].rolling(window=window_size).mean()

# Get the file name from the file path
file_name = os.path.basename(file_path)

# Plot the graph
plt.figure(figsize=(10, 6))
plt.bar(weekly_energy.index, weekly_energy['Gesamtverbrauch (Wh)'], width=5, color='skyblue', edgecolor='black', label='Weekly Consumption')
plt.plot(weekly_energy.index, moving_average, color='red', linestyle='--', label='Moving Average (Window Size = {})'.format(window_size))
plt.title('Weekly Energy Consumption with Moving Average - {}'.format(file_name))
plt.xlabel('Date')
plt.ylabel('Energy Consumption (kW/h)')
plt.legend()
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show(block=True)
