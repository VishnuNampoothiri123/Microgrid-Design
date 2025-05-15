import numpy as np
import pandas as pd
import os

# Loading Files
load_profile_excel = pd.read_excel("Load Profile.xlsx", usecols=['Load (kW)'])
UNMMaxLoad = 3006.1082
NGEngineSize = UNMMaxLoad * 0.6
PV_profile_excel = pd.read_excel("PV Profile.xlsx", usecols=['PV Profile (kW)']) 
BESS_sizekW = float(input('Enter the BESS_Size (kW): '))
BESS_sizekWhr = float(input('Enter the BESS_Size (kWhr): '))

def calculate_result(B, C, D, kW):
    return np.where((B - C - D) > 0, np.where((B - C - D) < kW, B - C - D, kW), 0)

# Initialize an empty list to store dataframes for each 8-row window
dfs = []

# Initialize lists to store statistics for each window
outage_numbers = []
start_hours = []
end_hours = []
min_excess_powers = []
max_excess_powers = []
bess_soc_checks = []
excess_generation_power_checks = []
max_loads = []
min_excess_generation_percent = []

# Initialize variables to store overall statistics
dispatch_violations = 0
min_excess_power = float('inf')
min_excess_power_window = None
max_excess_power = float('-inf')
max_excess_power_window = None

# Iterate over the rows with a step of 1, and take 8-row windows
for i in range(len(load_profile_excel) - 7):
	window = load_profile_excel.iloc[i:i+8]
	# Create a dataframe for the current window
	dispatchModelTable = pd.DataFrame()
	dispatchModelTable['Hours'] = range(i+1, i+9)  # Adjust hours based on the window
	dispatchModelTable['Load'] = window['Load (kW)'].values
	dispatchModelTable['NG'] = [NGEngineSize * 0.9] * 8  
	dispatchModelTable['PV'] = (PV_profile_excel.iloc[i:i+8]['PV Profile (kW)'] * UNMMaxLoad * 0.2).values
	dispatchModelTable['BESS'] = calculate_result(dispatchModelTable['Load'], dispatchModelTable['NG'], dispatchModelTable['PV'], BESS_sizekW)
	# Calculate BESS-Energy within each window
	BESS_Energy = [BESS_sizekWhr - dispatchModelTable['BESS'].iloc[0]]
	for j in range(1, 8):
		BESS_Energy.append(BESS_Energy[j-1] - dispatchModelTable['BESS'].iloc[j])

	dispatchModelTable['BESS-Energy'] = BESS_Energy

	# Calculate BESS_SOC%
	BESS_SOC = (dispatchModelTable['BESS-Energy'] / BESS_sizekWhr) * 100
	# Set values less than 0 to 0
	BESS_SOC = np.where(BESS_SOC < 0, 0, BESS_SOC)

	# Add BESS_SOC% column to the dataframe
	dispatchModelTable['BESS_SOC%'] = BESS_SOC
	dfs.append(dispatchModelTable)

	# Calculate Excess Power
	dispatchModelTable['Excess Power'] = dispatchModelTable['NG'] + dispatchModelTable['PV'] + BESS_sizekW - dispatchModelTable['Load']

	# Calculate Excess Generation Power Check for the current window
	excess_generation_power_check = (0.2 * window['Load (kW)'].max())

	# Append statistics for the current window to the lists
	outage_numbers.append(i + 1)
	start_hours.append(i + 1)
	end_hours.append(i + 8)
	max_loads.append(window['Load (kW)'].max())
	min_excess_powers.append(dispatchModelTable['Excess Power'].min())
	max_excess_powers.append(dispatchModelTable['Excess Power'].max())
	bess_soc_checks.append(BESS_SOC.min())
	excess_generation_power_checks.append(excess_generation_power_check)
	min_excess_generation_percent.append((dispatchModelTable['Excess Power'].min() / excess_generation_power_check) * 100)

	# Update overall statistics
	dispatch_violations += ((dispatchModelTable['Excess Power'] < 0).any() or (BESS_SOC < 10).any())
	if dispatchModelTable['Excess Power'].min() < min_excess_power:
		min_excess_power = dispatchModelTable['Excess Power'].min()
		min_excess_power_window = (i + 1, i + 8)
	if dispatchModelTable['Excess Power'].max() > max_excess_power:
		max_excess_power = dispatchModelTable['Excess Power'].max()
		max_excess_power_window = (i + 1, i + 8)

# Create a new dataframe with the desired columns
summary_df = pd.DataFrame({
    'Outage Number': outage_numbers,
    'Start Hour': start_hours,
    'End Hour': end_hours,
    'Maximum Load': max_loads,
    'Min Excess Power': min_excess_powers,
    'BESS SOC% Check': bess_soc_checks,
    'Excess Generation Power Check': excess_generation_power_checks,
    'Min Excess Generation %': min_excess_generation_percent
})

# Concatenate all dataframes in the list into a single dataframe
final_df = pd.concat(dfs, ignore_index=True)

# Export dataframe to Excel
final_df.to_excel("dispatch_model.xlsx", index=False)
print("Excel file 'dispatch_model.xlsx' has been created successfully.")

# Export the summary dataframe to Excel
summary_df.to_excel("dispatch_model_summary.xlsx", index=False)
print("Summary dataframe has been exported to 'dispatch_model_summary.xlsx'.")
print("\n")

print("The number of dispatch violations:", dispatch_violations)
print("\n")
if (final_df['Excess Power'] < 0).any():
	max_excess_powerkW = final_df[final_df['Excess Power'] < 0]['Excess Power'].min()
	print("There is a generation shortage.")
	print("The maximum amount of generation storage is: ",max_excess_powerkW,"kW")
else:
	print("There is no excess amount of generation.")
	print("Min Excess Power of all the windows in kW as well as which window this is located in:")
	print("The minimum excess power from hour", min_excess_power_window, "is", min_excess_power, "kW.")
	min_excess_value_percent = (min_excess_power / (0.2 * max(max_loads))) * 100
	print("Percentage of the minimum excess value that is 20% of the max load:", min_excess_value_percent, "%")
	print("\n")



