import pandas as pd
import numpy as np
from scipy.interpolate import interp1d

def linear_calibration(input_file_path, output_file_path):
    # Load data from Excel file
    df = pd.read_excel(input_file_path)

    # Extract energy and depth columns from the DataFrame
    energy_data = df["Voltage"].values
    depth_data = df["Depth"].values

    # Create an interpolation function
    interp_func = interp1d(depth_data, energy_data, kind='linear', fill_value=(energy_data[0], energy_data[-1]))

    # Define the desired range of depths and the increment
    start_depth = 25
    end_depth = 120
    depth_increment = 0.1
    depth = start_depth

    # Create an empty list to store the results
    results = []

    # Iterate through the desired range of depths with the specified increment
    while depth <= end_depth:
        estimated_energy = interp_func(depth)
        results.append((float(depth), float(estimated_energy)))
        depth += depth_increment

    # Create a DataFrame from the results
    estimated_df = pd.DataFrame(results, columns=["Depth", "Voltage"])

    # Save the estimated DataFrame to a new Excel file
    estimated_df.to_excel(output_file_path, index=False)

    print(f"Estimated energies added and saved to {output_file_path}.")

if __name__ == "__main__":
    input_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VED_Calibration080923a.xlsx"
    output_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VD_Interpolated080923a.xlsx"
    linear_calibration(input_path, output_path)
