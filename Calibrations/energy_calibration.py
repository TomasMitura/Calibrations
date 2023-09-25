import pandas as pd
from sklearn.linear_model import LinearRegression

def generate_voltage_calibration(calibration_file_path, energy_area_depth_calib_file_path, output_path):
    # Read energy area depth calibration data from Excel file
    energy_area_depth_data = pd.read_excel(energy_area_depth_calib_file_path)

    # Read data from the calibration
    new_data = pd.read_excel(calibration_file_path, header=None, skiprows=1, names=['Voltage', 'Energy'])

    # Remove data points where energy output is larger than the next data point
    new_data = new_data[new_data['Energy'] <= new_data['Energy'].shift(-1)]

    # Remove data points with energy output <= 5e-08
    new_data = new_data[new_data['Energy'] > 5e-08]

    # Perform linear regression to find the relationship between voltage and energy
    regressor_energy = LinearRegression()
    regressor_energy.fit(new_data[['Energy']], new_data['Voltage'])  # Invert the regression equation

    # Calculate estimated voltages based on inverted linear relationship
    estimated_voltages = regressor_energy.predict(energy_area_depth_data[['Energy']].values.reshape(-1, 1))

    # Create a new DataFrame with calculated voltages and other columns from the original data
    output_data = {
        'Voltage': estimated_voltages,
        'Energy': energy_area_depth_data['Energy'],   
        'Depth': energy_area_depth_data['Depth']
    }
    output_df = pd.DataFrame(output_data)

    # Save the modified DataFrame to a new Excel file
    output_df.to_excel(output_path, index=False)

    print(f"Estimated voltages added and saved to {output_path}.")

# Call the function with appropriate inputs
if __name__ == "__main__":
    calibration_file_path = r"C:\Users\mitura\Documents\Python_scripts\Calibrations\Calibration060923.xlsx"
    energy_area_depth_calib_file_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\EnergyAreaDepthCalibration.xlsx"
    output_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VEDCalibration060923.xlsx"
    
    generate_voltage_calibration(calibration_file_path, energy_area_depth_calib_file_path, output_path)
