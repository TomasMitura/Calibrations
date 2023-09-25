import pandas as pd
from sklearn.linear_model import LinearRegression

def calculate_energy_and_save():
    file_pathway = r"C:\Users\mitura\Documents\Python_scripts\Calibrations\Calibration080923.xlsx"
    energy_calib_file_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\L8_A_up_filtered.xlsx"
    output_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VED_calibration080923a.xlsx"

    # Read energy calibration data from Excel file
    energy_data = pd.read_excel(energy_calib_file_path, header=0, usecols=[0])

    # Read Calibration data from an Excel file without headers
    data = pd.read_excel(file_pathway, header=None, names=['Voltage', 'Energy'])

    # Remove data points where energy is larger than the next data point
    data = data[data['Energy'] <= data['Energy'].shift(-1)]

    # Remove data points with energy output <= 5e-08
    data = data[data['Energy'] > 5e-08]

    # Perform linear regression to find the relationship between voltage and energy
    regressor = LinearRegression()
    regressor.fit(data[['Voltage']], data['Energy'])

    slope = regressor.coef_[0]
    intercept = regressor.intercept_

    print("Linear Regression Parameters:")
    print(f"Slope: {slope}")
    print(f"Intercept: {intercept}")

    # Calculate estimated energy output per area based on linear relationship
    estimated_energy_output = regressor.predict(energy_data.values.reshape(-1, 1))

    # Read depth data from EnergyCalibration.xlsx
    depth_data = pd.read_excel(energy_calib_file_path, header=0, usecols=[1], names=['Depth'])
    
    # Convert depth values to absolute values and divide by 1000
    depth_data['Depth'] = abs(depth_data['Depth']) / 1000

    # Create a new DataFrame with calculated energy output per area and depth
    output_data = {
        'Voltage': energy_data['Voltage'],
        'Energy': estimated_energy_output,
        'Depth': depth_data['Depth']
    }
    output_df = pd.DataFrame(output_data)

    # Save the modified DataFrame to a new Excel file
    output_df.to_excel(output_path, index=False)

    print(f"Estimated energy output per area and depth added and saved to {output_path}.")
if __name__ == "__main__":
    calculate_energy_and_save()

