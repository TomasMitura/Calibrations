
from energy_calibration import generate_voltage_calibration #Based on the uFab calibration of voltage against energy output we get our VED calibration

if __name__ == "__main__":
    calibration_file_path = r"C:\Users\mitura\Documents\Python_scripts\Calibrations\Calibration220923.xlsx"
    energy_area_depth_calib_file_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VED_calibration080923a.xlsx"
    output_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VEDCalibration220923.xlsx"
    
    generate_voltage_calibration(calibration_file_path, energy_area_depth_calib_file_path, output_path)


from linearize_energy_calibration import linear_calibration

if __name__ == "__main__":
    input_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VEDCalibration220923.xlsx"
    output_path = r"C:\Users\mitura\Documents\Python_scripts\Depth\VD_Interpolated220923.xlsx"
    
    linear_calibration(input_path, output_path)



