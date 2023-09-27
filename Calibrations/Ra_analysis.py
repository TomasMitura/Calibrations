import pandas as pd

# Read the Excel file into a DataFrame, assuming there are no headers
file_path = r"C:\Users\mitura\Documents\Python_scripts\Calibrations\Ra_calculation_w1um_2mms.xlsx"
df = pd.read_excel(file_path, header=None, names=['Position', 'Depth'])

# Set the threshold for filtering
threshold_value = 96000

# Filter rows where the absolute value of 'Depth' is greater than the threshold
filtered_df = df[abs(df['Depth']) > threshold_value]

# Calculate the total number of segments
segment_size = 250 #60 points is 10 um
total_segments = len(filtered_df) // segment_size

# Initialize a list to store the Ra values for each segment
ra_values = []

# Calculate Ra for each segment
for i in range(total_segments):
    start_idx = i * segment_size
    end_idx = start_idx + segment_size
    segment_data = filtered_df.iloc[start_idx:end_idx]
    
    # Exclude the first 30 data points from the first segment
    if i == 0:
        segment_data = segment_data.iloc[50:]
    
    mean_depth = segment_data['Depth'].mean()
    ra_segment = (1 / segment_size) * segment_data['Depth'].sub(mean_depth).abs().sum()
    ra_values.append(ra_segment)

    print(f"Ra for segment {i+1}: {ra_segment}")

# Calculate the mean of all Ra values
mean_ra = sum(ra_values) / len(ra_values)

print("Mean Ra value for all segments:", mean_ra)
