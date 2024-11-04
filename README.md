# Excel VBA Multi-Plot Chart Generator

This Excel VBA script, `MuiliPlot`, automates the creation of multiple smooth XY scatter charts from data stored in the **DataSheet** worksheet. The script dynamically processes data and generates charts based on the specified configurations, with the ability to customize chart titles and axis labels.

![Interface Screenshot](path_to_image_file) <!-- Replace with the actual path to your image in the repository -->

## Overview

The script is designed to:
- Create scatter plots for specified data ranges in the **DataSheet**.
- Customize the chart appearance, including titles, font styles, and axis labels.
- Automatically adjust the position and size of each chart on the worksheet.
- Set specific axis properties, such as hiding negative values on the x-axis and using round-dot gridlines.

## Key Features

- **Dynamic Data Processing**: The script loops through a specified column, identifying data ranges and generating corresponding charts.
- **Customizable Chart Titles**: Uses values from specified cells to create unique chart titles and axis labels.
- **Series Renaming**: Custom names are assigned to data series for better chart readability.
- **Formatted Axes and Titles**: Applies font customization to titles and axes, including the use of "Times New Roman".

## Requirements

- Microsoft Excel with VBA enabled.
- A worksheet named **DataSheet** containing the required data in a specified format.
- Data columns: 
  - Column E: Title for the charts.
  - Column F: Start of the data range.
  - Column H: High Flood Level (HFL) values.
- Cells K1, K2, K3, K4, and K5 in **DataSheet** for configuration:
  - **K1**: Base title for the charts.
  - **K2**: Label for the X-axis.
  - **K3**: Label for the Y-axis.
  - **K4**: Chart width (in inches).
  - **K5**: Chart height (in inches).

## Usage Instructions

1. **Setup DataSheet**: Ensure that your data is correctly arranged in the **DataSheet** worksheet:
   - Titles in Column E.
   - X-axis and Y-axis data in Columns F to H.
   - HFL values in Column H.
   - Configuration settings in cells K1, K2, K3, K4, and K5.

2. **Run the Script**:
   - Press `ALT + F11` to open the VBA editor in Excel.
   - Copy and paste the provided `MuiliPlot` script into a new module.
   - Press `F5` or run the `MuiliPlot` macro from the VBA editor to generate charts.

## Outputs

- **Smooth XY Scatter Charts**: Each chart is plotted using data from Columns F to H.
- **Customized Titles and Labels**: Charts are titled and labeled based on the values in **DataSheet**.
- **Formatted Gridlines**: Value axis gridlines use a round-dot dash style for a cleaner look.

## Notes

- Ensure that your data does not contain any empty rows in the specified range, as the script relies on contiguous data.
- Adjust the chart size and position using the values in **K4** and **K5**.
- The series names are customized for the first chart; modify the code if additional customization is needed.

## Troubleshooting

- **Error Handling**: The script includes basic error handling for renaming series. If an error occurs, a message box will display the error description.
- **Customization**: The script can be easily modified to suit different data structures or chart types.

