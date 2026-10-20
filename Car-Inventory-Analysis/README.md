# Car Inventory Analysis Using Excel

## Project Overview

This project involved cleaning, transforming, and exploring a car inventory dataset in Microsoft Excel. The work focused on extracting useful fields from a combined Car ID, standardising vehicle makes and models, calculating additional metrics, and summarising driver mileage.

## Tools Used

- Microsoft Excel

## Work Performed

### 1. Data Extraction and Transformation
- Used `LEFT` and `MID` to extract the car make, model, and manufacturing year from the Car ID.
- Calculated vehicle age from the manufacturing year.
- Created a new Car ID by combining relevant information.

### 2. Data Cleaning and Error Handling
- Adjusted the vehicle-age calculation to account for older manufacturing years represented with two digits.
- Accounted for potential division-by-zero errors when calculating miles driven per year.
- Identified and corrected extra spaces in abbreviated make names that caused a `VLOOKUP` mismatch.

### 3. Data Standardisation and Calculated Fields
- Used `VLOOKUP` to map abbreviated car makes and models to their full names.
- Used `IF` to check whether a car was covered by insurance.
- Calculated miles driven per year.

### 4. Exploratory Analysis and Visualisation
- Created a scatter chart to explore the relationship between vehicle age and mileage.
- Observed a positive relationship in this dataset: older cars generally had higher mileage.
- Created a PivotTable to summarise the miles driven by each driver and a bar chart to compare the results.

## Key Observation

The scatter chart suggested that older cars generally had higher mileage in this dataset. This is an observed relationship in the data, not evidence that vehicle age alone causes higher mileage.

## Skills Practised

- Excel formulas: `LEFT`, `MID`, `IF`, and `VLOOKUP`
- Data cleaning and standardisation
- Error handling and troubleshooting
- Creating calculated fields
- PivotTables
- Charts and exploratory data analysis

## Project File

- `Car_Inventory_Analysis.xlsx` — the Excel workbook for this project.

## Conclusion

This project provided practical experience in turning combined and inconsistent data into a more structured format for analysis. It also reinforced the importance of checking data quality when formulas return unexpected results.
