import wbdata
import pandas as pd
from datetime import datetime

# Define countries and indicators (e.g., GDP for Turkey and the USA)
countries = ["TUR", "USA"]
indicator = {"NY.GDP.MKTP.CD": "GDP (Current USD)"}

# Set date range
start_date = datetime(2020, 1, 1)
end_date = datetime(2024, 1, 1)

# Specify the output path for the Excel file
output_path = "world_bank_gdp_data.xlsx"

# Create the Excel file with each country on a separate sheet
with pd.ExcelWriter(output_path) as writer:
    for country in countries:
        data = wbdata.get_dataframe(indicator, country=country)
        if not data.empty:
            data = data.reset_index()
            data.to_excel(writer, sheet_name=country)
        else:
            print(f"Warning: No data found for {country}.")

print(f"Data successfully saved to '{output_path}'")
