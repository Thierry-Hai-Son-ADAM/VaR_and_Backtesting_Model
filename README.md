# Portfolio Analysis and VaR Backtesting Tool (Google Colab)

This Python script, designed to run in Google Colab, provides a tool for analyzing a stock portfolio and performing Value at Risk (VaR) backtesting. It retrieves up-to-date asset information using `yfinance`, updates an Excel file with this data, downloads historical price data, calculates both Historical and Monte Carlo VaR, and visualizes the results against daily portfolio returns.

You need to use the excel file called "My Portfolio Analysis.xlxs" that you can retrieve from this repo. You might want to use you're own personnal portfolio by changing the tickers and the quantity in the excel file.

## Features

* **Portfolio Data Update:** Fetches current company names, countries, industries, sectors, market capitalization, and latest prices for tickers listed in your Excel file.
* **Currency Conversion:** Automatically converts asset prices to EUR if they are in a different currency using current exchange rates from `yfinance`.
* **Market Value Calculation:** Calculates and updates the adjusted market value for each holding in the Excel file based on the latest prices.
* **Historical Data Export:** Exports historical closing prices for your portfolio assets to a dedicated sheet in the Excel file.
* **VaR Backtesting:** Performs backtesting of both Historical and Monte Carlo VaR over a specified period.
* **VaR Visualization:** Generates a plot showing the daily portfolio returns overlaid with the calculated VaR levels, highlighting breaches.
* **Excel Output:** Saves the updated portfolio information, historical prices, and VaR backtesting results into a new Excel file ready for download.

## Prerequisites

* **Google Colab Environment:** This script is specifically written for Google Colab due to its use of `google.colab.files` for file uploads and downloads.
* **Required Libraries:** The script depends on the following Python libraries. These are typically pre-installed in Google Colab, but the installation command is included in the script comments just in case.
    * `numpy`
    * `pandas`
    * `yfinance`
    * `datetime`
    * `matplotlib`
    * `openpyxl`
    * `scipy`
* **Portfolio Excel File:** You need an Excel file named `My Portfolio Analysis.xlsx` (or you will be prompted to upload one). This file must contain at least two sheets:
    * **'Portfolio Summary'**: This sheet should have columns for 'Asset', 'Ticker', and 'Quantity'. A 'Market Value' column is also expected for portfolio value calculation, and columns like 'Country', 'Industry', 'Sector', 'Market Cap', and 'yfinance adjusted MV' will be updated or created by the script.
    * **'Settings'**: This sheet should contain settings for the VaR calculation, including 'VaR Horizon (in days)', 'VaR percentile', and 'Simulation Number', and 'VaR range (in years)'. The script includes default values if these settings are not found.

## How to Use

1.  **Open in Google Colab:** Upload the Python script to your Google Drive and open it in Google Colab.
2.  **Run the Script:** Execute the Colab notebook cells sequentially.
3.  **Upload Your Excel File:** When prompted, upload your `My Portfolio Analysis.xlsx` file using the file uploader.
4.  **Enter VaR Parameters (Optional):** You will be asked to provide the last date for the VaR calculation and the backtesting period in days. You can press Enter to use the default values (today's date and 30 days respectively).
5.  **Wait for Execution:** The script will then proceed to fetch data, perform calculations, update the workbook, and generate the plot. This may take some time depending on the number of assets and the historical data range.
6.  **Download the Output:** Once the script finishes, a new Excel file with updated data (named like `Updated_My Portfolio Analysis_YYYYMMDD_HHMMSS.xlsx`) will be automatically downloaded to your local machine.
7.  **View the Plot:** The VaR backtesting plot will be displayed directly in the Google Colab output.

## Code Structure

The script is organized into several functions:

* `upload_excel_file()`: Handles the upload of the portfolio Excel file from your local machine to the Colab environment.
* `retrieve_asset_info(tickers)`: Fetches detailed information and the latest closing price for a list of tickers using `yfinance`.
* `copy_datas(wb, asset_info)`: Updates the 'Portfolio Summary' sheet in the Excel workbook with the retrieved asset information and calculated adjusted market values.
* `export_close_prices(wb, close_adj_df)`: Creates or updates a sheet in the Excel workbook with historical closing prices.
* `perform_var_backtesting(tickers, combined_data_df, settings, ptf_value, user_date, user_VaR_range)`: Calculates Historical and Monte Carlo VaR over a specified backtesting period.
* `main()`: The main function that orchestrates the entire process, including file upload, data retrieval, Excel updates, VaR calculation, plotting, and file download.

## Notes

* Ensure your 'Ticker' column in the 'Portfolio Summary' sheet contains valid Yahoo Finance tickers.
* The script assumes equal weighting for assets when calculating portfolio returns for VaR. You might need to modify this if you have different weighting requirements.
* The currency conversion relies on the availability of exchange rate data from `yfinance` (e.g., `EUR=X` for EUR to USD).
* The historical data download period is determined by the 'VaR range (in years)' in the 'Settings' sheet, plus the VaR backtesting period.
* The plot is displayed in the Colab output. If you need to save it locally, you can uncomment the `plt.savefig()` and relevant `files.download()` lines in the `main` function.

Feel free to contribute or modify the script to suit your specific portfolio analysis needs!
