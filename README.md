# Automated Stock Portfolio Tracker with Real-Time Data Integration

Welcome to the Automated Stock Portfolio Tracker with Real-Time Data Integration! This project aims to simplify the process of tracking stock investments by automatically fetching current stock prices and updating an Excel file with relevant data. It also provides profit and loss calculations for each stock in the portfolio and creates backups of the data.

## Features

- **Real-Time Stock Price Fetching**: Uses Yahoo Finance API to get the latest stock prices.
- **Excel File Updates**: Updates the Excel file (`User.xlsx`) with current stock prices.
- **Profit and Loss Calculation**: Calculates the profit or loss for each stock based on the current price and the buying price.
- **Data Backup**: Automatically creates backups of the Excel file to ensure data safety.

## Getting Started

### Prerequisites

- Python 3.x
- `openpyxl` library
- `yfinance` library

You can install the required libraries using pip:

```sh
pip install openpyxl yfinance
```

### Files

- **Update Current Price.py**: The main script that performs all operations.
- **User.xlsx**: The Excel file containing your stock portfolio.

### Setting Up

1. **Clone the Repository**:
    ```sh
    git clone https://github.com/kayteekay1412/Automated-Stock-Portfolio-Tracker-with-Real-Time-Data-Integration.git
    cd Automated-Stock-Portfolio-Tracker-with-Real-Time-Data-Integration
    ```

2. **Prepare Your Excel File**:
    - Ensure your `User.xlsx` file is in the `Portfolio Tracker` directory.
    - The file should have the following columns:
        - **A**: Stock Name
        - **B**: Quantity
        - **C**: Buying Price
        - **K**: Stock Symbol (used to fetch data from Yahoo Finance)
        - **E**: Current Price (this will be updated by the script)

3. **Run the Script**:
    ```sh
    python Update\ Current\ Price.py
    ```

## Detailed Description

The script performs the following steps:

1. **Read Stock Data**:
    - Reads stock names, symbols, quantities, and buying prices from the `User.xlsx` file.

2. **Fetch Current Prices**:
    - Uses the Yahoo Finance API to fetch the latest prices for each stock.

3. **Update Excel File**:
    - Writes the current prices back to the `User.xlsx` file.

4. **Calculate Profit or Loss**:
    - Computes the profit or loss for each stock based on the current price and the buying price.

5. **Create Backup**:
    - Creates a backup of the `User.xlsx` file in the `Records` directory with a timestamp.

## Contributing

Contributions are welcome! If you have any suggestions or improvements, please create an issue or submit a pull request.
Feel free to reach out if you have any questions or need further assistance.
