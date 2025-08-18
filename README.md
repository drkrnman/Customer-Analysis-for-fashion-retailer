# Data-analisys-project---Global-fascion-retail-sales

## Project Goal
The goal of this project is to determine the difference in customer profitability across different customer segments.  
We use **LTV (LifeTime Value) for 6 months** — the average revenue per customer during the first 6 months since their first purchase.

## Dataset
The dataset is provided and contains information about:
- Customers
- Transactions
- Products
- Stores
- Discounts
- Employees

CSV files included:
- `customers.csv`
- `transactions.csv`
- `products.csv`
- `stores.csv`
- `discounts.csv`
- `employees.csv`
- `customer_stats.csv` (calculated metrics)
- `purchases_final.csv` (processed dataset)

## User’s CLI
The project has a command-line interface (CLI) that allows users to:
1. Select a report or analysis type from a menu.
2. View aggregated metrics and visualizations for different customer segments.

*The dataset is preloaded, so the user does not need to provide any CSV files.*

## Installation
1. Clone the repository.
2. Install dependencies using: `pip install -r requirements.txt`.
3. Run the project: `python main.py`.

## Requirements
All required Python packages are listed in `requirements.txt`.

## Notes
- Python 3.10+ is recommended.
- Large CSV files are tracked using Git LFS to ensure compatibility with GitHub.

## Author
Darya Korenman
