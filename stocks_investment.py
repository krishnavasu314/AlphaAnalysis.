import pandas as pd
import yfinance as yf

def clean_data(file_path):
    """Load and clean the CSV file."""
    
    stocks_data = pd.read_csv(file_path, usecols=['Ticker', 'Weightage'])
    
    # Remove whitespace and drop missing rows
    stocks_data['Ticker'] = stocks_data['Ticker'].str.strip() + '.NS'  # Append `.NS`
    stocks_data.dropna(inplace=True)

    return stocks_data

def process_stocks(stocks_data, start_date, end_date, investment_amount):
    """Fetch stock data and calculate investment details."""
    results = []

    for _, row in stocks_data.iterrows():
        ticker = row['Ticker']
        weightage = row['Weightage']

        try:
            # Fetch stock data from Yahoo Finance
            stock_data = yf.download(ticker, start=start_date, end=end_date, progress=False)

            if stock_data.empty:
                print(f"No data found for {ticker}")
                continue

            # Calculate investment 
            daily_closing_prices = stock_data['Close']
            amount_to_invest = investment_amount * weightage
            num_shares = amount_to_invest / daily_closing_prices

            # Save results for this stock
            for date, price in daily_closing_prices.items():
                results.append({
                    'Date': date,
                    'Ticker': ticker,
                    'Closing Price': price,
                    'Weightage': weightage,
                    'Investment Amount': amount_to_invest,
                    'Number of Shares': num_shares[date]
                })

        except Exception as e:
            print(f"Error fetching data for {ticker}: {e}")

    return results

def save_to_excel(results, output_file):
    """Save the results to an Excel file."""
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

if __name__ == "__main__":
   
    file_path = "Stocks.csv"

  
    investment_amount = float(input("Enter the investment amount: "))
    start_date = input("Enter the start date (YYYY-MM-DD): ")
    end_date = input("Enter the end date (YYYY-MM-DD): ")


    stocks_data = clean_data(file_path)


    stock_results = process_stocks(stocks_data, start_date, end_date, investment_amount)

    
    output_file = "investment_results.xlsx"
    save_to_excel(stock_results, output_file)
