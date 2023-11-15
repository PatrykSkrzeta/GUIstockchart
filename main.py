import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoLocator, AutoMinorLocator
from datetime import datetime, timedelta
import customtkinter as ctk

def format_yaxis(y, pos):
    billions = y / 1e9
    millions = y / 1e6
    units = ""
    
    if billions >= 1:
        units = f'{billions:.0f}B'
    elif millions >= 1:
        units = f'{millions:.0f}M'
    else:
        units = f'{y:.0f}'
    
    return units

def plot_stock_price(ticker, year_count, save_to_excel):
    end_date = datetime.today()
    start_date = end_date - timedelta(days=year_count * 365)

    close_df = pd.DataFrame()

    data = yf.download(ticker, start=start_date, end=end_date)
    close_df[ticker] = data['Close']

    if save_to_excel:
        excel_filename = "stock_price_report.xlsx"
        close_df.to_excel(excel_filename, sheet_name='Stock Prices', index=True)
        print(f'DataFrame saved to {excel_filename}')
    else:
        print('Data not saved to Excel.')

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(close_df.index, close_df[ticker], label=ticker)
    ax.set_title(f'Price Chart for {ticker}')
    ax.set_xlabel('Date')
    ax.set_ylabel('Closing Price')
    ax.yaxis.set_major_locator(AutoLocator())
    ax.yaxis.set_minor_locator(AutoMinorLocator())
    ax.yaxis.set_major_formatter(format_yaxis)

    ax.legend()
    plt.show()

    root.destroy()

def on_button_click():
    ticker = entry_ticker.get()
    year_count = int(entry_years.get())
    save_to_excel = entry_save_to_excel.get().lower()

    plot_stock_price(ticker, year_count, save_to_excel)


ctk.set_default_color_theme("dark-blue")
ctk.set_appearance_mode("dark")


root = ctk.CTk()
root.geometry("600x350")
root.title("Stock Chart Generator by Patryk Skrzęta")

frame = ctk.CTkFrame(root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

ctk.CTkLabel(frame, text="Enter the symbol (e.g., AAPL or BTC-USD):", font=("Arial", 15)).pack()
entry_ticker = ctk.CTkEntry(frame)
entry_ticker.pack(pady=5, padx=10)

ctk.CTkLabel(frame, text="Enter years count:", font=("Arial", 15)).pack()
entry_years = ctk.CTkEntry(frame)
entry_years.pack(pady=5, padx=10)

ctk.CTkLabel(frame, text="Do you want to save the data to an Excel file? (yes/no):", font=("Arial", 15)).pack()
entry_save_to_excel = ctk.CTkEntry(frame)
entry_save_to_excel.pack(pady=5, padx=10)

button_plot = ctk.CTkButton(frame, text="Generate Chart", command=on_button_click)
button_plot.pack(pady=25, padx=10)

root.mainloop()
