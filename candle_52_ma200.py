import csv
import openpyxl
import pandas as pd
import yfinance as yf
from tqdm import tqdm
from niftystocks import ns


# Screener Constants
fixed_target = 60           # 60 %
fixed_stoploss = 15         # 30 %
number_of_position = 20     # Infinite or fixed
wallet = 200000             # Wallet Balance
max_entry_amount = 100000   # Max entry Amount
entry_amount = 10000        # per entry
increase_percent = 5        # When profit is greater then 10% then only entry amount increased by 5%
fixed_entry_amount_flag = False
fixed_target_flag = True

file_name = 'v15_n'
symbol_list_unfiltered = ns.get_nifty200_with_ns()


exclude_symbol = ['MAHINDCIE.NS', 'ORIENTREF.NS', 'PVR.NS', 'WABCOINDIA.NS', 'SRTRANSFIN.NS', 'LTI.NS', 'L&TFH.NS', 'MINDAIND.NS', 'CADILAHC.NS', 'IIFLWAM.NS', 'MOTHERSUMI.NS', 'BURGERKING.NS', 'SUNCLAYLTD.NS', 'SHRIRAMCIT.NS', 'ANGELBRKG.NS', 'WELSPUNIND.NS', 'KALPATPOWR.NS', 'AMARAJABAT.NS', 'HDFC.NS', 'SUPPETRO.NS', 'ADANITRANS.NS', 'PHILIPCARB.NS', 'MINDTREE.NS', 'UJJIVAN.NS', 'TATACOFFEE.NS', 'GODREJCP.NS', 'MCDOWELL-N.NS', 'AEGISCHEM.NS']

symbol_list = [symbol for symbol in symbol_list_unfiltered if symbol not in exclude_symbol]

multiple_data_frame = yf.download(symbol_list, interval="1d", start='1999-04-01', end='2024-03-31', group_by='ticker', rounding=True)

def get_change(current, previous):
    if current == previous:
        return 0
    try:
        return (abs(current - previous) / previous) * 100.0
    except ZeroDivisionError:
        return float('inf')


def Entry(date_time, data_frame, symbol, active_entry, wallet, entry_amount, sheet_data):
    fixed_target_price = round(data_frame['Close'] + data_frame['Close']*fixed_target/100, 2)
    fixed_stoploss_price = round(data_frame['Close'] - data_frame['Close']*fixed_stoploss/100, 2)

    if entry_amount > data_frame['Close'] and wallet > entry_amount:
        invested_amount = 0
        shares = 0
        while True:
            invested_amount += data_frame['Close']
            shares += 1
            if invested_amount > entry_amount:
                invested_amount -= data_frame['Close']
                shares -= 1
                break
        wallet = wallet - invested_amount
        active_entry[symbol] = {
            'buy': True,
            'tr_sl': False,
            'fixed_target': fixed_target_price,
            'fixed_stoploss': fixed_stoploss_price,
            'tr_stoploss': 0,
            'price': data_frame['Close'],
            'datetime': date_time,
            'max_high': 0,
            'max_low': 0,
            'invested_amount': invested_amount,
            'shares': shares,
            'change_in_investment': 0,
        }
        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Entry', 'Buy', active_entry[symbol]['price'], active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], '', '', '', '', '', active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], '', '', len(active_entry), wallet, entry_amount]
        sheet_data.append(sht_data)
    return wallet, entry_amount, active_entry


def Exit(date_time, data_frame, symbol, active_entry, wallet, entry_amount, sheet_data):

    
    # gapdown
    if data_frame['Open'] < active_entry[symbol]['fixed_stoploss']:
        sell_price = data_frame['Open']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = round((price_diff/active_entry[symbol]['price']) * 100, 2)
        days = (date_time - active_entry[symbol]['datetime']).days
        active_entry[symbol]['max_low'] = pnl
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl < 0:
                entry_amount = entry_amount - entry_amount * increase_percent/100
        

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Stoploss', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]

    # gapup
    elif data_frame['Open'] > active_entry[symbol]['fixed_target'] and fixed_target_flag:
        sell_price = data_frame['Open']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = round((price_diff/active_entry[symbol]['price']) * 100, 2)
        days = (date_time - active_entry[symbol]['datetime']).days
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl > 10 and entry_amount <= max_entry_amount:
                if wallet > entry_amount + entry_amount * increase_percent/100:
                    entry_amount = entry_amount + entry_amount * increase_percent/100

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Target', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]

    elif data_frame['High'] > active_entry[symbol]['fixed_target'] and fixed_target_flag:
        sell_price = active_entry[symbol]['fixed_target']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = round((price_diff/active_entry[symbol]['price']) * 100, 2)
        days = (date_time - active_entry[symbol]['datetime']).days
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl > 10 and entry_amount <= max_entry_amount:
                if wallet > entry_amount + entry_amount * increase_percent/100:
                    entry_amount = entry_amount + entry_amount * increase_percent/100

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Target', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]
    
    elif active_entry[symbol]['tr_sl'] and data_frame['Low'] < active_entry[symbol]['tr_stoploss']:
        sell_price = active_entry[symbol]['tr_stoploss']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = round((price_diff/active_entry[symbol]['price']) * 100, 2)
        days = (date_time - active_entry[symbol]['datetime']).days
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl > 10 and entry_amount <= max_entry_amount:
                if wallet > entry_amount + entry_amount * increase_percent/100:
                    entry_amount = entry_amount + entry_amount * increase_percent/100

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Tr-Sl', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]
    
    elif data_frame['Low'] < active_entry[symbol]['fixed_stoploss']:
        sell_price = active_entry[symbol]['fixed_stoploss']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = round((price_diff/active_entry[symbol]['price']) * 100, 2)
        days = (date_time - active_entry[symbol]['datetime']).days
        active_entry[symbol]['max_low'] = pnl
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl < 0:
                entry_amount = entry_amount - entry_amount * increase_percent/100

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Stoploss', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]
    return wallet, entry_amount, active_entry


# Screener Variable
sheet_data = [
    ['Year', 'Month', 'Datetime', 'Symbol', 'Indicate', 'Type', 'Price', 'Target', 'Stoploss', 'Tr-Stoploss', 'PriceDifference', 'PnL(%)', 'Max High(%)', 'Max Low(%)', 'Entry-Stays(days)', 'No of shares', 'Invested AMT', 'Gained AMT', 'Actual Gain', 'Current Open Entry', 'Wallet', 'Entry Amount']
]
# Screener Variable
stats_sheet_data = [
    ['Stats', 'Values'],
    [' ', ' '],
    ['Initial Entry Amount', entry_amount],
    ['Initial Capital', wallet],
]
initial_wallet = wallet
active_entry = {}
flag_trsl = {}
high_dict = {}
number_of_entry_at_a_time = 0
max_portfolio_change = 0
min_portfolio_change = 0

for index, date_time in enumerate(tqdm(multiple_data_frame.index)):
    for symbol in symbol_list:
        if len(active_entry) > number_of_entry_at_a_time:
            number_of_entry_at_a_time = len(active_entry)

        # Take Entry
        if index > 200 and not active_entry.get(symbol) and len(active_entry) < number_of_position and max(multiple_data_frame[symbol]['High'].iloc[index-52:index]) < multiple_data_frame.iloc[index][symbol]['High'] and (sum(multiple_data_frame[symbol]['High'].iloc[index-200:index])/200) < multiple_data_frame.iloc[index][symbol]['High']:
            wallet, entry_amount, active_entry = Entry(date_time, multiple_data_frame.iloc[index][symbol], symbol, active_entry, wallet, entry_amount, sheet_data)           
        
        # Take Exit
        elif active_entry.get(symbol) and active_entry.get(symbol).get('buy'):
            high_price_diff = multiple_data_frame.iloc[index][symbol]['High'] - active_entry[symbol]['price']
            high_pnl = round((high_price_diff/active_entry[symbol]['price']) * 100, 2)
            low_price_diff = multiple_data_frame.iloc[index][symbol]['Low'] - active_entry[symbol]['price']
            low_pnl = round((low_price_diff/active_entry[symbol]['price']) * 100, 2)
            close_price_diff = multiple_data_frame.iloc[index][symbol]['Close'] - active_entry[symbol]['price']
            close_pnl = round((close_price_diff/active_entry[symbol]['price']) * 100, 2)
            active_entry[symbol]['change_in_investment'] = close_pnl

            if high_pnl > active_entry[symbol]['max_high']:
                active_entry[symbol]['max_high'] = high_pnl
            
            if low_pnl < active_entry[symbol]['max_low']:
                active_entry[symbol]['max_low'] = low_pnl
            
            if high_pnl < fixed_target and multiple_data_frame.iloc[index][symbol]['Low'] > active_entry[symbol]['fixed_stoploss']:
                active_entry[symbol]['tr_sl'] = True
                active_entry[symbol]['tr_stoploss'] = min(multiple_data_frame[symbol]['Low'].iloc[index-6:index])
            elif high_pnl > fixed_target:
                active_entry[symbol]['tr_sl'] = True
                active_entry[symbol]['tr_stoploss'] = multiple_data_frame.iloc[index][symbol]['High'] - multiple_data_frame.iloc[index][symbol]['High'] * 0.05
            wallet, entry_amount, active_entry = Exit(date_time, multiple_data_frame.iloc[index][symbol], symbol, active_entry, wallet, entry_amount, sheet_data)
    
    # Get Portfolio Value change
    change = 0
    for i in active_entry:
        change += active_entry[i]['change_in_investment']
    
    if change > max_portfolio_change:
        max_portfolio_change = change
    if change < min_portfolio_change:
        min_portfolio_change = change


amount_invested = 0
for i in active_entry:
    amount_invested = amount_invested + active_entry[i]['invested_amount']

stats_sheet_data.append(['Changed to Entry Amount', entry_amount])
stats_sheet_data.append(['Gained Capital', round(wallet, 2)])
stats_sheet_data.append(['Still Invested Amount', round(amount_invested, 2)])
stats_sheet_data.append(['Total Wallet Amount', round(wallet+amount_invested, 2)])
stats_sheet_data.append(['Max Portfolio Change %', round(max_portfolio_change, 2)])
stats_sheet_data.append(['Min Portfolio Change %', round(min_portfolio_change, 2)])

# Create a new Excel workbook
workbook = openpyxl.Workbook()
# Select the default sheet (usually named 'Sheet')
sheet = workbook.active

for row in sheet_data:
    sheet.append(row)
# Save the workbook to a file
workbook.save(f"{file_name}.xlsx")
# Print a success message
print("Excel file created successfully!")


# Read and store content 
# of an excel file  
read_file = pd.read_excel(f"{file_name}.xlsx") 
  
# Write the dataframe object 
# into csv file 
read_file.to_csv (f"{file_name}.csv", index = None, header=True)

def trim_data(data_):
    """
        Remove Whitespaces
    """
    for data in data_:
        if data_[data]:
            data_[data] = data_[data].strip()
    return data_

with open(f'{file_name}.csv', mode='r') as csv_file:

    # Variables
    winners = []
    losers = []
    num_consecutive_win = 0
    num_consecutive_loss = 0
    consecutive_win = []
    consecutive_loss = []
    total_entry = 0
    total_exit = 0
    total_number_of_win = 0
    total_number_of_loss = 0
    total_number_of_targets = 0
    total_number_of_stoploss = 0
    total_number_of_trsl = 0
    entry_stays_days_bars = 0
    
    csv_reader = csv.DictReader(csv_file)
    for row_data in csv_reader:

        # Removing Whitespaces from start and end
        row = trim_data(row_data)

        if row_data['Indicate'] == 'Entry':
            total_entry += 1
        elif row_data['Indicate'] == 'Exit':
            if float(row['Entry-Stays(days)']) > entry_stays_days_bars:
                entry_stays_days_bars = float(row['Entry-Stays(days)'])
            total_exit += 1


        if row_data['PnL(%)'] not in ['', ' ', None]:

            if float(row_data['PnL(%)']) > 0:
                total_number_of_win += 1
                winners.append(float(row_data['PnL(%)']))
                
                if num_consecutive_loss != 0:
                    consecutive_loss.append(num_consecutive_loss)
                    num_consecutive_loss = 0
                num_consecutive_win += 1
                
                if row_data['Type'] == 'Target':
                    total_number_of_targets += 1
                if row_data['Type'] == 'Tr-Sl':
                    total_number_of_trsl += 1
            
            elif float(row_data['PnL(%)']) < 0:
                total_number_of_loss += 1
                losers.append(float(row_data['PnL(%)']))

                if num_consecutive_win != 0:
                    consecutive_win.append(num_consecutive_win)
                    num_consecutive_win = 0
                num_consecutive_loss += 1

                if row_data['Type'] == 'Tr-Sl':
                    total_number_of_trsl += 1
                if row_data['Type'] == 'Stoploss':
                    total_number_of_stoploss += 1
        

    stats_sheet_data.append([' ', ' '])
    stats_sheet_data.append(['All Trade', ' '])
    stats_sheet_data.append(['Total Number of Entry at a Time', number_of_entry_at_a_time])
    stats_sheet_data.append(['Entry Stays(Days/Bars)', entry_stays_days_bars])
    stats_sheet_data.append(['Total Profit/Loss(%)', round(get_change(wallet+amount_invested, initial_wallet), 2)])
    stats_sheet_data.append(['Avg Profit/Loss(%)', round(sum(winners+losers)/len(winners+losers), 2)])
    stats_sheet_data.append(['Total Entry', total_entry])
    stats_sheet_data.append(['Total Exit', total_exit])
    stats_sheet_data.append(['Total Active Entries', total_entry - total_exit])
    stats_sheet_data.append(['Total Number of Win', total_number_of_win])
    stats_sheet_data.append(['Total Win %', round((total_number_of_win/total_exit)*100, 2)])
    stats_sheet_data.append(['Total Number of Loss', total_number_of_loss])
    stats_sheet_data.append(['Total Loss %', round((total_number_of_loss/total_exit)*100, 2)])
    stats_sheet_data.append(['Total Number of Concutive Win', max(consecutive_win)])
    stats_sheet_data.append(['Total Number of Concutive Loss', max(consecutive_loss)])
    stats_sheet_data.append(['Total Number of Target', total_number_of_targets])
    stats_sheet_data.append(['Total Number of Tr-Sl', total_number_of_trsl])
    stats_sheet_data.append(['Total Number of Stoploss', total_number_of_stoploss])
    stats_sheet_data.append([' ', ' '])

    stats_sheet_data.append(['Winners', ' '])
    stats_sheet_data.append(['Total Win', round(sum(winners), 2)])
    stats_sheet_data.append(['Avg Win', round(sum(winners)/len(winners), 2)])
    stats_sheet_data.append(['Max Win', max(winners)])
    stats_sheet_data.append(['Min Win', min(winners)])
    stats_sheet_data.append([' ', ' '])
    
    stats_sheet_data.append(['Losers', ' '])
    stats_sheet_data.append(['Total Loss', round(sum(losers), 2)])
    stats_sheet_data.append(['Avg Loss', round(sum(losers)/len(losers), 2)])
    stats_sheet_data.append(['Max Loss', min(losers)])
    stats_sheet_data.append(['Min Loss', max(losers)])
    stats_sheet_data.append([' ', ' '])


# Create a new Excel workbook
workbook = openpyxl.Workbook()
# Select the default sheet (usually named 'Sheet')
sheet = workbook.active

for row in stats_sheet_data:
    sheet.append(row)
# Save the workbook to a file
workbook.save(f"stats_{file_name}.xlsx")
# Print a success message
print("Statistics Excel file created successfully!")
