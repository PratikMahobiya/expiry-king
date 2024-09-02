import csv
import openpyxl
import requests
import pandas as pd
from tqdm import tqdm
from datetime import datetime, timedelta


# Screener Constants
fixed_target = 60           # 60 %
fixed_stoploss = 30         # 30 %
number_of_position = 10     # Infinite or fixed
wallet = 10000             # Wallet Balance
max_entry_amount = 100000   # Max entry Amount
entry_amount = 1000        # per entry
increase_percent = 5        # When profit is greater then 10% then only entry amount increased by 5%
fixed_entry_amount_flag = False
fixed_target_flag = False

file_name = 'V8'

def convert_to_ist(timestamp_ms):
    # Convert milliseconds to seconds
    timestamp_s = timestamp_ms / 1000
    # Convert to datetime object in UTC
    dt_utc = pd.to_datetime(timestamp_s, unit='s', utc=True)
    # IST offset
    ist_offset = timedelta(hours=5, minutes=30)
    # Convert to IST
    dt_ist = dt_utc + ist_offset
    return dt_ist

def Get_Candle(collection, interval, from_date, to_date):
    try:
        # url = "https://api.coindcx.com/exchange/v1/derivatives/futures/data/active_instruments"

        # response = requests.get(url)
        # data = response.json()
        crypto_all = ['B-CTSI_USDT', 'B-LPT_USDT', 'B-RLC_USDT', 'B-MKR_USDT', 'B-ENS_USDT', 'B-BTC_USDT', 'B-PEOPLE_USDT', 'B-ROSE_USDT', 'B-DUSK_USDT', 'B-FLOW_USDT', 'B-IMX_USDT', 'B-API3_USDT', 'B-GMT_USDT', 'B-APE_USDT', 'B-WOO_USDT', 'B-JASMY_USDT', 'B-DAR_USDT', 'B-OP_USDT', 'B-INJ_USDT', 'B-SPELL_USDT', 'B-LDO_USDT', 'B-ICP_USDT', 'B-APT_USDT', 'B-QNT_USDT', 'B-FXS_USDT', 'B-MAGIC_USDT', 'B-T_USDT', 'B-HIGH_USDT', 'B-MINA_USDT', 'B-ASTR_USDT', 'B-PHB_USDT', 'B-GMX_USDT', 'B-CFX_USDT', 'B-STX_USDT', 'B-BNX_USDT', 'B-ACH_USDT', 'B-CKB_USDT', 'B-PERP_USDT', 'B-TRU_USDT', 'B-LQTY_USDT', 'B-ARB_USDT', 'B-JOE_USDT', 'B-TLM_USDT', 'B-AMB_USDT', 'B-LEVER_USDT', 'B-HFT_USDT', 'B-XVS_USDT', 'B-BICO_USDT', 'B-LOOM_USDT', 'B-BOND_USDT', 'B-WAXP_USDT', 'B-RIF_USDT', 'B-POLYX_USDT', 'B-GAS_USDT', 'B-POWR_USDT', 'B-CAKE_USDT', 'B-TWT_USDT', 'B-BADGER_USDT', 'B-CHESS_USDT', 'B-SSV_USDT', 'B-ILV_USDT', 'B-ETH_USDT', 'B-BCH_USDT', 'B-XRP_USDT', 'B-EOS_USDT', 'B-LTC_USDT', 'B-ETC_USDT', 'B-LINK_USDT', 'B-TRX_USDT', 'B-XLM_USDT', 'B-ADA_USDT', 'B-DASH_USDT', 'B-ZEC_USDT', 'B-XTZ_USDT', 'B-BNB_USDT', 'B-ATOM_USDT', 'B-ONT_USDT', 'B-IOTA_USDT', 'B-BAT_USDT', 'B-VET_USDT', 'B-NEO_USDT', 'B-QTUM_USDT', 'B-IOST_USDT', 'B-THETA_USDT', 'B-ALGO_USDT', 'B-ZIL_USDT', 'B-KNC_USDT', 'B-ZRX_USDT', 'B-COMP_USDT', 'B-OMG_USDT', 'B-DOGE_USDT', 'B-SXP_USDT', 'B-KAVA_USDT', 'B-BAND_USDT', 'B-SNX_USDT', 'B-DOT_USDT', 'B-YFI_USDT', 'B-BAL_USDT', 'B-CRV_USDT', 'B-TRB_USDT', 'B-RUNE_USDT', 'B-SUSHI_USDT', 'B-EGLD_USDT', 'B-SOL_USDT', 'B-ICX_USDT', 'B-STORJ_USDT', 'B-BLZ_USDT', 'B-UNI_USDT', 'B-AVAX_USDT', 'B-FTM_USDT', 'B-ENJ_USDT', 'B-FLM_USDT', 'B-REN_USDT', 'B-KSM_USDT', 'B-NEAR_USDT', 'B-AAVE_USDT', 'B-FIL_USDT', 'B-LRC_USDT', 'B-MATIC_USDT', 'B-BEL_USDT', 'B-AXS_USDT', 'B-ALPHA_USDT', 'B-ZEN_USDT', 'B-SKL_USDT', 'B-GRT_USDT', 'B-1INCH_USDT', 'B-CHZ_USDT', 'B-SAND_USDT', 'B-ANKR_USDT', 'B-LIT_USDT', 'B-UNFI_USDT', 'B-REEF_USDT', 'B-RVN_USDT', 'B-SFP_USDT', 'B-XEM_USDT', 'B-COTI_USDT', 'B-STMX_USDT', 'B-CELR_USDT', 'B-HOT_USDT', 'B-MTL_USDT', 'B-OGN_USDT', 'B-NKN_USDT', 'B-BAKE_USDT', 'B-GTC_USDT', 'B-IOTX_USDT', 'B-C98_USDT', 'B-MASK_USDT', 'B-ATA_USDT', 'B-DYDX_USDT', 'B-GALA_USDT', 'B-CELO_USDT', 'B-AR_USDT', 'B-ARPA_USDT', 'B-YGG_USDT', 'B-BNT_USDT', 'B-OXT_USDT', 'B-CHR_USDT', 'B-SUPER_USDT', 'B-USTC_USDT', 'B-ONG_USDT', 'B-AUCTION_USDT', 'B-MOVR_USDT', 'B-LSK_USDT', 'B-OM_USDT', 'B-GLM_USDT', 'B-RARE_USDT', 'B-SYN_USDT', 'B-SYS_USDT', 'B-NULS_USDT', 'B-MANA_USDT', 'B-ALICE_USDT', 'B-HBAR_USDT', 'B-DENT_USDT', 'B-ONE_USDT', 'B-LINA_USDT', 'B-KLAY_USDT', 'B-UMA_USDT', 'B-KEY_USDT', 'B-VIDT_USDT', 'B-MBOX_USDT', 'B-NMR_USDT', 'B-HIFI_USDT', 'B-ALPACA_USDT', 'B-SUN_USDT']

        crypto_100 = ['B-LPT_USDT', 'B-MKR_USDT', 'B-ENS_USDT', 'B-BTC_USDT', 'B-FLOW_USDT', 'B-API3_USDT', 'B-APE_USDT', 'B-JASMY_USDT', 'B-OP_USDT', 'B-INJ_USDT', 'B-LDO_USDT', 'B-APT_USDT', 'B-MINA_USDT', 'B-CFX_USDT', 'B-STX_USDT', 'B-ACH_USDT', 'B-LQTY_USDT', 'B-ARB_USDT', 'B-JOE_USDT', 'B-LEVER_USDT', 'B-ETH_USDT', 'B-BCH_USDT', 'B-XRP_USDT', 'B-EOS_USDT', 'B-LTC_USDT', 'B-ETC_USDT', 'B-LINK_USDT', 'B-TRX_USDT', 'B-XLM_USDT', 'B-ADA_USDT', 'B-XTZ_USDT', 'B-BNB_USDT', 'B-ATOM_USDT', 'B-QTUM_USDT', 'B-KNC_USDT', 'B-DOGE_USDT', 'B-SXP_USDT', 'B-BAND_USDT', 'B-SNX_USDT', 'B-DOT_USDT', 'B-YFI_USDT', 'B-CRV_USDT', 'B-TRB_USDT', 'B-RUNE_USDT', 'B-SOL_USDT', 'B-ICX_USDT', 'B-STORJ_USDT', 'B-BLZ_USDT', 'B-UNI_USDT', 'B-AVAX_USDT', 'B-FTM_USDT', 'B-FLM_USDT', 'B-NEAR_USDT', 'B-AAVE_USDT', 'B-FIL_USDT', 'B-LRC_USDT', 'B-MATIC_USDT', 'B-BEL_USDT', 'B-1INCH_USDT', 'B-CHZ_USDT', 'B-SAND_USDT', 'B-ANKR_USDT', 'B-UNFI_USDT', 'B-SFP_USDT', 'B-STMX_USDT', 'B-MTL_USDT', 'B-OGN_USDT', 'B-BAKE_USDT', 'B-GTC_USDT', 'B-IOTX_USDT', 'B-C98_USDT', 'B-DYDX_USDT', 'B-GALA_USDT', 'B-CELO_USDT', 'B-ARPA_USDT', 'B-MANA_USDT', 'B-HBAR_USDT', 'B-LINA_USDT', 'B-KLAY_USDT', 'B-KEY_USDT', 'B-NMR_USDT']
        
        data = crypto_100 if collection == '100' else crypto_all

        print("Total symbol: ", len(data))
        symbol = []
        df = []
        for index, pair in enumerate(tqdm(data)):
            # url = "https://public.coindcx.com/market_data/candles"
            # query_params = {
            #     "interval": interval,
            #     "pair": pair,
            #     "startTime": from_date.timestamp() * 1000,
            #     "endTime": to_date.timestamp() * 1000,
            #     "limit": 1000
            # }
            url = "https://public.coindcx.com/market_data/candlesticks"
            query_params = {
                "pair": pair,
                "from": from_date.timestamp(),
                "to": to_date.timestamp(),
                "resolution": interval,  # '1' OR '5' OR '60' OR '1D'
                "pcode": "f"
            }
            response = requests.get(url, params=query_params)
            data = response.json()
            data = data['data']
            # if len(data) > 364:
            symbol.append(pair[2:-1].replace('_', '-'))
            # Apply the conversion to the DataFrame
            data_frame = pd.DataFrame(data[::-1])
            data_frame['time'] = data_frame['time'].apply(convert_to_ist)
            data_frame.rename(columns={
                            'open': 'Open',
                            'high': 'High',
                            'low': 'Low',
                            'close': 'Close',
                            'volume': 'Volume',
                            'time': 'Time'}, inplace=True)
            df.append(data_frame)

        mdf = pd.concat([ i.set_index('Time') for i in df], axis=1, keys=symbol)
    except Exception as e:
        raise Exception(f"Error: {e}")
    return mdf, symbol

def get_change(current, previous):
    if current == previous:
        return 0
    try:
        return (abs(current - previous) / previous) * 100.0
    except ZeroDivisionError:
        return float('inf')


def Entry(date_time, data_frame, symbol, active_entry, wallet, entry_amount, sheet_data):
    fixed_target_price = data_frame['Close'] + data_frame['Close']*fixed_target/100
    fixed_stoploss_price = data_frame['Close'] - data_frame['Close']*fixed_stoploss/100

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

    if data_frame['High'] > active_entry[symbol]['fixed_target'] and fixed_target_flag:
        sell_price = active_entry[symbol]['fixed_target']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = (price_diff/active_entry[symbol]['price']) * 100
        days = (date_time - active_entry[symbol]['datetime']).days
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl > 10 and entry_amount <= max_entry_amount:
                if wallet > (entry_amount + entry_amount * increase_percent/100)*3:
                    entry_amount = entry_amount + entry_amount * increase_percent/100

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Target', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]
    
    elif active_entry[symbol]['tr_sl'] and data_frame['Low'] < active_entry[symbol]['tr_stoploss']:
        sell_price = active_entry[symbol]['tr_stoploss']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = (price_diff/active_entry[symbol]['price']) * 100
        days = (date_time - active_entry[symbol]['datetime']).days
        gained_amount = active_entry[symbol]['shares'] * sell_price
        actual_amount = gained_amount - active_entry[symbol]['invested_amount']
        wallet = wallet + gained_amount
        if not fixed_entry_amount_flag:
            if pnl > 10 and entry_amount <= max_entry_amount:
                if wallet > (entry_amount + entry_amount * increase_percent/100)*3:
                    entry_amount = entry_amount + entry_amount * increase_percent/100

        sht_data = [date_time.year, date_time.month, date_time, symbol, 'Exit', 'Tr-Sl', sell_price, active_entry[symbol]['fixed_target'], active_entry[symbol]['fixed_stoploss'], active_entry[symbol]['tr_stoploss'], price_diff, pnl, active_entry[symbol]['max_high'], active_entry[symbol]['max_low'], days, active_entry[symbol]['shares'], active_entry[symbol]['invested_amount'], gained_amount, actual_amount, len(active_entry) - 1, wallet, entry_amount]
        sheet_data.append(sht_data)
        del active_entry[symbol]
    
    elif data_frame['Low'] < active_entry[symbol]['fixed_stoploss']:
        sell_price = active_entry[symbol]['fixed_stoploss']
        price_diff = sell_price - active_entry[symbol]['price']
        pnl = (price_diff/active_entry[symbol]['price']) * 100
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

multiple_data_frame, symbol_list = Get_Candle('100', interval='5', from_date=datetime(2024, 3, 28), to_date=datetime(2024, 3, 31))

for index, date_time in enumerate(tqdm(multiple_data_frame.index)):
    date_time = date_time.replace(tzinfo=None)
    for symbol in symbol_list:
        if len(active_entry) > number_of_entry_at_a_time:
            number_of_entry_at_a_time = len(active_entry)

        # Take Entry
        if index > 50 and not active_entry.get(symbol) and len(active_entry) < number_of_position and max(multiple_data_frame[symbol]['High'].iloc[index-50:index]) < multiple_data_frame.iloc[index][symbol]['High']:
            wallet, entry_amount, active_entry = Entry(date_time, multiple_data_frame.iloc[index][symbol], symbol, active_entry, wallet, entry_amount, sheet_data)

        # Take Exit
        elif active_entry.get(symbol) and active_entry.get(symbol).get('buy'):
            high_price_diff = multiple_data_frame.iloc[index][symbol]['High'] - active_entry[symbol]['price']
            high_pnl = (high_price_diff/active_entry[symbol]['price']) * 100
            low_price_diff = multiple_data_frame.iloc[index][symbol]['Low'] - active_entry[symbol]['price']
            low_pnl = (low_price_diff/active_entry[symbol]['price']) * 100
            close_price_diff = multiple_data_frame.iloc[index][symbol]['Close'] - active_entry[symbol]['price']
            close_pnl = (close_price_diff/active_entry[symbol]['price']) * 100
            active_entry[symbol]['change_in_investment'] = close_pnl
            
            if high_pnl < fixed_target and multiple_data_frame.iloc[index][symbol]['Low'] > active_entry[symbol]['fixed_stoploss']:
                active_entry[symbol]['tr_sl'] = True
                active_entry[symbol]['tr_stoploss'] = min(multiple_data_frame[symbol]['Low'].iloc[index-10:index])
            elif high_pnl > fixed_target and high_pnl > active_entry[symbol]['max_high']:
                active_entry[symbol]['tr_sl'] = True
                active_entry[symbol]['tr_stoploss'] = multiple_data_frame.iloc[index][symbol]['High'] - multiple_data_frame.iloc[index][symbol]['High'] * 0.05
            
            if high_pnl > active_entry[symbol]['max_high']:
                active_entry[symbol]['max_high'] = high_pnl
            
            if low_pnl < active_entry[symbol]['max_low']:
                active_entry[symbol]['max_low'] = low_pnl

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
stats_sheet_data.append(['Gained Capital', wallet])
stats_sheet_data.append(['Still Invested Amount', amount_invested])
stats_sheet_data.append(['Total Wallet Amount', wallet+amount_invested])
stats_sheet_data.append(['Max Portfolio Change %', max_portfolio_change])
stats_sheet_data.append(['Min Portfolio Change %', min_portfolio_change])

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
    stats_sheet_data.append(['Total Profit/Loss(%)', get_change(wallet+amount_invested, initial_wallet)])
    stats_sheet_data.append(['Avg Profit/Loss(%)', sum(winners+losers)/len(winners+losers)])
    stats_sheet_data.append(['Total Entry', total_entry])
    stats_sheet_data.append(['Total Exit', total_exit])
    stats_sheet_data.append(['Total Active Entries', total_entry - total_exit])
    stats_sheet_data.append(['Total Number of Win', total_number_of_win])
    stats_sheet_data.append(['Total Win %', (total_number_of_win/total_exit)*100])
    stats_sheet_data.append(['Total Number of Loss', total_number_of_loss])
    stats_sheet_data.append(['Total Loss %', (total_number_of_loss/total_exit)*100])
    stats_sheet_data.append(['Total Number of Concutive Win', max(consecutive_win)])
    stats_sheet_data.append(['Total Number of Concutive Loss', max(consecutive_loss)])
    stats_sheet_data.append(['Total Number of Target', total_number_of_targets])
    stats_sheet_data.append(['Total Number of Tr-Sl', total_number_of_trsl])
    stats_sheet_data.append(['Total Number of Stoploss', total_number_of_stoploss])
    stats_sheet_data.append([' ', ' '])

    stats_sheet_data.append(['Winners', ' '])
    stats_sheet_data.append(['Total Win', sum(winners)])
    stats_sheet_data.append(['Avg Win', sum(winners)/len(winners)])
    stats_sheet_data.append(['Max Win', max(winners)])
    stats_sheet_data.append(['Min Win', min(winners)])
    stats_sheet_data.append([' ', ' '])
    
    stats_sheet_data.append(['Losers', ' '])
    stats_sheet_data.append(['Total Loss', sum(losers)])
    stats_sheet_data.append(['Avg Loss', sum(losers)/len(losers)])
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
