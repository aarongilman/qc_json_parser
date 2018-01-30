import json
import pandas as pd

##Put the json file exported from QuantConnect in the same directory as this file. Change json filename below to match
##export name from QuantConnect you placed in directory. Modify the strategy name below as it will ultimately be the
##filename of the excel workbook we are creating for analysis.

##Where are you storing the json files to read in?
json_directory = 'input'
##Where do you want the excel files to save to?
output_directory = 'output'

##Put the name of the json file below that is stored in the input directory.
json_filename = 'strategy1.json'
##Put the name of the strategy or whatever you want the excel output to be named.
strategy_name = 'Strategy1'

##Load the json file and read into python
with open(json_directory + "/" + json_filename) as json_data:
    d = json.load(json_data)
    print(d)

##Parse the json file and grab the elements we want to save to a workbook
orders = d['Orders']
total_performance = d['TotalPerformance']
trade_statistics = total_performance['TradeStatistics']
closed_trades = total_performance['ClosedTrades']
portfolio_statistics = total_performance['PortfolioStatistics']
statistics = d['Statistics']
profit_loss = d['ProfitLoss']

##Take the parsed json and save as pandas dataframes, which ultimately save as individual worksheets in the workbook
orders_df = pd.DataFrame.from_dict(orders, orient='index')
trade_statistics_df = pd.DataFrame.from_dict(trade_statistics, orient='index')
closed_trades_df = pd.DataFrame.from_records(closed_trades)
portfolio_statistics_df = pd.DataFrame.from_dict(portfolio_statistics, orient='index')
statistics_df = pd.DataFrame.from_dict(statistics, orient='index')
profit_loss_df = pd.DataFrame.from_dict(profit_loss, orient='index')

writer = pd.ExcelWriter(output_directory + "/" + strategy_name + '.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
orders_df.to_excel(writer, sheet_name='Orders')
trade_statistics_df.to_excel(writer, sheet_name='TradeStatistics')
closed_trades_df.to_excel(writer, sheet_name='ClosedTrades')
portfolio_statistics_df.to_excel(writer, sheet_name='PortfolioStatistics')
statistics_df.to_excel(writer, sheet_name='Statistics')
profit_loss_df.to_excel(writer, sheet_name='ProfitLoss')
# Close the Pandas Excel writer and output the Excel file.

rolling_window = d['RollingWindow']

windows = rolling_window.keys()

index_list = ['AverageEndTradeDrawdown',
              'AverageLosingTradeDuration',
              'AverageLoss',
              'AverageMAE',
              'AverageMFE',
              'AverageProfit',
              'AverageProfitLoss',
              'AverageTradeDuration',
              'AverageWinningTradeDuration',
              'EndDateTime',
              'LargestLoss',
              'LargestMAE',
              'LargestMFE',
              'LargestProfit',
              'LossRate',
              'MaxConsecutiveLosingTrades',
              'MaxConsecutiveWinningTrades',
              'MaximumClosedTradeDrawdown',
              'MaximumDrawdownDuration',
              'MaximumEndTradeDrawdown',
              'MaximumIntraTradeDrawdown',
              'NumberOfLosingTrades',
              'NumberOfWinningTrades',
              'ProfitFactor',
              'ProfitLossDownsideDeviation',
              'ProfitLossRatio',
              'ProfitLossStandardDeviation',
              'ProfitToMaxDrawdownRatio',
              'SharpeRatio',
              'SortinoRatio',
              'StartDateTime',
              'TotalFees',
              'TotalLoss',
              'TotalNumberOfTrades',
              'TotalProfit',
              'TotalProfitLoss',
              'WinLossRatio',
              'WinRate']

trades_dataframe = pd.DataFrame(index=index_list, columns=windows)

for window in windows:
    new_window = rolling_window[window]
    trade_stats = new_window['TradeStatistics']
    trades_dataframe[window] = pd.DataFrame.from_dict(trade_stats, orient='index')

trades_dataframe = trades_dataframe.T
trades_dataframe.to_excel(writer, sheet_name='RollingTradeStatistics')
# Close the Pandas Excel writer and output the Excel file.
writer.save()
