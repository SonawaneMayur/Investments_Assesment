##To access the Trade files
##
import pandas as pd

class Trade_Files:

    __trade_orders = 'C://Users//Mayur//Documents//Assesment//Qtron_Investments//Trade_Files//Trade_Orders.xlsx'
    __trade_confirms = 'C://Users//Mayur//Documents//Assesment//Qtron_Investments//Trade_Files//Trade_Confirms.xlsx'
    __trade_stats = 'C://Users//Mayur//Documents//Assesment//Qtron_Investments//Trade_Files//Trade_Stats.xlsx'

    def __init__(self, trade_orders = None, trade_confirms = None, trade_stats = None):

        if trade_orders == None:
            self.trade_orders = Trade_Files.__trade_orders
        else:
            self.trade_orders = trade_orders
        if trade_confirms == None:
            self.trade_confirms = Trade_Files.__trade_confirms
        else: self.trade_confirms = None

        if trade_stats == None:
            self.trade_stats = Trade_Files.__trade_stats
        else: self.trade_stats = None

    def get_trade_order_dataframe(self):
        return pd.read_excel(self.trade_orders, sheetname='Sheet1', header=0)

    def get_trade_confirms_dataframe(self):
        return pd.read_excel(self.trade_confirms,sheetname='KAS-Web export 22062017', header=0)

    def get_trade_output_file(self):
        return self.trade_stats
#e = Trade_Files("d","da")
#print(e.trade_orders)
def xyz():
    pass

#e= Trade_Files()
#print(e.get_trade_output_file())