# -*- coding: utf-8 -*-
"""
@author: mk9999
"""

import backtrader as bt
import pandas as pd
import datetime
import win32com.client

global mkt
global message
message = ""

class TD(bt.Strategy):
    
    '''
        These are the control parameters for TD Sequential, more on these can be
        found in: Advanced Techincal Analysis (2011, page 127) by Trevor Neil 
        candles_past_to_compare:
            Don't change, used to evaluate the trigger
        cancel_1:
            Cancel if high of buy setup is exceeded anytime by a low of countdown
            Cancel if low of sell setup is exceeded anytime by a high of countdown
        cancel_2:
            Cancel if a setup is completed in opposite direction of current
            countdown and start countdown in the new direction
        cancel_3:
            Cancel if a new setup in the same direction occurs -> recycling
            0 = Don't use -> ignore new setups;
            1 = Always use -> never ignore new setups;
            2 = Use with exception -> use the cancel method but ignore new
                setups with all closes in the range of previous setup's high/low
        recycle_12:
            If True it uses 12 bars instead of 9 as a valid new setup in the same
            direction (Cancel_3) if countdown has already started
        aggressive_countdown:
            If true, compares lows to lows instead od lows to closes in countdown
        cancel_1618:
            If a new setup's High-Low is 1.618 times larger than original 
            setups High-Low, the new setup is not used
    '''
    
    params = dict(
      candles_past_to_compare = 4, # DO NOT CHANGE
      cancel_1 = True,
      cancel_2 = True,
      cancel_3 = 2,
      recycle_12 = True,
      aggressive_countdown = False,
      cancel_1618 = True,
    )
  
    '''*********************** DO NOT CHANGE CODE BELOW THIS LINE *************'''


    '''Initialization '''    
    def __init__(self):
        # Keep a reference to the primary data feed
        self.dataprimary = self.datas[0]
        # Keep a reference to the "close" line in the primary dataseries
        self.dataclose = self.dataprimary.close
        # Control vars
        self.live = True
        self.order = None

        self.buyTrig = False
        self.sellTrig = False

        self.tdsl = 0 # TD sequence long
        self.tdss = 0 # TD sequence short
        self.buySetup = False # TD buy setup 
        self.sellSetup = False # TD sell setup
        self.buyCountdown = 0 # TD buy countdown
        self.sellCountdown = 0 # TD sell countdown
        self.buyVal = 0 # buy value at intersection
        self.sellVal = 0 # sell value at intersection
        
        self.buyprice = 0
        self.buycomm = 0
        
        self.buySig = False
        self.idealBuySig = False
        self.sellSig = False
        self.idealSellSig = False
        
        # Used only for plotting the last setup number
        self.buy_nine = False
        self.sell_nine = False
        
        # Used for determining Cancel_1
        self.buy_high = 999999
        self.buy_low = 0
        self.sell_high = 999999
        self.sell_low = 0
                      
        self.use_cancel_1 = self.p.cancel_1
        self.use_cancel_2 = self.p.cancel_2
        self.use_cancel_3 = self.p.cancel_3
        self.recycle_12 = self.p.recycle_12
        self.aggressive_countdown = self.p.aggressive_countdown
        self.use_cancel_1618 = self.p.cancel_1618

    def log(self, txt, dt=None, doprint=False):
        ''' Logging function for this strategy'''
        if doprint:
            dt = dt or self.datas[0].datetime.date(0)
            print('%s, %s' % (dt.isoformat(), txt))

    ''' Data notification '''
    def notify_data(self, data, status, *args, **kwargs):
#        print('*' * 5, 'DATA NOTIF:', data._getstatusname(status), *args)
        if status == data.LIVE:
            self.live = True
        elif status == data.DELAYED:
            self.live = False


    ''' Store notification '''
    def notify_order(self, order):
        if order.status in[order.Submitted, order.Accepted]:
            return
        if order.status in [order.Completed]:
            if order.isbuy():
                self.log(
                    'BUY EXECUTED, Price: %.2f, Cost: %.2f, Comm %.2f' %
                    (order.executed.price,
                     order.executed.value,
                     order.executed.comm))
                
                self.buyprice = order.executed.price
                self.buycomm = order.executed.comm
                
            elif order.issell():
                #self.log('SELL EXECUTED, %.2f' % order.executed.price)
                
                self.log('SELL EXECUTED, Price: %.2f, Cost: %.2f, Comm %.2f' %
                         (order.executed.price,
                          order.executed.value,
                          order.executed.comm))
            self.bar_executed = len(self)
            
        elif order.status in [order.Canceled, order.Margin, order.Rejected]:
            self.log('Order Canceled/Margin/Rejected')

        self.order = None


    ''' Trade notification '''
    def notify_trade(self, trade):
#        print('-' * 50, 'TRADE BEGIN', datetime.datetime.now())
#        print(trade)
#        print('-' * 50, 'TRADE END')
        pass
        if not trade.isclosed:
            return

#        self.log('OPERATION PROFIT, GROSS %.2f, NET %.2f' %
#            (trade.pnl, trade.pnlcomm))


    def reset_on_cancel_1(self):
        '''
        This method handles setup cancelation based on cancel method 1.
        Called by next().
        '''
        self.buySetup = False
        self.sellSetup = False
        self.buyCountdown = 0
        self.sellCountdown = 0
        self.buy_high = 999999
        self.buy_low = 0
        self.sell_high = 999999
        self.sell_low = 0
        
    def reset_setup(self, buy_or_sell):
        '''
        This method resets just the setup variables without changing countdown
        variables. Used only for setup reset.
        '''
        if buy_or_sell == "B":
            self.buyTrig = False
            self.tdsl = 0

        elif buy_or_sell =="S":
            self.sellTrig = False
            self.tdss = 0
    
    def reset_countdown(self, buy_or_sell, count):
        '''
        This method handles countdown variable setting and setup cancelation
        after a setup 9 count has formed. Also handles cancel_2 setup
        cancelation.
        Called by next() and reset_on_cancel_3().
        '''
        
        if(buy_or_sell == "B"):
            # This is after buy setup confirms, if 8 or 9 are below 6 and 7 it triggers a setup buy signal
            self.buySig = ((self.dataprimary.low[0] < self.dataprimary.low[-3]) and (self.dataprimary.low[0] < self.dataprimary.low[-2])) or ((self.dataprimary.low[-1] < self.dataprimary.low[-2]) and (self.dataprimary.low[-1] < self.dataprimary.low[-3]))
            if self.tdsl == 9:
                self.buy_nine = True
            self.reset_setup(buy_or_sell)
            self.buySetup = True
            
            # Cancel_2 logic
            if self.use_cancel_2:
                self.sellSetup = False
                self.sellCountdown = 0
            self.buyCountdown = 0
            self.buy_high = max(self.dataprimary.high[n] for n in range(-(count-1),0))
            self.buy_low = min(self.data.low[n] for n in range(-(count-1),0))
    
    
        if(buy_or_sell == "S"):
            # This is  after sell setup confirms, if 8 or 9 are below 6 and 7 it triggers a setup sell signal
            self.sellSig = ((self.dataprimary.high[0] > self.dataprimary.high[-2]) and (self.dataprimary.high[0] > self.dataprimary.high[-3])) or ((self.dataprimary.high[-1] > self.dataprimary.high[-3]) and (self.dataprimary.high[-1] > self.dataprimary.high[-2]))
            if self.tdss == 9:
                self.sell_nine = True
            self.reset_setup(buy_or_sell)
            self.sellSetup = True
            
            # Cancel_2 logic
            if self.use_cancel_2:
                self.buySetup = False
                self.buyCountdown = 0
            self.sellCountdown = 0
            self.sell_high = max(self.dataprimary.high[n] for n in range(-(count-1),0))
            self.sell_low = min(self.data.low[n] for n in range(-(count-1),0))
            
            
    def reset_on_cancel_3(self, buy_or_sell, mode, count):
        
        '''
        This method handles countdown cancelation based on the settings for
        cancel_2, cancel_3, cancel_1618. Variable count determines whether the 
        TD setup was 9 bars or 12 bars.
        '''
        
        if (buy_or_sell == "B"):
            buy_diff_1618 = 0
            if self.use_cancel_1618 and self.buySetup:
                buy_high_1618 = max(self.dataprimary.high[n] for n in range(-(count-1),0))
                buy_low_1618 = min(self.data.low[n] for n in range(-(count-1),0))
                buy_diff_1618 = (buy_high_1618 - buy_low_1618) - (self.buy_high-self.buy_low)*1.618
            
#           
            # 1618 buy cancel
            if (buy_diff_1618 > 0) or (mode == 0):
                # Keep the current countdown and reset setup count
#                    print("1618 buy cancel", self.dataprimary.datetime.datetime(0))
                    self.reset_setup(buy_or_sell)
                    return
            
            if(mode == 1):
               self.reset_countdown(buy_or_sell, count)
                
            if(mode == 2):
                low = all(self.buy_low < self.dataprimary.close[n] for n in range(-(count-1),0))
                high = all(self.buy_high > self.dataprimary.close[n] for n in range(-(count-1),0))
               
                # Keep the current countdown and reset setup count
                if high and low:
#                    print("Buy setup cancel 3",self.dataprimary.datetime.datetime(0))
                    self.reset_setup(buy_or_sell)
                    return
                # Start a new buy countdown
                else:
                    self.reset_countdown(buy_or_sell, count)
                
    
        elif (buy_or_sell == "S"):
            sell_diff_1618 = 0
            if self.use_cancel_1618 and self.sellSetup:
                sell_high_1618 = max(self.dataprimary.high[n] for n in range(-(count-1),0))
                sell_low_1618 = min(self.data.low[n] for n in range(-(count-1),0))
                sell_diff_1618 = (sell_high_1618 - sell_low_1618) - (self.sell_high-self.sell_low)*1.618
            
            # 1618 sell cancel
            if (sell_diff_1618 > 0) or (mode == 0):
            # Keep the current countdown and reset setup count
#                print("1618 sell cancel", self.dataprimary.datetime.datetime(0))
                self.reset_setup(buy_or_sell)
                return
                
            if(mode == 1):
                self.reset_countdown(buy_or_sell, count)
                
            if(mode == 2):
                low = all(self.sell_low < self.dataprimary.close[n] for n in range(-(count-1),0))
                high = all(self.sell_high > self.dataprimary.close[n] for n in range(-(count-1),0))
                
                # Keep the current sell countdown and reset the setup count
                if high and low:
#                    print("Sell setup cancel 3",self.dataprimary.datetime.datetime(0))
                    self.reset_setup(buy_or_sell)
                    return
                # Start a new sell countdown
                else:
                    self.reset_countdown(buy_or_sell, count)
    
    
    
    ''' Next '''        
    def next(self):
        self.buySig = False
        self.sellSig = False
        self.idealBuySig = False
        self.idealSellSig = False
        self.buy_nine = False
        self.sell_nine = False
        self.mkt_data = [0,0,0,0]
    
        # Calculate td sequential values if enough bars
        if len(self.dataclose) > self.p.candles_past_to_compare:
            
            # Begin of sequence, bullish or bearish trigger
            # Buy; first candle and trigger
            if self.dataclose[0] < self.dataclose[-self.p.candles_past_to_compare] and self.dataclose[-1] > self.dataclose[-(self.p.candles_past_to_compare + 1)]:
                self.buyTrig = True
                self.sellTrig = False
                self.tdsl = 0
                self.tdss = 0
                
            # Sell; first candle and trigger 
            elif self.dataclose[0] > self.dataclose[-self.p.candles_past_to_compare] and self.dataclose[-1] < self.dataclose[-(self.p.candles_past_to_compare + 1)]:
                self.sellTrig = True
                self.buyTrig = False
                self.tdss = 0
                self.tdsl = 0

            # Setup start
            # Buy setup numbering
            if self.dataclose[0] < self.dataclose[-self.p.candles_past_to_compare] and self.buyTrig:
                self.tdsl += 1
                self.mkt_data[0] = self.tdsl
                
            # Sell setup numbering
            elif self.dataclose[0] > self.dataclose[-self.p.candles_past_to_compare] and self.sellTrig:
                self.tdss += 1
                self.mkt_data[1] = self.tdss
                
#   BUY SETUPS            
            ''' Buy setup reaches 9, check for buy signal and cancelation of previous signal'''
            if self.tdsl == 9:
                # If you're already in a countdown, set things based on cancellation method 3 and recycling settings (9 or 12 bars)
                if self.buySetup is True:
                    if (self.recycle_12 is False) or (self.use_cancel_3 == 0):
                        self.reset_on_cancel_3("B", self.use_cancel_3, self.tdsl)
                # Else set a new countdown
                else:
                    self.reset_countdown("B", self.tdsl)
                    
            # Used only if self.recycle_12 is set to true        
            elif self.tdsl == 12:
                self.reset_on_cancel_3("B", self.use_cancel_3, self.tdsl)
                
#   SELL SETUPS            
            ''' Sell setup reaches 9, check for sell signal and cancelation of previous signal'''
            if self.tdss == 9:
                # If you're already in a countdown, set things based on cancellation method 3 and recycling settings (9 or 12 bars)
                if self.sellSetup is True:
                    if (self.recycle_12 is False) or (self.use_cancel_3 == 0):
                        self.reset_on_cancel_3("S", self.use_cancel_3, self.tdss)
                # Else set a new countdown
                else:
                    self.reset_countdown("S", self.tdss)
                    
            # Used only if self.recycle_12 is set to true
            elif self.tdss == 12:
                self.reset_on_cancel_3("S", self.use_cancel_3, self.tdss)
                
                
            
            # Cancel setup 1
            # Cancel buy
            if self.use_cancel_1 and self.buySetup and (self.buy_high < self.dataprimary.low[0]):
                self.reset_on_cancel_1()
#                print("Cancel 1 buy",self.dataprimary.datetime.datetime(0))
            # Cancel sell
            elif self.use_cancel_1 and self.sellSetup and (self.sell_low > self.dataprimary.high[0]):
                self.reset_on_cancel_1()
#                print("Cancel 1 sell",self.dataprimary.datetime.datetime(0))

            # 
            if self.aggressive_countdown:
                countdown_compare = self.dataprimary.low[0]
                
            else:
                countdown_compare = self.dataprimary.close[0]
            
            '''Sequence countdown'''
            #Buy countdown    
            if self.buySetup:
                if countdown_compare <= self.dataprimary.low[-2]:
                    self.buyCountdown += 1
                    self.mkt_data[2] = self.buyCountdown
                    if self.buyCountdown > 13:
                        self.buyCountdown = 13
                if self.buyCountdown == 8:
                    self.buyVal = countdown_compare
                elif self.buyCountdown == 13:
                    if self.dataprimary.low[0] <= self.buyVal:
                        self.idealBuySig = True
                        self.buySetup = False
                        self.buyCountdown = 0
            
            # Sell countdown
            if self.sellSetup:
                if countdown_compare >= self.dataprimary.high[-2]:
                    self.sellCountdown += 1
                    self.mkt_data[3] = self.sellCountdown 
                    if self.sellCountdown > 13:
                        self.sellCountdown = 13
                if self.sellCountdown == 8:
                    self.sellVal = countdown_compare
                elif self.sellCountdown == 13:
                    if self.dataprimary.high[0] >= self.sellVal:
                        self.idealSellSig = True
                        self.sellSetup = False
                        self.sellCountdown = 0
            date=self.dataprimary.datetime.datetime(0).strftime("%d.%m.%Y")
            if date == datetime.datetime.today().strftime("%d.%m.%Y"):#datetime.date(2018,02,06).strftime("%d.%m.%Y"):
                set_email_data(self.mkt_data)
            
def set_email_data(data):
    
    '''
    This function sets and filters the data to be sent via email
    '''
    
    setup_list = [7,9,11,12]
    countdown_list = [9,10,11,12,13]
    global message
    
    if data[0] in setup_list:
        message = message +mkt+ " Buy Setup: " + str(data[0])+"\n <br>"
        
    if data[1] in setup_list:
        message = message +mkt+ " Sell Setup: " + str(data[1])+"\n <br>"
        
    if data[2] in countdown_list:
        message = message +mkt+ " Buy Countdown: " + str(data[2])+"\n <br>"
        
    if data[3] in countdown_list:
        message = message +mkt+ " Sell Countdown: " + str(data[3])+"\n <br>"


def add_text():
#    return "\n<br> This is a TD sequential signal."
    return ""

def send_mail(message, recipients, sender):
    '''
    This function sends the email if the message is not empty
    '''
    
    
#    message = message + add_text()
    o = win32com.client.Dispatch("Outlook.Application")
    Msg = o.CreateItem(0)
    Msg.Subject = "TD Sequential " + datetime.datetime.today().strftime("%d.%m.%Y")
    Msg.HTMLBody = message
    Msg.To = recipients
    Msg.SentOnBehalfOfName = sender
    if message == "":
        Msg.HTMLBody = "No updates to TD seq"
        Msg.To =""
    Msg.Send()
    print(message)
    
    


def td_start(markets, recipients, sender, datapath):
    '''
    This function starts backtrader. It is called from the master file.
    '''
    
    global message
    for n in range(len(markets)):
        global mkt
        mkt=markets[n]
        cerebro = bt.Cerebro()
        cerebro.addstrategy(TD)
        try:
            dataframe = pd.read_excel(datapath,
                                      sheet_name=markets[n],
                                      index_col=1)
        except Exception as e:
            print(e)
            continue
    
        data = bt.feeds.PandasData(dataname=dataframe)
        cerebro.adddata(data, name="real")
        cerebro.run(optreturn=True)
#    return message
    send_mail(message,recipients, sender)
        