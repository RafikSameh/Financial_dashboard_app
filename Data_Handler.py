import pandas as pd
import numpy as np
import plotly.io as pio


class Data_Handler:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.CashFlow = self.excel_sheet_handler()
        self.calculations_dict, self.totals = self.total_calculations(self.CashFlow)
        self.operating_cf_in = None
        self.operating_cf_out = None
        self.investing_cf_out = None
        self.financing_cf = None
        self.flow_classification()
        self.Cash_Flow_Forecast = None
        self.Cash_Flow_Forecast_Eg = None
        self.Cash_Flow_Forecast_KSA = None
        self.Cash_Flow_Forecast_Both = None
        self.cashflow_forecast_handler()
        self.operating_cf_in = self.cash_inflow_handler()
        self.calculations_dict, self.totals = self.total_calculations(self.CashFlow)

    def excel_sheet_handler(self):
        CashFlow = pd.read_excel(self.file_path, sheet_name="Cash Flow - User Input")
        CashFlow.drop(axis=0, labels=[0, 1], inplace=True)
        CashFlow.drop(axis=1, labels=["Unnamed: 0"], inplace=True)
        CashFlow.columns = CashFlow.iloc[0].values
        CashFlow.drop(axis=0, labels=[2], inplace=True)
        return CashFlow

    def total_calculations(self, CashFlow):
        cash_beginning_balance = CashFlow[CashFlow['Item'].str.strip() == 'Cash Beginning Balance'].iloc[:, 4:5].values[0][0]
        cash_ending_balance = CashFlow[CashFlow['Item'].str.strip() == 'Cash Ending Balance'].iloc[:, -1].values[0]
        totals = CashFlow[CashFlow['Item'].str.startswith("Total") == True]
        calculations_dict = {}
        calculations_dict['Cash Beginning Balance'] = cash_beginning_balance

        for i in range(len(totals)):
            calculations_dict[totals['Item'].iloc[i]] = totals.iloc[i, 4:].apply(pd.to_numeric, errors='coerce').sum()
        calculations_dict.pop('Total Operating Cash Outflow', None)
        calculations_dict.pop('Total Change in cash', None)
        calculations_dict['Cash Ending Balance'] = cash_ending_balance
        return calculations_dict, totals
    
    def flow_classification(self):
        for i in range(1,5):
            start_idx = self.CashFlow[self.CashFlow['Category'] == i].index[0]
            if i == 4:
                end_idx = self.CashFlow[self.CashFlow['Item'] == "Total Cash outflow from Financing "].index[0]
            else:
                end_idx = self.CashFlow[self.CashFlow['Category'] == i+1].index[0]
            if i == 1:
                self.operating_cf_in = self.CashFlow.loc[start_idx:end_idx-1]
            elif i == 2:
                self.operating_cf_out = self.CashFlow.loc[start_idx:end_idx-1]
            elif i == 3:
                self.investing_cf_out = self.CashFlow.loc[start_idx:end_idx-1]
            elif i == 4:
                self.financing_cf = self.CashFlow.loc[start_idx:end_idx]

    def cashflow_forecast_handler(self):
        self.Cash_Flow_Forecast = self.CashFlow[self.CashFlow[self.CashFlow['Item'].str.strip() == "Operating Cash Outflow"].index[0]-3:self.CashFlow[self.CashFlow['Item'].str.strip() == "Total Operating Cash Outflow"].index[0]-2]
        self.Cash_Flow_Forecast_Eg = pd.concat([self.Cash_Flow_Forecast[self.Cash_Flow_Forecast['Country'].str.strip() == "Egypt"].groupby('Category').sum().reset_index(False), self.Cash_Flow_Forecast[self.Cash_Flow_Forecast['Item'].str.strip() == "Total Cash outflow For Egypt"]]).drop(columns=['Country','Item','Cash Flow Type'])
        self.Cash_Flow_Forecast_KSA = pd.concat([self.Cash_Flow_Forecast[self.Cash_Flow_Forecast['Country'].str.strip() == "KSA"].groupby('Category').sum().reset_index(False), self.Cash_Flow_Forecast[self.Cash_Flow_Forecast['Item'].str.strip() == "Total Cash outflow For KSA"]]).drop(columns=['Country','Item','Cash Flow Type'])
        self.Cash_Flow_Forecast_Both = self.Cash_Flow_Forecast.groupby('Category').sum().drop(columns=['Country','Item','Cash Flow Type']).iloc[1:].reset_index()
        self.Cash_Flow_Forecast_Both.loc[len(self.Cash_Flow_Forecast_Both)] = ['Total Operating Cash Outflow'] + self.Cash_Flow_Forecast[self.Cash_Flow_Forecast['Item'].str.strip() == "Total Operating Cash Outflow"].iloc[:, 4:].apply(pd.to_numeric, errors='coerce').sum().tolist()
        self.Cash_Flow_Forecast_Eg["Category"].fillna("Total Operating Cash Outflow", inplace=True)
        self.Cash_Flow_Forecast_KSA["Category"].fillna("Total Operating Cash Outflow", inplace=True)
        self.Cash_Flow_Forecast_Both.iloc[:,1:] = self.Cash_Flow_Forecast_Both.iloc[:,1:].abs().fillna(0)
        self.Cash_Flow_Forecast_Eg.iloc[:,1:] = self.Cash_Flow_Forecast_Eg.iloc[:,1:].abs().fillna(0)
        self.Cash_Flow_Forecast_KSA.iloc[:,1:] = self.Cash_Flow_Forecast_KSA.iloc[:,1:].abs().fillna(0)



    def cash_inflow_handler(self):
        self.operating_cf_in.drop(columns=['Country','Cash Flow Type'],inplace=True)
        self.operating_cf_in = self.operating_cf_in.iloc[1:,:]
        self.operating_cf_in.fillna(0, inplace=True)
        MWAN = self.operating_cf_in[self.operating_cf_in['Category'].str.strip() == "MWAN Project"]
        MWAN.drop(columns=['Category'],inplace=True)
        MWAN.set_index('Item',inplace=True)

        NWC = self.operating_cf_in[self.operating_cf_in['Category'].str.strip() == "NWC Project"]
        NWC.drop(columns=['Category'],inplace=True)
        NWC.set_index('Item',inplace=True)

        return self.operating_cf_in