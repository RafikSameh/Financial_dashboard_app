from urllib.request import DataHandler
import pandas as pd
import numpy as np
import plotly.io as pio
from plotly.graph_objs import Figure, Waterfall
import plotly.graph_objs as go
from Data_Handler import Data_Handler
from plotly.graph_objs import Figure, Waterfall
import plotly.graph_objs as go
from plotly.subplots import make_subplots
""" import ipywidgets as widgets
from IPython.display import display
from ipywidgets import VBox """


def excel_sheet_handler():
    CashFlow = pd.read_excel("Cash flow Forecasting. Yomn - RS.xlsx",sheet_name="Cash Flow - User Input")
    CashFlow.drop(axis=0,labels=[0,1],inplace=True)
    CashFlow.drop(axis=1,labels=["Unnamed: 0"],inplace=True)
    CashFlow.columns =CashFlow.iloc[0].values
    CashFlow.drop(axis=0,labels=[2],inplace=True)
    return CashFlow

def total_calculations(CashFlow):
    cash_beginning_balance = CashFlow[CashFlow['Item'].str.strip() == 'Cash Beginning Balance'].iloc[: , 4:5].values[0][0]
    cash_ending_balance = CashFlow[CashFlow['Item'].str.strip() == 'Cash Ending Balance'].iloc[: , -1].values[0]
    totals = CashFlow[CashFlow['Item'].str.startswith("Total") == True]
    dict = {}
    dict['Cash Beginning Balance'] = cash_beginning_balance

    for i in range(len(totals)):
        dict[totals['Item'].iloc[i]] = totals.iloc[i, 4:].apply(pd.to_numeric, errors='coerce').sum()
    dict.pop('Total Operating Cash Outflow', None)
    dict.pop('Total Change in cash', None)
    dict['Cash Ending Balance'] = cash_ending_balance
    return dict,totals


class plot_diagrams(Data_Handler):
    def __init__(self):
        self.DataHandler = Data_Handler("Cash flow Forecasting. Yomn - RS.xlsx")
        self.waterfall_cash_movement_fig = self.waterfall_cash_movement(self.DataHandler.calculations_dict)
        self.monthly_cash_flow_fig = self.monthly_cash_flow(self.DataHandler.totals)
        self.operating_cash_flow_diagram_fig = self.operating_cash_flow_diagram(self.DataHandler.Cash_Flow_Forecast_Both,
                                                                                self.DataHandler.Cash_Flow_Forecast_KSA,
                                                                                self.DataHandler.Cash_Flow_Forecast_Eg)
        #self.create_pie_and_bar_with_interactive_slider_fig = self.create_pie_and_bar_with_interactive_slider(self.DataHandler.operating_cf_in[:len(self.DataHandler.operating_cf_in)])
        

    def waterfall_cash_movement(self,dict):
        fig = go.Figure(
            data=[
                Waterfall(
                    x=list(dict.keys()),
                    y=list(dict.values()),
                    connector={"line": {"color": "rgb(63, 63, 63)"}},
                    increasing={"marker": {"color": "#28a745"}},
                    decreasing={"marker": {"color": "#dc3545"}},
                    totals={"marker": {"color": "#0066cc"}}
                )
            ]
        )
        fig.update_layout(
            title="Cash Movement Summary Water Fall (000s SAR)",
            yaxis_title="Amount (000s SAR)",
            xaxis_title="Category",
            autosize=True,
            margin={"l":40, "r":40, "t":60, "b":40}
        )
        return fig

    def monthly_cash_flow(self,totals):
        # Extract month columns (assuming columns 4 onward are months)
        months = totals.columns[4:]
        categories = ['Total Operating Cash Inflow', 'Total Operating Cash Outflow', 'Total Investing Cash Outflow', 'Total Cash outflow from Financing']

        # Prepare data for each category using only the 'totals' DataFrame
        data = []
        ending_balance = None
        for cat in categories:
            # Find the row in totals where the category name appears in the 'Item' column
            row = totals[totals['Item'].str.contains(cat, case=False, na=False)]
            if not row.empty:
                y = row.iloc[0, 4:].apply(pd.to_numeric, errors='coerce')
                bar = go.Bar(
                    name=cat,
                    x=months,
                    y=y/1000,
                )
                data.append(bar)



        # Add ending cash balance as a line
        ending_balance_row = self.DataHandler.CashFlow[self.DataHandler.CashFlow['Item'].str.strip() == 'Cash Ending Balance']
        if not ending_balance_row.empty:
            ending_balance = ending_balance_row.iloc[0, 4:].apply(pd.to_numeric, errors='coerce') / 1000
            line = go.Scatter(
                name='Ending Cash Balance',
                x=months,
                y=ending_balance,
                mode='lines+markers',
                yaxis='y',
                text=y,
                textposition='top center',  
            )
            data.append(line)
        # Calculate overall total for each month and add as text annotations
        annotations = []
        for i, val in enumerate(ending_balance):
            color = 'green' if val >= 0 else 'black'
            annotations.append(
                {
                    "x": months[i],
                    "y": val if val >= 0 else 0,  # Place label above zero for negative values
                    "text": f'{val:,.0f}',
                    "showarrow": False,
                    "font": {"color": color, "size": 14},
                    "yanchor": 'bottom' if val >= 0 else 'top'
                }
            )
        fig = go.Figure(data=data)
        fig.update_layout(
            barmode='stack',
            title='Monthly Cash Flow by Category',
            xaxis_title='Month',
            yaxis_title='Amount (000s SAR)',
            autosize=True,
            legend_title='Category',
            annotations=annotations,
            legend={
                "orientation":"h",   # Horizontal legend
                "y":-0.2,            # Position below the plot
                "x":0,
                "xanchor":'left',
                "yanchor":'top'
            }  
        )
        return fig


    def operating_cash_flow_diagram(self,Cash_Flow_Forecast_Both, Cash_Flow_Forecast_KSA, Cash_Flow_Forecast_Eg):
        # Helper function to create traces from a given DataFrame
        def create_traces(df, name):
            months = df.columns[1:]  # assuming first col = Category
            categories = df['Category'].tolist()

            traces = []
            # Bars for categories
            for cat in categories[:-1]:  # exclude the last row ("Total")
                row = df[df['Category'] == cat]
                if not row.empty:
                    y = row.iloc[0, 1:].apply(pd.to_numeric, errors='coerce')
                    traces.append(go.Bar(
                        name=cat,
                        x=months,
                        y=y/1000,
                        visible=False
                    ))

            # Line for total cash (last row)
            total_row = df.iloc[-1, 1:].apply(pd.to_numeric, errors='coerce') / 1000
            traces.append(go.Scatter(
                name=f'Total ({name})',
                x=months,
                y=total_row,
                mode='lines+markers+text',   # <-- line + points + labels
                text=[f"{val:,.0f}" for val in total_row],  # label each point
                textposition="top center",
                visible=False,
                line={"width": 3, "color": "black"}
            ))
            return traces


        # Collect traces for each dataframe
        traces = []
        df_list = [
            ("Total", Cash_Flow_Forecast_Both),
            ("KSA", Cash_Flow_Forecast_KSA),
            ("Egypt", Cash_Flow_Forecast_Eg)
        ]

        for name, df in df_list:
            traces.extend(create_traces(df, name))

        # Build dropdown buttons
        buttons = []
        trace_counts = [len(create_traces(df, name)) for name, df in df_list]

        start = 0
        for i, (name, df) in enumerate(df_list):
            mask = [False] * sum(trace_counts)  # all hidden by default
            for j in range(start, start + trace_counts[i]):
                mask[j] = True
            buttons.append({
                "label": name,
                "method": "update",
                "args": [{"visible": mask},
                      {"title": f"Monthly Cash Flow: {name}"}]
            })
            start += trace_counts[i]

        # Create the figure
        fig = go.Figure(data=traces)

        fig.update_layout(
        updatemenus=[{
            "active": 0,
            "buttons": buttons,
            "x": 1.2,
            "y": 1.1
        }],
        barmode='stack',
        title="Monthly Cash Flow",  # default view
        xaxis_title="Month",
        yaxis_title="Amount (000s SAR)",
        autosize=True,
        legend_title="Category",
        xaxis={
            "rangeselector": {
                "buttons": [
                    {"count": 3, "label": "3m", "step": "month", "stepmode": "backward"},
                    {"count": 6, "label": "6m", "step": "month", "stepmode": "backward"},
                    {"count": 12, "label": "1y", "step": "month", "stepmode": "backward"},
                    {"step": "all"}
                ]
            },
            "rangeslider": {"visible": True},
            "type": "date"
        }

    )

        # Show only the first dataframe (KSA) initially
        for i in range(trace_counts[0]):
            fig.data[i].visible = True

        return fig


""" def create_pie_and_bar_with_interactive_slider(self,df):
        # Melt into long format
        df_long = df[:len(df)-1].melt(id_vars=["Category", "Item"], var_name="Month", value_name="Value")
        df_long["Month"] = pd.to_datetime(df_long["Month"])
        months = sorted(df_long["Month"].unique())

        # Slider for selecting range
        options = [(f"{m.strftime('%Y-%m')}", m) for m in months]
        slider = widgets.SelectionRangeSlider(
            options=options,
            index=(0, len(months)-1),
            description='Period:',
            continuous_update=False,
            layout={'width': '80%'},
            style={'description_width': 'initial'}
        )

        # Dropdown for category selection
        category_selector = widgets.Dropdown(
            options=["All"] + sorted(df_long["Category"].astype(str).unique().tolist()),
            value="All",
            description="Category:",
            style={'description_width': 'initial'}
        )

        out = widgets.Output()

        def update_chart(change=None):
            with out:
                out.clear_output()

                # Selected range
                start, end = slider.value
                mask = (df_long["Month"] >= start) & (df_long["Month"] <= end)
                selected = df_long[mask]

                # Apply category filter if not "All"
                if category_selector.value != "All":
                    selected = selected[selected["Category"] == category_selector.value]

                # Pie aggregation
                if category_selector.value == "All":
                    pie_data = selected.groupby("Category")["Value"].sum().reset_index()
                    pie_labels = pie_data["Category"]
                else:
                    pie_data = selected.groupby("Item")["Value"].sum().reset_index()
                    pie_labels = pie_data["Item"]

                # Bar aggregation
                if category_selector.value == "All":
                    bar_data = selected.groupby(["Month", "Category"])["Value"].sum().reset_index()
                    bar_groups = bar_data["Category"].unique()
                    group_col = "Category"
                else:
                    bar_data = selected.groupby(["Month", "Item"])["Value"].sum().reset_index()
                    bar_groups = bar_data["Item"].unique()
                    group_col = "Item"

                # Create figure
                fig = make_subplots(
                    rows=1, cols=2,
                    specs=[[{"type": "domain"}, {"type": "xy"}]]
                )

                # Pie
                fig.add_trace(go.Pie(
                    labels=pie_labels,
                    values=pie_data["Value"],
                    textinfo="label+percent"
                ), row=1, col=1)

                # Bar (stacked by Category or Item)
                for grp in bar_groups:
                    grp_data = bar_data[bar_data[group_col] == grp]
                    fig.add_trace(go.Bar(
                        x=grp_data["Month"],
                        y=grp_data["Value"]/1000,  # divide by 1000 if needed
                        name=grp,
                    ), row=1, col=2)

                fig.update_layout(barmode="stack", showlegend=True,autosize=True)
                #fig.show()

        # Attach callbacks
        slider.observe(update_chart, names="value")
        category_selector.observe(update_chart, names="value")

        #display(slider, category_selector, out)

        # Trigger initial plot
        #update_chart()
        #ui = VBox([slider, category_selector, out])
        #display(ui)
        #update_chart()
        #return ui   """ 
