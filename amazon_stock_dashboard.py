import pandas as pd
import plotly.graph_objs as go
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc

# Load the dataset
file_path = r'C:\Users\13025\OneDrive\Desktop\coding\amazon monthly.xlsx'
try:
    data = pd.read_excel(file_path, sheet_name='AMAZON_monthly')
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Convert the 'Date' column to datetime format
data['Date'] = pd.to_datetime(data['Date'])

# Set 'Date' column as index
data.set_index('Date', inplace=True)

# Calculate KPIs
max_price = data['Close'].max()
total_growth = ((data['Close'].iloc[-1] - data['Close'].iloc[0]) / data['Close'].iloc[0]) * 100
avg_diff_open_adj_close = (data['Open'] - data['Adj Close']).mean()
avg_monthly_growth = data['% of change monthly'].mean()

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Layout of the dashboard
app.layout = dbc.Container([
    html.H1("Amazon Stock Analysis Dashboard", className="mt-3 mb-4"),

    # KPIs
    dbc.Row([
        dbc.Col([
            html.Div([
                html.H3("Maximum Price"),
                html.P(f"${max_price:.2f}", id='max-price', className="kpi-value")
            ], className="kpi"),
        ], width=4),
        dbc.Col([
            html.Div([
                html.H3("Total Percentage Growth"),
                html.P(f"{total_growth:.2f}%", id='total-growth', className="kpi-value")
            ], className="kpi"),
        ], width=4),
        dbc.Col([
            html.Div([
                html.H3("Average Difference Between Open and Adj. Close Prices"),
                html.P(f"${avg_diff_open_adj_close:.2f}", id='avg-diff', className="kpi-value")
            ], className="kpi"),
        ], width=4),
    ], className="kpi-container mb-4"),

    # Descriptive statistics
    dbc.Card([
        dbc.CardHeader("Descriptive Statistics"),
        dbc.CardBody([
            html.P("Summary statistics of the dataset:"),
            dbc.Table.from_dataframe(data.describe().loc[['mean', 'std', 'min', 'max', '25%', '50%', '75%']].transpose(), striped=True, bordered=True, hover=True)
        ]),
    ], className="mb-4"),

    # Date range slider
    dcc.RangeSlider(
        id='date-slider',
        min=data.index.min().timestamp(),
        max=data.index.max().timestamp(),
        step=None,
        marks={str(timestamp): str(date.date()) for timestamp, date in zip(data.resample('MS').first().index.astype(str), data.resample('MS').first().index)},
        value=[data.index.min().timestamp(), data.index.max().timestamp()]
    ),

    # Time series plot for Open, High, Low, Close prices
    dcc.Graph(id='price-trends'),

    # Volume trends over time
    dcc.Graph(id='volume-trends'),

    # Percentage change graphs
    dcc.Graph(id='percentage-change'),
], className="pt-3")

# Callback to update the graphs based on the selected date range
@app.callback(
    [Output('price-trends', 'figure'),
     Output('volume-trends', 'figure'),
     Output('percentage-change', 'figure')],
    [Input('date-slider', 'value')]
)
def update_graphs(selected_date_range):
    filtered_data = data.loc[pd.to_datetime(selected_date_range[0], unit='s'):pd.to_datetime(selected_date_range[1], unit='s')]

    # Time series plot for Open, High, Low, Close prices
    price_trends_fig = {
        'data': [
            go.Scatter(x=filtered_data.index, y=filtered_data['Open'], mode='lines', name='Open', line=dict(color='blue')),
            go.Scatter(x=filtered_data.index, y=filtered_data['High'], mode='lines', name='High', line=dict(color='green')),
            go.Scatter(x=filtered_data.index, y=filtered_data['Low'], mode='lines', name='Low', line=dict(color='red')),
            go.Scatter(x=filtered_data.index, y=filtered_data['Close'], mode='lines', name='Close', line=dict(color='orange'))
        ],
        'layout': go.Layout(
            title='Amazon Stock Prices Over Time',
            xaxis={'title': 'Date'},
            yaxis={'title': 'Price'},
            hovermode='closest'
        )
    }

    # Volume trends over time
    volume_trends_fig = {
        'data': [
            go.Bar(x=filtered_data.index, y=filtered_data['Volume'], name='Volume', marker=dict(color='purple'))
        ],
        'layout': go.Layout(
            title='Amazon Monthly Trading Volume',
            xaxis={'title': 'Date'},
            yaxis={'title': 'Volume'},
            hovermode='closest'
        )
    }

    # Percentage change graphs
    percentage_change_fig = {
        'data': [
            go.Scatter(x=filtered_data.index, y=filtered_data['% of change monthly'], mode='lines', name='% Change Monthly', line=dict(color='blue')),
            go.Scatter(x=filtered_data.index, y=filtered_data['% of change from start'], mode='lines', name='% Change from Start', line=dict(color='green'))
        ],
        'layout': go.Layout(
            title='Amazon Percentage Change Over Time',
            xaxis={'title': 'Date'},
            yaxis={'title': 'Percentage Change'},
            hovermode='closest'
        )
    }

    return price_trends_fig, volume_trends_fig, percentage_change_fig

if __name__ == '__main__':
    app.run_server(debug=True)
