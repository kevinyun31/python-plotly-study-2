# Import all the required packages and dataset 

from dash import Dash, html, dcc, Output, Input
import plotly.express as px
import pandas as pd
import openpyxl

df = pd.read_excel("supermarket_sales4.xlsx")

# --------------------------------------------------------------------------

#instantiating Dash
app = Dash(__name__)

#Defining the web layout
app.layout=html.Div([
    html.H1("Excel to Python App"),
    dcc.RadioItems(id='col-choice', options=['Gender','Customer type','City'], value='Gender'),
    dcc.Graph(id='our-graph', figure={}),
])

# --------------------------------------------------------------------------

@app.callback(
    Output('our-graph', 'figure'),
    Input('col-choice', 'value')
)
def update_graphs(column_selected):
    pivot_df = pd.pivot_table(df, values='Total', index=column_selected, columns='Payment', aggfunc='sum')
    fig = px.imshow(pivot_df)
    return fig

# --------------------------------------------------------------------------

if __name__=='__main__':
    app.run_server(port=8050)
    