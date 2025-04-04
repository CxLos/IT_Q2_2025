# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
import datetime
import folium
import os
import sys
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component
# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
data_path = 'data/IT_Responses.xlsx'
file_path = os.path.join(script_dir, data_path)
data = pd.read_excel(file_path)
df = data.copy()

# Trim leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# Filtered df where 'Date:' is between Ocotber to December:
df['Date:'] = pd.to_datetime(df['Date:'], errors='coerce')
df = df[(df['Date:'] >= '2024-10-01') & (df['Date:'] <= '2024-12-31')]

# print(df_m.head())
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns)
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())

# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

# Column Names: 
#  Index([
#        'Timestamp', 
#        'Which form are you filling out?',
#        'Person completing this form:',
#        'Was all required IT equipment purchased/ serviced this month?',
#        'If answered "No" to above question, please specify why:',
#        'Were all IT equipment support request addressed?',
#        'Was phone system maintained and any support issues resolved within 48 hours?',
#        'Was page speed optimization completed this month?',
#        'Were any 404 errors identified and fixed?',
#        'Were Updates made to the sitemap, internal linking, or robot.txt as needed?',
#        'Was a Database / cloud backup completed?',
#        'Did you complete any content or layout updates on the website?',
#        'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:',
#        'Person completing this form:.1', 'Was a Security Audit Conducted?',
#        'If yes, were all issues addressed?',
#        'Were any new Cybersecurity vulnerabilities identified?',
#        'Did all Scheduled Cybersecurity training sessions occur this month?',
#        'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.1',
#        'Person completing this form',
#        'Were automated Workflows reviewed and adjusted as necessary?',
#        'Were all planned email campaigns executed?',
#        'Were A/B tests conducted on landing pages or customer journeys?',
#        'Were any necessary SEO updates made this month?',
#        'Were all monthly analytics reports generated and reviewed?',
#        'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.2',
#        'Person completing this form:.2',
#        'Did you attend or host all scheduled events this month?',
#        'Were all new or potential community partnerships engaged as planned?',
#        'Did you follow up with all attendees or participants from recent events?',
#        'Were all planned outreach campaigns completed?',
#        'Did you track social media engagement metrics?',
#        'Was feedback from the community collected and reviewed?',
#        'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.3',
#        'Person completing this form:.3',
#        'Did all planned technical training sessions for staff occur?',
#        'Did all scheduled employees complete the training?',
#        'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.4',
#        'Person completing this form:.4',
#        'Were all Weekly and Monthly reports completed and submitted on time?',
#        'Was data collected accurately and reviewed for quality?',
#        'Did you identify any actionable insights from this month's data?',
#        'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.5',
#        'Date:', 
# 'Briefly describe what tasks you worked on:',
#        'How much time did you spend on these tasks? (minutes)'],
#       dtype='object')

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

# ============================== Data Preprocessing ========================== #



# ========================= Filtered DataFrames ========================== #

# List of columns to include
columns_to_include = [
    'Person completing this form:.4',
    'Were all Weekly and Monthly reports completed and submitted on time?',
    'Was data collected accurately and reviewed for quality?',
    "Did you identify any actionable insights from this month's data?",
    'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.5',
    'Date:',
    'Briefly describe what tasks you worked on:',
    'How much time did you spend on these tasks? (minutes)'
]

columns_to_include_2 = [
    'Date:',
    'Person completing this form:.4',
    'Briefly describe what tasks you worked on:',
    'How much time did you spend on these tasks? (minutes)'
]

# Create a new DataFrame with only the specified columns
df_data = df[columns_to_include]
df_data2 = df[columns_to_include_2]
# print(df_data.head(10))

# Total data events
data_events = len(df_data)

# Total Data Hours
total_data_hours = df_data['How much time did you spend on these tasks? (minutes)'].sum()/60
total_data_hours = round(total_data_hours)

# "Person completing this form:" dataframe:
df['Person completing this form:.4'] = df['Person completing this form:.4'].str.strip()
df_person = df.groupby('Person completing this form:.4').size().reset_index(name='Count')
# print(df_person.value_counts())

# Data Table 2
columns_to_include_2 = [
    'Date:',
    'Person completing this form:.4',
    'Briefly describe what tasks you worked on:',
    'How much time did you spend on these tasks? (minutes)'
]

df_data2 = df[columns_to_include_2]

# # ========================== DataFrame Table ========================== #

# df_data2 data table
data_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_data2.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_data2[col] for col in df_data2.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

data_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server   

app.layout = html.Div(
  children=[ 
    html.Div(
        className='divv', 
        children=[ 
          html.H1(
              'IT Report Q1 2025', 
              className='title'),
          html.Div(
              className='btn-box', 
              children=[
                  html.A(
                    'Repo',
                    href='https://github.com/CxLos/IT_Q1_2025',
                    className='btn'),
    ]),
  ]),    

# Data Table
# html.Div(
#     className='row0',
#     children=[
#         html.Div(
#             className='table',
#             children=[
#                 html.H1(
#                     className='table-title',
#                     children='Data Table'
#                 )
#             ]
#         ),
#         html.Div(
#             className='table2', 
#             children=[
#                 dcc.Graph(
#                     className='data',
#                     figure=marcom_table
#                 )
#             ]
#         )
#     ]
# ),

# ROW 1
html.Div(
    className='row0',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=['Total Data Events:']
            ),
            html.Div(
                className='circle1',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high3',
                            children=[data_events]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high2',
                children=['MarCom Hours:']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[total_data_hours]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph1',
            children=[
                html.Div(
                    className='table',
                    children=[
                        html.H1(
                            className='table-title',
                            children='Data Events Table'
                        )
                    ]
                ),
                html.Div(
                    className='table2',
                    children=[
                        dcc.Graph(
                            className='data',
                            figure=data_table
                        )
                    ]
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[
            # "Person completing this form:" bar chart
            dcc.Graph(
                className='bar',
                figure=px.bar(
                    df_person,
                    x='Person completing this form:.4',
                    y='Count',
                    color='Person completing this form:.4',
                    text='Count',
                ).update_layout(
                    title='Person Completing this Form',
                    xaxis_title='Person',
                    yaxis_title='Count',
                    title_x=0.5
                ).update_traces(
                        textposition='auto',
                        textangle=0, 
                        hovertemplate= '<b>%{label}</b><br><b>Count</b>: %{y}<extra></extra>'
                )
            )
            ],
        )
    ]
),
])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=True)
                #    False)
# =================================== Updated Database ================================= #

# updated_path = 'data/bmhc_q4_2024_cleaned.xlsx'
# data_path = os.path.join(script_dir, updated_path)
# df.to_excel(data_path, index=False)
# print(f"DataFrame saved to {data_path}")

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx