import requests
import pandas as pd

series_id = 'GDP' 
#when you search for a data point on the FRED homepage and click on a chart the Series ID is listed in parentheses next to the title
api_key = 'Replace with you key. Do not push to github with the key in it'

'''
Unsure how we want to procced with adding API key, for security reasons I will not add it here yet.
Working on getting an .env and .gitignore setup to keep it secure when pushing to github.
for now you can add your API key directly in the function call below. Just do not Push it to cloud with the key in it.
'''

#Function to retrieve data from FRED API and download it as a CSV to the working directory
def get_data(series_id, api_key):
    url = f'https://api.stlouisfed.org/fred/series/observations?series_id={series_id}&api_key={api_key}&file_type=json'
    response = requests.get(url)
    data = response.json()
    
    observations = data['observations']
    df = pd.DataFrame(observations)
    df['date'] = pd.to_datetime(df['date'])
    df['value'] = pd.to_numeric(df['value'], errors='coerce')
    
    return df[['date', 'value']]

# This function returns the primary FRED source used. We will definiitely need this helper function later on.
def fred_source_link(series_id):
    return f"https://fred.stlouisfed.org/series/{series_id}"

df = get_data(series_id, api_key)
df.to_csv('gdp_data.csv', index=True)