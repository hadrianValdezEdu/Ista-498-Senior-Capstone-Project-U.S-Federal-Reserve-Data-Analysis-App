import requests
import pandas as pd
import numpy as np

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

def fred_source_link(series_id):
    return f"https://fred.stlouisfed.org/series/{series_id}"

# This class will be used to create a search button within excel.
# currently the method "user_selection()" within this class has to be updated in such
# a way that makes it better compatible with excel. this class for now will be for archival purposes
# just in case  we need to explore another method.
class Search():
    def __init__(self, api_key):
        self.api_key = api_key

    def get_subcategories(self, category_id=0):
        url = (
            "https://api.stlouisfed.org/fred/category/children"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json"
        )
        return requests.get(url).json()
    
    def user_selection(self):
            selection_df = []
            for lst in self.get_subcategories().values():
                for dictionary in lst:
                    id = dictionary["id"]
                    name = dictionary["name"]
                    selection_df.append(name)
                    selection_df.append(id)

            # this df is created so that integer values are ints, instead of  "1", "2", strings, etc.
            selection_df = pd.DataFrame(selection_df)

            # .reshape(-1,2) --> (rows, columns) --> (auto-compute the amount of rows needed, 3 columns)
            # each item within the df is reshaped into a list that resembles this --> [category, series_id]
            selection_df = np.array(selection_df).reshape(-1,2)

            # each item (which is a list of 2 items) becomes a df containing 2 columns
            selection_df = pd.DataFrame(selection_df, columns=["category", "series_id"])
            return selection_df
    

if __name__ == "__main__":
    df = get_data(series_id, api_key)
    df.to_csv('gdp_data.csv', index=True)
    print(df)