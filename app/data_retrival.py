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

# This function returns the primary FRED source used. We will definiitely need this helper function later on.
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
    
    def get_category_id(self):
        # pandas has built in dataframe compatibility for lists containing dicts.
        # This pd.Dataframelist([dict{}]) data structure will be needed for
        # Excel and JSON compatibility.
        selection_df = []

        for lst in self.get_subcategories().values():
            for dictionary in lst:
                selection_df.append({
                    "category": dictionary["name"],
                    "category_id": dictionary["id"]
                    })
                
        return pd.DataFrame(selection_df)
    
    #Function to retrieve data from FRED API and download it as a CSV to the working directory
    def get_data(self, series_id):
        url = (
            "https://api.stlouisfed.org/fred/series/observations"
            f"?series_id={series_id}&api_key={self.api_key}&file_type=json"
        )
        response = requests.get(url)
        data = response.json()

        df = pd.DataFrame(data["observations"])
        df["date"] = pd.to_datetime(df["date"])
        df["value"] = pd.to_numeric(df["value"], errors="coerce")

        return df[["date", "value"]]

if __name__ == "__main__":
    df = get_data(series_id, api_key)
    df.to_csv('gdp_data.csv', index=True)
    print(df)