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
    
    def get_series(self, category_id):
        url = (
            "https://api.stlouisfed.org/fred/category/series"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json"
        )
        return requests.get(url).json()
    
    def get_series_df(self, category_id):
        # assigns this data var as a dict
        data = self.get_series(category_id)

        # returns a list containing ONLY what was in the key "seriess"
        series_lst = data.get("seriess", [])

        rows = []

        for series in series_lst:
            rows.append({
                "series_id": series["id"],
                "title": series["title"],
                "frequency": series["frequency"],
                "units": series["units"],
                "seasonal_adjustment": series["seasonal_adjustment"]
                })
        
        return pd.DataFrame(rows)
    
    def get_category_id(self, category_id=0):
        # pandas has built in dataframe compatibility for lists containing dicts.
        # This pd.Dataframelist([dict{}]) data structure will be needed for
        # Excel and JSON compatibility.
        selection_df = []

        # call helper function, it returns a dictionary
        for lst in self.get_subcategories(category_id).values():
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
    
    # If this function shouldn't be here let me know. -David
    # this function is designed to find any mising data points and replace those values with error.
    def missingValues(self, series_id):
        df = self.get_data(series_id)
        errors = df.isnull()
        df["error"] = errors["value"]
        return df

    # this function is safe to delete whenever needed.
def example():
    print("test_var.get_category_id()\n", test_var.get_category_id())
    print("test_var.get_data\n", test_var.get_data("GDP"))
    print("test_var.get_subcategories(category_id=32992)\n", test_var.get_subcategories(category_id=32992))
    print("test_var.get_series(32992)\n",test_var.get_series(32992))

if __name__ == "__main__":
    #df = get_data(series_id, api_key)
    #df.to_csv('gdp_data.csv', index=True)
    
    # example of Search()
    # api_key = "ENTER YOUR API KEY HERE"
    test_var = Search(api_key = "")
    example()
