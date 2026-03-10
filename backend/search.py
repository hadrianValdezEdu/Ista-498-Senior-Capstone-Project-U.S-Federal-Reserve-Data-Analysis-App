import requests
import pandas as pd
import numpy as np

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
        df = df[["date", "value"]]

        # convert to obect type for JSON compatibility with Excel, since Excel cannot handle NaN values
        df["value"] = df["value"].astype("object")

        # replace NaN values introduced from pandas into None for JSON compatibility with Excel
        df = df.where(pd.notnull(df), None)

        return df