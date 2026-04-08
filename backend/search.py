# --------------------------------------------------------------------------
# IMPORTS
# --------------------------------------------------------------------------

import requests
import pandas as pd
import numpy as np

# --------------------------------------------------------------------------
# SEARCH CLASS INITIALIZATION
# --------------------------------------------------------------------------

class Search():
    def __init__(self, api_key):
        self.api_key = api_key

    # ----------------------------------------------------------------------
    # CATEGORY RETRIEVAL
    # ----------------------------------------------------------------------

    def get_subcategories(self, category_id=0):
        url = (
            "https://api.stlouisfed.org/fred/category/children"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json"
        )
        return requests.get(url).json()

    def get_category_id(self, category_id=0):
        selection_df = []

        for lst in self.get_subcategories(category_id).values():
            for dictionary in lst:
                selection_df.append({
                    "category": dictionary["name"],
                    "category_id": dictionary["id"]
                })

        return pd.DataFrame(selection_df)

    # ----------------------------------------------------------------------
    # SERIES RETRIEVAL
    # ----------------------------------------------------------------------

    def get_series(self, category_id):
        url = (
            "https://api.stlouisfed.org/fred/category/series"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json"
        )
        return requests.get(url).json()

    def get_series_df(self, category_id):
        data = self.get_series(category_id)
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

    def get_series_info(self, series_id):
        url = (
            "https://api.stlouisfed.org/fred/series"
            f"?series_id={series_id}&api_key={self.api_key}&file_type=json"
        )
        data = requests.get(url).json()

        series_list = data.get("seriess", [])

        rows = []
        for s in series_list:
            rows.append({
                "series_id": s["id"],
                "title": s["title"],
                "frequency": s["frequency"],
                "units": s["units"],
                "seasonal_adjustment": s["seasonal_adjustment"]
            })

        return rows

    # ----------------------------------------------------------------------
    # DATA RETRIEVAL
    # ----------------------------------------------------------------------

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

        df["value"] = df["value"].astype("object")
        df = df.where(pd.notnull(df), None)

        return df