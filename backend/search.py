# --------------------------------------------------------------------------
# IMPORTS
# --------------------------------------------------------------------------

import requests
import pandas as pd
import numpy as np
import re

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
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()

    def get_category_id(self, category_id=0):
        selection_df = []
        data = self.get_subcategories(category_id)
        categories = []

        if isinstance(data, dict):
            categories = data.get("categories") or []

        for dictionary in categories:
            if not isinstance(dictionary, dict):
                continue

            selection_df.append({
                "category": dictionary.get("name", ""),
                "category_id": dictionary.get("id")
            })

        return pd.DataFrame(selection_df)

    # ----------------------------------------------------------------------
    # SERIES RETRIEVAL
    # ----------------------------------------------------------------------

    def get_series(self, category_id):
        url = (
            "https://api.stlouisfed.org/fred/category/series"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json&limit=1000"
        )
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()

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
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        data = response.json()

        series_list = data.get("seriess", [])

        rows = []
        for s in series_list:
            rows.append({
                "series_id": s["id"],
                "title": s["title"],
                "frequency": s["frequency"],
                "units": s["units"],
                "seasonal_adjustment": s["seasonal_adjustment"],
                "observation_start": s.get("observation_start")
            })

        return rows

    # ----------------------------------------------------------------------
    # DATA RETRIEVAL
    # ----------------------------------------------------------------------

    def get_data(self, series_id, observation_start=None):
        url = (
            "https://api.stlouisfed.org/fred/series/observations"
            f"?series_id={series_id}&api_key={self.api_key}&file_type=json"
        )
        if observation_start:
            url += f"&observation_start={observation_start}"

        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        data = response.json()

        observations = data.get("observations", [])
        
        if not observations:
            return pd.DataFrame(columns=["date", "value"]) # Return an empty DataFrame with correct columns safely
            
        df = pd.DataFrame(observations)
        df["date"] = pd.to_datetime(df["date"])
        # Format the date to 'YYYY-MM-DD' string for better readability in Excel.
        df["date"] = df["date"].dt.strftime('%Y-%m-%d')

        df["value"] = pd.to_numeric(df["value"], errors="coerce")
        df = df[["date", "value"]]

        # Replace numpy's Not a Number (NaN) with Python's None, which is JSON compliant.
        df.replace({np.nan: None}, inplace=True)

        return df

    # ----------------------------------------------------------------------
    # SEARCH PARSING LOGIC
    # ----------------------------------------------------------------------

    def parse_search_input(self, user_input):
        """
        Parses the user input to extract the Series ID. 
        Handles both full FRED URLs and raw Series IDs.
        """
        user_input = user_input.strip()
        url_match = re.search(r"fred\.stlouisfed\.org/series/([A-Za-z0-9_]+)", user_input, re.IGNORECASE)
        
        if url_match:
            return url_match.group(1).upper()
        else:
            # Only clean if it's not a URL match
            clean_id = re.sub(r'[^A-Za-z0-9_]', '', user_input)
            return clean_id.upper()