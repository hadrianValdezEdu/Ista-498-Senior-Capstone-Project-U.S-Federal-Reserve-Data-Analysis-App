"""
Search Module for Federal Reserve Economic Data (FRED) API Integration.

This module provides the Search class, which serves as a wrapper around the FRED API
to retrieve and process economic data. It handles API communication, data transformation,
and formatting for compatibility with Excel and JSON serialization.

The Search class queries FRED endpoints for:
  - Economic data categories and subcategories
  - Available time-series data within categories
  - Historical observations for specific series
  - Metadata about economic indicators

All data is processed into pandas DataFrames for flexible transformation and exported
as JSON-compatible formats to ensure seamless integration with Excel add-ins.

Dependencies:
  - requests: HTTP library for FRED API communication
  - pandas (pd): Data manipulation and DataFrame operations
  - numpy (np): Numerical computing utilities
"""
import requests
import pandas as pd
import numpy as np

class Search():
    """
    FRED API Wrapper for querying and retrieving Federal Reserve economic data.
    
    Provides methods to retrieve economic data categories, series listings, historical observations,
    and metadata from the FRED API. Transforms raw API responses into pandas DataFrames
    and JSON-compatible formats for integration with Excel and web clients.
    
    Attributes:
        api_key (str): FRED API authentication key for requests.
    """
    def __init__(self, api_key):
        """
        Initialize the Search class with a FRED API key.
        
        Args:
            api_key (str): The API key for authenticating requests to the FRED API.
        """
        self.api_key = api_key

    def get_subcategories(self, category_id=0):
        """
        Retrieve child categories for a given FRED category ID.
        
        Queries the FRED API for all subcategories belonging to the specified category.
        Returns raw JSON response from the API.
        
        Args:
            category_id (int, optional): The FRED category ID to query. Defaults to 0 (root).
            
        Returns:
            dict: JSON response containing category hierarchy and metadata.
        """
        url = (
            "https://api.stlouisfed.org/fred/category/children"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json"
        )
        return requests.get(url).json()
    
    def get_series(self, category_id):
        """
        Retrieve all economic time series within a specified FRED category.
        
        Queries the FRED API to obtain a list of all available data series for a given category.
        Returns raw JSON response from the API.
        
        Args:
            category_id (int): The FRED category ID to query.
            
        Returns:
            dict: JSON response containing series information and metadata.
        """
        url = (
            "https://api.stlouisfed.org/fred/category/series"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json"
        )
        return requests.get(url).json()
    
    def get_series_df(self, category_id):
        """
        Retrieve all economic series in a category and return as a formatted DataFrame.
        
        Queries FRED for series within a category, extracts relevant metadata fields,
        and returns a pandas DataFrame with standardized columns for series information.
        
        Args:
            category_id (int): The FRED category ID to query.
            
        Returns:
            pd.DataFrame: DataFrame with columns [series_id, title, frequency, units, seasonal_adjustment].
        """
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
    
    def get_category_id(self, category_id=0):
        """
        Retrieve all categories and subcategories and return as a formatted DataFrame.
        
        Recursively queries FRED for subcategories and formats them into a flat DataFrame
        with category names and IDs. Data structure is compatible with pandas operations
        and JSON serialization for Excel integration.
        
        Args:
            category_id (int, optional): The FRED category ID to query. Defaults to 0 (root).
            
        Returns:
            pd.DataFrame: DataFrame with columns [category, category_id].
        """
        selection_df = []

        for lst in self.get_subcategories(category_id).values():
            for dictionary in lst:
                selection_df.append({
                    "category": dictionary["name"],
                    "category_id": dictionary["id"]
                    })
                
        return pd.DataFrame(selection_df)
    
    def get_data(self, series_id):
        """
        Retrieve historical observations for a specific FRED economic series.
        
        Fetches all available historical data points (dates and values) for a given series ID
        from the FRED API. Converts dates to datetime format and numeric values for analysis.
        Ensures Excel and JSON compatibility by replacing NaN values with None.
        
        Args:
            series_id (str): The FRED series ID (e.g., 'GDP', 'UNRATE').
            
        Returns:
            pd.DataFrame: DataFrame with columns [date, value] containing historical observations.
                         Missing values are represented as None for JSON serialization.
        """
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

    def get_series_info(self, series_id):
        """
        Retrieve comprehensive metadata for a specific FRED economic series.
        
        Queries FRED API for detailed information about a series including title, frequency,
        units, and seasonal adjustment metadata. Returns data as a list of dictionaries
        for JSON and Excel compatibility.
        
        Args:
            series_id (str): The FRED series ID to lookup (e.g., 'GDP', 'UNRATE').
            
        Returns:
            list: List of dictionaries containing series metadata with keys:
                  [series_id, title, frequency, units, seasonal_adjustment].
        """
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