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

Modules:
  - requests: HTTP library for FRED API communication
  - pandas (pd): Data manipulation and DataFrame operations
  - numpy (np): Numerical computing utilities
    - re: Regular expressions for parsing user input
"""
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
        """
        self.api_key = api_key

    # ----------------------------------------------------------------------
    # CATEGORY RETRIEVAL
    # ----------------------------------------------------------------------

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
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()

    def get_category_id(self, category_id=0):
        """
        Retrieve all categories and subcategories and return as a formatted DataFrame.
        
        Queries FRED for subcategories and formats them into a flat DataFrame
        with category names and IDs. Data structure is compatible with pandas operations
        and JSON serialization for Excel integration.
        
        Args:
            category_id (int, optional): The FRED category ID to query. Defaults to 0 (root).
            
        Returns:
            pd.DataFrame: DataFrame with columns [category, category_id].
        """
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

    def get_series(self, category_id, order_by="popularity"):
        """
        Retrieve all economic time series within a specified FRED category.
        
        Queries the FRED API to obtain a list of all available data series for a given category.
        Returns raw JSON response from the API.
        
        Args:
            category_id (int): The FRED category ID to query.
            order_by (str, optional): Metric to sort the series by. Defaults to 'popularity'.
            
        Returns:
            dict: JSON response containing series information and metadata.
        """
        url = (
            "https://api.stlouisfed.org/fred/category/series"
            f"?category_id={category_id}&api_key={self.api_key}&file_type=json&limit=1000&order_by={order_by}&sort_order=desc"
        )
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()

    def get_series_df(self, category_id, order_by="popularity"):
        """
        Retrieve all economic series in a category and return as a formatted DataFrame.
        
        Queries FRED for series within a category, extracts relevant metadata fields,
        and returns a pandas DataFrame with standardized columns for series information.
        
        Args:
            category_id (int): The FRED category ID to query.
            order_by (str, optional): Metric to sort the series by. Defaults to 'popularity'.
            
        Returns:
            pd.DataFrame: DataFrame with columns [series_id, title, frequency, units, seasonal_adjustment].
        """
        data = self.get_series(category_id, order_by=order_by)
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
        """
        Retrieve comprehensive metadata for a specific FRED economic series.
        
        Queries FRED API for detailed information about a series including title, frequency,
        units, and seasonal adjustment metadata. Returns data as a list of dictionaries
        for JSON and Excel compatibility.
        
        Args:
            series_id (str): The FRED series ID to lookup (e.g., 'GDP', 'UNRATE').
            
        Returns:
            list: List of dictionaries containing series metadata and information.
        """
        url = (
            "https://api.stlouisfed.org/fred/series"
            f"?series_id={series_id}&api_key={self.api_key}&file_type=json"
        )
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        data = response.json()

        series_list = data.get("seriess", [])

        # Process the list to rename 'id' to 'series_id' for frontend consistency
        # and return all available metadata fields.
        processed_list = []
        for s in series_list:
            info_copy = s.copy()
            if 'id' in info_copy:
                info_copy['series_id'] = info_copy.pop('id')
            processed_list.append(info_copy)

        return processed_list

    # ----------------------------------------------------------------------
    # DATA RETRIEVAL
    # ----------------------------------------------------------------------

    def get_data(self, series_id, observation_start=None):
        """
        Retrieve historical observations for a specific FRED economic series.
        
        Fetches all available historical data points (dates and values) for a given series ID
        from the FRED API. Converts dates to datetime format and numeric values for analysis.
        Ensures Excel and JSON compatibility by replacing NaN values with None.
        
        Args:
            series_id (str): The FRED series ID (e.g., 'GDP', 'UNRATE').
            observation_start (str, optional): Date string to limit payload size.
            
        Returns:
            pd.DataFrame: DataFrame with columns [date, value] containing historical observations.
        """
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
        df = df.replace({np.nan: None})

        return df

    # ----------------------------------------------------------------------
    # SEARCH PARSING LOGIC
    # ----------------------------------------------------------------------

    def parse_search_input(self, user_input):
        """
        Parse and sanitize flexible user input to extract a valid FRED Series ID.
        
        Handles both full FRED URL strings and raw Series IDs. Uses regex to accurately
        match patterns and strips out invalid characters to ensure clean API requests.
        
        Args:
            user_input (str): Raw input from the user (URL or ID).
            
        Returns:
            str: The sanitized and capitalized FRED Series ID.
        """
        user_input = user_input.strip()
        # This regex pattern is a nightmare to read
        url_match = re.search(r"fred\.stlouisfed\.org/series/([A-Za-z0-9_]+)", user_input, re.IGNORECASE)
        
        if url_match:
            return url_match.group(1).upper()
        else:
            # Only clean if it's not a URL match
            clean_id = re.sub(r'[^A-Za-z0-9_]', '', user_input)
            return clean_id.upper()