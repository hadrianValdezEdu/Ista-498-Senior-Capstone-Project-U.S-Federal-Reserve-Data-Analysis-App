"""
REST API Proxy Backend for U.S. Federal Reserve Data Analysis Application.

This module implements a FastAPI-based REST API that serves as a secure proxy to the 
Federal Reserve Economic Data (FRED) API. The API enables Excel add-ins and web clients to:
  - Query economic data categories and subcategories
  - Retrieve available economic series within each category
  - Access historical time-series data for specific economic indicators
  - Fetch metadata and details about economic data series
  - Search for FRED series via full URLs or raw Series IDs

The application uses CORS middleware to facilitate cross-origin requests,
allowing seamless integration between the backend API and Excel-based data analysis tools.

Modules:
  - FastAPI & HTTPException: Web framework for building REST APIs and handling errors
  - CORSMiddleware: Enables cross-origin resource sharing for client compatibility
  - requests: HTTP library for catching underlying FRED API communication errors
  - Search: Custom wrapper class for querying and retrieving FRED data
"""
# --------------------------------------------------------------------------
# IMPORTS
# --------------------------------------------------------------------------
from fastapi import FastAPI
from fastapi import HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import requests # Import requests to catch its exceptions
from backend.search import Search
import os
from dotenv import load_dotenv

# --------------------------------------------------------------------------
# APP INITIALIZATION
# --------------------------------------------------------------------------

"""
Initialize the FastAPI application instance that serves as the web server proxy.
This app handles all HTTP requests, validates environment variables for API 
authentication, and routes them to appropriate endpoint handlers.
"""
app = FastAPI(title="FRED API Proxy", description="Proxy for FRED API to fetch economic data.")

load_dotenv()
FRED_API_KEY = os.environ.get("FRED_API_KEY", "")
if not FRED_API_KEY:
    print("WARNING: FRED_API_KEY not found in environment variables or .env file.")

# Initialize Search class instance with the loaded API key for FRED API communication
search = Search(api_key=FRED_API_KEY)

"""
Configure CORS (Cross-Origin Resource Sharing) middleware to enable the Excel add-in
and web clients to communicate with this API proxy without being blocked by browser security policies.
Allows specific local origins, methods, and headers to ensure full compatibility with the frontend.
"""
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "https://127.0.0.1:3000", "http://localhost:3000", "http://127.0.0.1:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------------------------------------------------------------------------
# CATEGORY ENDPOINTS
# --------------------------------------------------------------------------

"""
GET ENDPOINT 1: Retrieve subcategories for a given category ID.
URL Example: http://localhost:8080/categories/0
This endpoint queries the FRED database through the Search wrapper and returns all 
subcategories within a specified category. The response is converted to JSON format 
for consumption by Excel clients. Includes error handling for bad HTTP responses.
"""
@app.get("/categories/{category_id}")
def categories(category_id: int):
    try:
        df = search.get_category_id(category_id)
        return df.to_dict(orient="records")
    except requests.exceptions.HTTPError as e:
        print(f"Error fetching children for category {category_id}: {e}")
        raise HTTPException(status_code=e.response.status_code, detail=f"FRED API Error for category '{category_id}': {e.response.text}")
    except Exception as e:
        print(f"Unexpected error fetching children for category {category_id}: {e}")
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred while fetching children for category '{category_id}'.")

# --------------------------------------------------------------------------
# SERIES ENDPOINTS
# --------------------------------------------------------------------------

"""
GET ENDPOINT 2: Retrieve all economic data series within a specified category.
URL Example: http://localhost:8080/series/125?order_by=popularity
This endpoint returns a list of all available economic time series that belong to the given category.
Results are formatted as JSON records for Excel compatibility and can be dynamically 
ordered via query parameters.
"""
@app.get("/series/{category_id}")
def series(category_id: int, order_by: str = Query("popularity")):
    try:
        df = search.get_series_df(category_id, order_by=order_by)
        return df.to_dict(orient="records")
    except requests.exceptions.HTTPError as e:
        print(f"Error fetching series for category {category_id}: {e}")
        raise HTTPException(status_code=e.response.status_code, detail=f"FRED API Error for category '{category_id}': {e.response.text}")
    except Exception as e:
        print(f"Unexpected error fetching series for category {category_id}: {e}")
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred while fetching series for category '{category_id}'.")

# --------------------------------------------------------------------------
# SERIES INFO ENDPOINTS (for displaying details without data)
# --------------------------------------------------------------------------

"""
GET ENDPOINT 3: Retrieve metadata and information for a specific economic series ID.
URL Example: http://localhost:8080/info/GDP
This endpoint provides detailed metadata about a FRED series, including title, description,
frequency, units, and other attributes relevant for understanding the economic indicator,
without fetching the heavy historical data array. Serialized for seamless client integration.
"""
@app.get("/info/{series_id}")
def series_info(series_id: str):
    try:
        info = search.get_series_info(series_id)
        return info
    except requests.exceptions.HTTPError as e:
        print(f"Error fetching info for series {series_id}: {e}")
        raise HTTPException(status_code=e.response.status_code, detail=f"FRED API Error for series info '{series_id}': {e.response.text}")
    except Exception as e:
        print(f"Unexpected error fetching info for series {series_id}: {e}")
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred while fetching info for series '{series_id}'.")

# --------------------------------------------------------------------------
# DATA ENDPOINTS
# --------------------------------------------------------------------------

"""
GET ENDPOINT 4: Retrieve historical economic data for a specific FRED series ID.
URL Example: http://localhost:8080/data/GDP
This endpoint fetches the complete time-series data (dates and values) for a given economic indicator.
To prevent 500 errors on massive series, it safely looks up the observation_start date first.
The data is serialized as JSON records for seamless integration with Excel and other client applications.
"""
@app.get("/data/{series_id}")
def data(series_id: str):
    try:
        # To prevent 500 errors on large series, we get the observation_start date first.
        info = search.get_series_info(series_id)
        observation_start = None
        if info: # Check if info list is not empty
            observation_start = info[0].get('observation_start')

        df = search.get_data(series_id, observation_start=observation_start)
        return df.to_dict(orient="records")
    except requests.exceptions.HTTPError as e:
        print(f"Error fetching data for series {series_id}: {e}")
        raise HTTPException(status_code=e.response.status_code, detail=f"FRED API Error for series '{series_id}': {e.response.text}")
    except Exception as e:
        print(f"Unexpected error fetching data for series {series_id}: {e}")
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred while fetching data for series '{series_id}'.")

# --------------------------------------------------------------------------
# SEARCH ENDPOINTS
# --------------------------------------------------------------------------

"""
GET ENDPOINT 5: Search and retrieve data and metadata from user input.
URL Example: http://localhost:8080/logic/search?q=https://fred.stlouisfed.org/series/GDP
This endpoint parses flexible user input (either a full FRED URL or a raw Series ID),
extracts the valid ID, and fetches both the series metadata and historical observations.
Packages both responses into a combined JSON dictionary for flexible client-side rendering
and Excel integration.
"""
@app.get("/logic/search")
def search_logic(q: str = Query(..., description="Search query for a FRED series ID or URL")):
    series_id = search.parse_search_input(q)
    
    if series_id:
        try:
            info = search.get_series_info(series_id)
            observation_start = None
            if info: # Check if info list is not empty
                # Pass observation_start to prevent 500 errors on large series like GDP
                observation_start = info[0].get('observation_start')

            data_df = search.get_data(series_id, observation_start=observation_start)
            return {"type": "series", "info": info, "data": data_df.to_dict(orient="records")}
        except requests.exceptions.HTTPError as e:
            # Catch HTTP errors from FRED API calls
            print(f"Error fetching series {series_id}: {e}")
            raise HTTPException(status_code=e.response.status_code, detail=f"FRED API Error for series '{series_id}': {e.response.text}")
        except Exception as e: # Catch other potential errors (e.g., JSON parsing)
            print(f"Unexpected error fetching series {series_id}: {e}")
            raise HTTPException(status_code=500, detail=f"An unexpected error occurred while fetching series '{series_id}'. Details: {e}")
            
    raise HTTPException(status_code=404, detail="Invalid input. Please provide a valid FRED URL or Series ID.")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8080)
