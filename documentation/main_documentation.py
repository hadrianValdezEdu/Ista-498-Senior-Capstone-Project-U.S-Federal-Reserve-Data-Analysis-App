"""
REST API Backend for U.S. Federal Reserve Data Analysis Application.

This module implements a FastAPI-based REST API that provides access to Federal Reserve Economic Data (FRED)
through multiple endpoints. The API enables Excel add-ins and other client applications to:
  - Query economic data categories and subcategories
  - Retrieve available economic series within each category
  - Access historical time-series data for specific economic indicators
  - Fetch metadata and details about economic data series

The application uses CORS middleware to facilitate cross-origin requests from Excel clients,
allowing seamless integration between the backend API and Excel-based data analysis tools.

Modules:
  - FastAPI: Web framework for building REST APIs
  - CORSMiddleware: Enables cross-origin resource sharing for Excel client compatibility
  - Search: Custom class for querying and retrieving FRED data
"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from search import Search

"""
Initialize the FastAPI application instance that serves as the web server backend.
This app handles all HTTP requests and routes them to appropriate endpoint handlers.
"""
app = FastAPI()

# Initialize Search class instance with empty API key (to be configured at runtime)
search = Search(api_key="")

"""
Configure CORS (Cross-Origin Resource Sharing) middleware to enable the Excel add-in
to communicate with this API backend without being blocked by browser security policies.
Allows all origins, methods, and headers to ensure full compatibility with Excel clients.
"""
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


"""
GET ENDPOINT 1: Retrieve subcategories for a given category ID.
URL Example: http://localhost:5000/categories/0
This endpoint queries the FRED database and returns all subcategories within a specified category.
The response is converted to JSON format for consumption by Excel clients.
"""
@app.get("/categories/{category_id}")
def categories(category_id: int):
    df = search.get_category_id(category_id)
    return df.to_dict(orient="records")

"""
GET ENDPOINT 2: Retrieve all economic data series within a specified category.
URL Example: http://localhost:5000/series/125
This endpoint returns a list of all available economic time series that belong to the given category.
Results are formatted as JSON records for Excel compatibility.
"""
@app.get("/series/{category_id}")
def series(category_id: int):
    df = search.get_series_df(category_id)
    return df.to_dict(orient="records")

"""
GET ENDPOINT 3: Retrieve historical economic data for a specific FRED series ID.
URL Example: http://localhost:5000/data/GDP
This endpoint fetches the complete time-series data (dates and values) for a given economic indicator.
The data is serialized as JSON records for integration with Excel and other client applications.
"""
@app.get("/data/{series_id}")
def data(series_id: str):
    df = search.get_data(series_id)
    return df.to_dict(orient="records")

"""
GET ENDPOINT 4: Retrieve metadata and information for a specific economic series ID.
URL Example: http://localhost:5000/search/GDP
This endpoint provides detailed metadata about a FRED series, including title, description,
frequency, units, and other attributes relevant for understanding the economic indicator.
"""
@app.get("/search/{series_id}")
def search_series(series_id: str):
    data = search.get_series_info(series_id)
    return data