from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from search import Search

# creates FastAPI web server for this backend
app = FastAPI()

# Initialize Search class
search = Search(api_key="ENTER YOUR API KEY HERE") # create an instance of the Search class, which will be used to call the methods in the Search class

# Allow Excel add-in to call this server
# without this, excel would block everything
app.add_middleware(
    CORSMiddleware,     # This allows what can access this api
    allow_origins=["*"],# This allows everything to access, without this line excel blocks request
    allow_methods=["*"],# Allows all HTTP methods
    allow_headers=["*"],# Allows headers in excel
)

# These are get end points. They act as a URL on our middleware server. When the excel
# add in calls this URL, it will trigger the functions below from our Search() class.

# GET ENDPOINT 1 — Return subcategories for a category ID
# Example: http://localhost:5000/categories/0
# Excel will go to this URL and trigger the function below, which will call the class method
@app.get("/categories/{category_id}")
def categories(category_id: int):
    df = search.get_category_id(category_id) # call the Search() class method
    return df.to_dict(orient="records")      # convert pandas df  into JSON friendly format for excel to read

# GET ENDPOINT 2 — Return all series inside a category
# Example: http://localhost:5000/series/125
@app.get("/series/{category_id}")
def series(category_id: int):
    df = search.get_series_df(category_id) # call the Search() class method
    return df.to_dict(orient="records")    # convert pandas df into JSON friendly format for excel to read

# GET ENDPOINT 3 — Return all data for a series ID
# Example: http://localhost:5000/data/GDP
@app.get("/data/{series_id}")
def data(series_id: str):
    df = search.get_data(series_id)        # call the Search() class method
    return df.to_dict(orient="records")    # convert pandas df into JSON friendly format for excel to read