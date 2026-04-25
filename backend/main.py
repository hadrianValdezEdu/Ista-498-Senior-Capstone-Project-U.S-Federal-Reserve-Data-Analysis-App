# --------------------------------------------------------------------------
# IMPORTS
# --------------------------------------------------------------------------
from fastapi import FastAPI
from fastapi import HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import requests # Import requests to catch its exceptions
from search import Search

# --------------------------------------------------------------------------
# APP INITIALIZATION
# --------------------------------------------------------------------------

app = FastAPI(title="FRED API Proxy", description="Proxy for FRED API to fetch economic data.")
search = Search(api_key="")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "https://127.0.0.1:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --------------------------------------------------------------------------
# CATEGORY ENDPOINTS
# --------------------------------------------------------------------------

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

@app.get("/series/{category_id}")
def series(category_id: int):
    try:
        df = search.get_series_df(category_id)
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
