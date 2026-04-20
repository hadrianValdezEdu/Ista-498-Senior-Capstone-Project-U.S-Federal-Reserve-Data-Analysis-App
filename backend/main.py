# --------------------------------------------------------------------------
# IMPORTS
# --------------------------------------------------------------------------

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from search import Search

# --------------------------------------------------------------------------
# APP INITIALIZATION
# --------------------------------------------------------------------------

app = FastAPI()
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
    df = search.get_category_id(category_id)
    return df.to_dict(orient="records")

# --------------------------------------------------------------------------
# SERIES ENDPOINTS
# --------------------------------------------------------------------------

@app.get("/series/{category_id}")
def series(category_id: int):
    df = search.get_series_df(category_id)
    return df.to_dict(orient="records")

# --------------------------------------------------------------------------
# DATA ENDPOINTS
# --------------------------------------------------------------------------

@app.get("/data/{series_id}")
def data(series_id: str):
    df = search.get_data(series_id)
    return df.to_dict(orient="records")

# --------------------------------------------------------------------------
# SEARCH ENDPOINTS
# --------------------------------------------------------------------------

@app.get("/logic/search/{series_id}")
def search_series(series_id: str):
    info = search.get_series_info(series_id)
    data_df = search.get_data(series_id)
    return {"info": info, "data": data_df.to_dict(orient="records")}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8080)
