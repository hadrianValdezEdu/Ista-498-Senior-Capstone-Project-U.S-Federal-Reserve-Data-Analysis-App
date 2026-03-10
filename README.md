Hadrian's Note:

March 10 I added a working back end server

Excel Add‑in (JavaScript UI) ***SET UP BUT NOT FINISHED***
        ↓
HTTP Requests
        ↓
Python Backend (FastAPI) ***I FINISHED THIS PART***
        ↓
Your Search Class ***FINISHED EARLIER***
        ↓
FRED API ***GIVEN***

When our project is finished. When a user selects a button from our Excel add in UI, the 
button calls the python backend which is a HTTP server, and that server calls our python
code's functions which in turn fetchs data from Fred's Api.