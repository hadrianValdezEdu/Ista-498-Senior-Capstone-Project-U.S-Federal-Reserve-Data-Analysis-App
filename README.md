# U.S. Federal Reserve Data Analysis App

**ISTA 498 - Senior Capstone Project**  

---

## What is this program?
This is a custom Microsoft Excel Add-in that connects directly to the Federal Reserve Economic Data (FRED) database. We're building this for our senior capstone project to make financial analysis a lot less tedious.

## What does it do?
If you've ever tried to analyze U.S. economic trends (like GDP or inflation), you know the normal process sucks. You have to go to the FRED website, manually download a bunch of CSV files, import them into Excel, and format everything before you can actually get to work. 

This app completely automates that. You just open the add-in inside Excel, search for the economic indicators you need, click a button, and it instantly dumps the historical data right into your spreadsheet.

## How does it do it?
Here's the step-by-step of how the program actually works under the hood:

1. **The UI (Excel/JavaScript):** You type your query into the custom search bar we built inside Excel.
2. **The Request:** When you hit "Search", the frontend fires off an HTTP request to our backend.
3. **The Server (Python/FastAPI):** We have a FastAPI server running locally that catches that request.
4. **The Search Logic:** The server hands the request to our custom Python `Search` class. This handles all the business logic and formats your search into a valid query.
5. **The API:** The `Search` class pings the official FRED API and grabs the raw, up-to-date data.
6. **The Return Trip:** That data is passed back through the FastAPI server straight to the Excel Add-in, which then writes the numbers into your active workbook.