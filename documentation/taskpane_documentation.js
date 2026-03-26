/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/* global document, Office, Excel */

// ===========================
// GLOBAL STATE VARIABLES
// ===========================

// selectedCategoryId: Stores the ID of the currently selected category from the FRED API.
// Used to track which category the user has selected when browsing the category tree.
// Type: String (category ID) or null (if no selection made)
// Initial Value: null
let selectedCategoryId = null;

// selectedSeriesId: Stores the ID of the currently selected time series from the FRED API.
// Used to track which series the user has selected before fetching data.
// Type: String (series ID) or null (if no selection made)
// Initial Value: null
let selectedSeriesId = null;

// ===========================
// OFFICE INITIALIZATION
// ===========================

// Office.onReady(callback):
//   Called when the Office JavaScript API has finished loading and is ready to use.
//   This ensures the Excel API is available before attempting to interact with it.
//   Param: callback - Function to execute once Office is ready (in this case, an arrow function that sets up button handlers)
//
// This initialization block performs the critical task of attaching click event listeners
// to the three main buttons in the task pane. Without these handlers, button clicks would not
// trigger any functionality.
Office.onReady(() => {
  // -----
  // BUTTON 1: "Get Categories" Button Setup
  // -----
  // This button fetches the ROOT-level categories from the FRED API (starting at category ID 0).
  //
  // - document.getElementById("btnGetCategories"): Retrieves the HTML element with ID "btnGetCategories" from the DOM
  //   The DOM (Document Object Model) is the programming interface for HTML documents running in the Excel task pane.
  //   The task pane is essentially a mini web browser window embedded within Excel.
  //
  // - .onclick = () => getSubcategories(0): Assigns an event listener (a function) to the button's click event.
  //   When the user clicks the button, the function getSubcategories(0) is executed.
  //   The parameter 0 represents the root category ID in the FRED API, so this fetches top-level categories.
  //   URL example: https://fred.stlouisfed.org/categories (without specifying a parent category)
  document.getElementById("btnGetCategories").onclick = () => {getSubcategories(0)};

  // -----
  // BUTTON 2: "Get Series" Button Setup
  // -----
  // This button fetches all time series belonging to the category selected by the user.
  // It uses the global variable selectedCategoryId to determine which category to query.
  //
  // - document.getElementById("btnGetSeries"): Retrieves the button element from the DOM
  // - .onclick = getSeries: Assigns the getSeries function as the click handler
  //   Note: We pass the function reference directly (not calling it with ()) so it executes on click, not immediately.
  document.getElementById("btnGetSeries").onclick = getSeries;

  // -----
  // BUTTON 3: "Get Data" Button Setup
  // -----
  // This button fetches time series data for the series selected by the user and inserts it into Excel.
  // It uses the global variable selectedSeriesId to determine which series to query.
  //
  // - document.getElementById("btnGetData"): Retrieves the button element from the DOM
  // - .onclick = getData: Assigns the getData function as the click handler
  //   Again, we pass the function reference so it executes when clicked.
  document.getElementById("btnGetData").onclick = getData;
});

// ===========================
// GET CATEGORIES FUNCTION
// ===========================

// getSubcategories(categoryId):
//   Fetches a list of subcategories for a given parent category from the FRED API via the backend.
//   This function is asynchronous (uses async/await) to handle the network request without blocking the UI.
//
//   Parameters:
//     - categoryId (string): The ID of the parent category. 
//       Special case: 0 = root category (no parent), fetches top-level categories.
//       Example: "13" fetches subcategories of category 13 (Manufacturing).
//
//   Return Value: None (void). Results are displayed in the task pane UI via renderCategories().
//
//   Process:
//     1. Displays a loading message in the task pane UI
//     2. Makes an HTTP GET request to the backend endpoint: http://localhost:5000/categories/{categoryId}
//     3. Parses the JSON response from the backend
//     4. Calls renderCategories() to display the categories in the task pane
//     5. Catches and logs any errors to console
//
async function getSubcategories(categoryId) {
  // Display a loading message to provide user feedback that the app is fetching data
  // document.getElementById("output"): Gets the HTML element with ID "output" from the DOM
  // .innerHTML: Sets the HTML content inside this element (replaces any existing content)
  document.getElementById("output").innerHTML = "Fetching categories...";

  try {
    // Make an HTTP GET request to the backend server for the specified category ID.
    // The backend (FastAPI server in main.py) has an endpoint "/categories/{category_id}"
    // that handles this request and communicates with the FRED API.
    //
    // await: Pauses execution until the fetch request completes and returns a response.
    //        Without await, the code would continue before the response arrives, causing errors.
    //
    // Template literal (backticks): Allows ${categoryId} to be replaced with the actual category ID value.
    //        Example: If categoryId = 0, the URL becomes http://localhost:5000/categories/0
    //
    // Response: A Promise that resolves to a Response object containing the HTTP response.
    const response = await fetch(
      `http://localhost:5000/categories/${categoryId}`
    );

    // Parse the JSON response body into a JavaScript object or array
    // response.json(): Returns a Promise that resolves to the parsed JSON data.
    // await: Waits for the JSON parsing to complete before assigning to the data variable.
    //
    // data: Contains an array of category objects from the FRED API, each with:
    //       - category_id: The unique identifier for the category
    //       - category: The display name of the category
    //       - description: (optional) A description of what the category contains
    const data = await response.json();

    // Call the renderCategories function to display the fetched categories in the task pane
    // This function takes the parsed data and generates HTML to present it to the user.
    // We also pass categoryId so the UI knows which category level we're viewing.
    renderCategories(data, categoryId);

    // Catch any errors that occur during the fetch or JSON parsing
    // Errors could include: network failure, invalid JSON response, server error, etc.
  } catch (error) {
    // Display an error message in the task pane output area
    document.getElementById("output").innerHTML = "Error fetching categories.";

    // Log the full error object to the browser console for debugging
    // This helps developers understand what went wrong (network error, parsing error, etc.)
    console.error(error);
  }
}

// renderCategories(categories, parentCategoryId):
//   Generates HTML to display a list of categories and adds click handlers to enable category navigation.
//   This function is synchronous (does not use async/await) because it only manipulates the DOM, not network requests.
//
//   Parameters:
//     - categories (array): An array of category objects from the FRED API. Each object has:
//       - category_id: The unique identifier (number)
//       - category: The display name (string)
//       - description: (optional) Descriptive text
//
//     - parentCategoryId (string/number): The ID of the parent category level. Used for display only.
//       Helps the user understand where they are in the category tree hierarchy.
//
//   Return Value: None (void). Modifies the DOM by setting innerHTML and adding event listeners.
//
//   Side Effects:
//     - Updates the output element HTML
//     - Adds click event listeners to category items
//     - Sets the global selectedCategoryId when a category is clicked
//
function renderCategories(categories, parentCategoryId) {
  // Get reference to the output HTML element where categories will be displayed
  // This element is defined in taskpane.html with id="output"
  const output = document.getElementById("output");

  // Initialize the HTML string that will be built up with category list items
  // Template literal allows us to include the parentCategoryId value in the heading
  // <h3>...<h3>: HTML heading element for the section title
  // <ul>: HTML unordered list opening tag. Categories will be rendered as list items inside.
  let html = `<h3>Select a Category (Parent ID: ${parentCategoryId})</h3><ul>`;

  // forEach(callback): Iterates over each category in the categories array
  // cat: Parameter representing the current category object in the iteration
  // This loops once for each category returned from the FRED API
  categories.forEach(cat => {
    // Build an HTML list item for the current category
    // Template literal allows us to embed the category properties dynamically:
    //   - ${cat.category_id}: The unique identifier from the API response
    //   - ${cat.category}: The display name from the API response
    //
    // data-id attribute: Custom HTML attribute that stores the category ID
    //   Used later to identify which category was clicked when adding event listeners
    //
    // class="category-item": CSS class assigned for styling and for selecting these elements via querySelectorAll
    html += `
      <li class="category-item" data-id="${cat.category_id}">
        ${cat.category} (ID: ${cat.category_id})
      </li>`;
  });

  // Close the unordered list by adding the closing </ul> tag
  // This completes the HTML structure
  html += "</ul>";

  // Add a button that allows the user to view series for the CURRENT category level
  // This gives the user two options: explore subcategories OR jump directly to series data
  // Template literal embeds the parentCategoryId so the button label is informative
  html += `
    <button id="btnViewSeries">View Series in Category ${parentCategoryId}</button>
  `;

  // Set the HTML content of the output element
  // output.innerHTML replaces all existing content with the newly generated HTML
  // This causes the categories list to appear in the task pane for the user to see
  output.innerHTML = html;

  // Add click event handlers to each category list item for interactive navigation
  // querySelectorAll(".category-item"): Returns a NodeList of all elements with class "category-item"
  //   These are the <li> elements we just created in the loop above
  //
  // forEach(item => {...}): Iterates over each category item element
  //   item: Parameter representing the current DOM element being processed
  document.querySelectorAll(".category-item").forEach(item => {
    // item.onclick: Assigns a click event handler to this specific category item
    // When the user clicks on a category in the task pane, this function executes
    item.onclick = () => {
      // Retrieve the category ID from the "data-id" attribute of the clicked element
      // .getAttribute("data-id"): Gets the value of the data-id attribute
      // This ID was set when we generated the HTML above with data-id="${cat.category_id}"
      //
      // selectedCategoryId: Global variable updated to reflect the user's selection
      // Used by other functions (getSeries, getData) to know which category to work with
      selectedCategoryId = item.getAttribute("data-id");

      // Automatically fetch and display the subcategories of the clicked category
      // This enables full category-tree navigation, allowing users to drill down into categories
      // Passing selectedCategoryId (just set above) as the parameter for the next level
      getSubcategories(selectedCategoryId);
    };
  });

  // Add click handler for the "View Series" button
  // This allows the user to fetch series data for the CURRENT category level without going deeper
  // document.getElementById("btnViewSeries"): Gets the button element we added above
  document.getElementById("btnViewSeries").onclick = () => {
    // Update the global selectedCategoryId to the current level (parentCategoryId)
    // This tells getSeries() which category to fetch series from
    selectedCategoryId = parentCategoryId;

    // Fetch and display the series for the current category level
    getSeries();
  };
}

// ===========================
// GET SERIES FOR SELECTED CATEGORY
// ===========================

// getSeries():
//   Fetches all time series belonging to the currently selected category from the FRED API via the backend.
//   This is an asynchronous function that handles network requests without blocking the UI.
//   Requires that the user has previously selected a category (stored in global selectedCategoryId).
//
//   Parameters: None (uses the global variable selectedCategoryId)
//
//   Return Value: None (void). Results are displayed in the task pane UI via renderSeries().
//
//   Validations:
//     - Checks if selectedCategoryId is set; displays error if not
//
//   Process:
//     1. Validates that a category has been selected
//     2. Displays a loading message in the task pane
//     3. Makes an HTTP GET request to the backend: http://localhost:5000/series/{categoryId}
//     4. Parses the JSON response (array of series objects)
//     5. Calls renderSeries() to display the series in the task pane
//     6. Catches and logs any errors
//
async function getSeries() {
  // Validate that the user has selected a category before attempting to fetch series
  // !selectedCategoryId: Checks if selectedCategoryId is null, undefined, empty string, 0, false, or any other falsy value
  // If selectedCategoryId is falsy, it means the user hasn't selected a category yet
  if (!selectedCategoryId) {

    // Inform the user that they need to select a category first before fetching series
    document.getElementById("output").innerHTML = "Please select a category first.";

    // Exit the function immediately without making a backend request
    // This prevents errors that would occur if we tried to fetch series with a null/undefined category ID
    return;
  }

  // Display a loading message to provide user feedback that the app is fetching series data
  document.getElementById("output").innerHTML = "Fetching series for selected category...";

  try {
    // Make an HTTP GET request to the backend for the series in the selected category
    // The backend endpoint "/series/{category_id}" handles this request and calls the FRED API
    // selectedCategoryId: The global variable set when the user clicks a category in renderCategories()
    //
    // await: Waits for the fetch request to complete before proceeding
    const response = await fetch(
      `http://localhost:5000/series/${selectedCategoryId}`
    );

    // Parse the JSON response into a JavaScript object/array
    // The parsed data should be an array of series objects from the FRED API, each with:
    //   - series_id: The unique identifier for the series (e.g., "GDP", "UNRATE")
    //   - title: The full descriptive name of the series (e.g., "Real Gross Domestic Product")
    //   - units: The measurement units (e.g., "Billions of Dollars")
    //   - frequency: How often the data is updated (e.g., "Quarterly")
    const data = await response.json();

    // Pass the parsed series data to renderSeries() for display in the task pane
    // renderSeries() will generate HTML and add click handlers for series selection
    renderSeries(data);

    // Catch any errors occurring during fetch or JSON parsing
  } catch (error) {
    // Display an error message in the task pane output
    document.getElementById("output").innerHTML = "Error fetching series.";

    // Log the error to the console for debugging purposes
    console.error(error);
  }
}
// renderSeries(seriesList):
//   Generates HTML to display a list of time series and adds click handlers to enable series selection.
//   This is a synchronous function that only manipulates the DOM.
//
//   Parameters:
//     - seriesList (array): An array of series objects from the FRED API. Each object contains:
//       - series_id: The unique identifier (e.g., "GDP")
//       - title: The descriptive name (e.g., "Real Gross Domestic Product")
//       - units: Measurement units (e.g., "Billions of Dollars")
//       - frequency: Update frequency (e.g., "Quarterly")
//
//   Return Value: None (void). Modifies the DOM and sets the global selectedSeriesId when a series is clicked.
//
//   Side Effects:
//     - Updates the output element HTML with the series list
//     - Adds click event listeners to each series item
//     - Sets global selectedSeriesId when a series is clicked
//
function renderSeries(seriesList) {
  // Get reference to the output HTML element where series will be displayed
  const output = document.getElementById("output");

  // Initialize the HTML string with a heading and opening unordered list tag
  // This heading is different from categories - it prompts the user to select a series
  let html = "<h3>Select a Series</h3><ul>";

  // Iterate through each series in the seriesList array
  // series: Parameter representing the current series object in the iteration
  seriesList.forEach(series => {
    // Build an HTML list item for the current series
    // Template literal embeds the series properties:
    //   - ${series.series_id}: The unique identifier (e.g., "GDP")
    //   - ${series.title}: The descriptive name of the series
    //
    // data-id attribute: Custom HTML attribute that stores the series ID
    //   Used when adding click handlers to identify which series was selected
    //
    // class="series-item": CSS class for styling and element selection
    html += `
      <li class="series-item" data-id="${series.series_id}">
        ${series.title} (ID: ${series.series_id})
      </li>`;
  });

  // Close the unordered list tag
  html += "</ul>";

  // Set the HTML content of the output element
  // This displays the series list in the task pane for the user to interact with
  output.innerHTML = html;

  // Add click event handlers to each series list item for series selection
  // querySelectorAll(".series-item"): Returns a NodeList of all elements with class "series-item"
  // forEach(item => {...}): Iterates over each series item element
  document.querySelectorAll(".series-item").forEach(item => {
    // item.onclick: Assigns a click event handler to this specific series item
    // When the user clicks on a series, this function executes
    item.onclick = () => {
      // Retrieve the series ID from the "data-id" attribute of the clicked element
      // .getAttribute("data-id"): Gets the value of the data-id attribute
      // This ID was set when we generated the HTML above
      //
      // selectedSeriesId: Global variable updated to reflect the user's series selection
      // Used by getData() to know which series to fetch data for
      selectedSeriesId = item.getAttribute("data-id");

      // Display a confirmation message showing the selected series ID
      // This provides visual feedback to the user that their selection was registered
      output.innerHTML = `<p>Selected Series ID: ${selectedSeriesId}</p>`;
    };
  });
}

// ===========================
// GET DATA FOR SELECTED SERIES + INSERT INTO EXCEL
// ===========================

// getData():
//   Fetches time series data for the currently selected series from the FRED API via the backend,
//   then inserts that data into the active Excel worksheet.
//   This is an asynchronous function that combines network requests and Excel API interactions.
//
//   Parameters: None (uses the global variable selectedSeriesId)
//
//   Return Value: None (void). Results are inserted directly into Excel cells.
//
//   Validations:
//     - Checks if selectedSeriesId is set; displays error if not
//
//   Process:
//     1. Validates that a series has been selected
//     2. Displays a loading message in the task pane
//     3. Makes an HTTP GET request to the backend: http://localhost:5000/data/{seriesId}
//     4. Parses the JSON response (array of date/value objects)
//     5. Calls insertDataIntoExcel() to insert the data into Excel
//     6. Displays a success message when complete
//     7. Catches and logs any errors
//
// Called By: The "Get Data" button in the task pane (set up in Office.onReady())
async function getData() {
  // Validate that the user has selected a series before attempting to fetch data
  // If selectedSeriesId is falsy (null, undefined, etc.), display an error
  if (!selectedSeriesId) {
    document.getElementById("output").innerHTML =
      "Please select a series first.";
    return;
  }

  // Display a loading message to provide user feedback
  document.getElementById("output").innerHTML =
    "Fetching data for selected series...";

  try {
    // Make an HTTP GET request to the backend for the selected series data
    // The backend endpoint "/data/{series_id}" handles this request and calls the FRED API
    // selectedSeriesId: The global variable set when the user clicks a series in renderSeries()
    //
    // await: Waits for the fetch request to complete
    const response = await fetch(
      `http://localhost:5000/data/${selectedSeriesId}`
    );

    // Parse the JSON response into a JavaScript object/array
    // The parsed data should be an array of objects, each containing:
    //   - date: The date of the data point (e.g., "2023-01-01")
    //   - value: The numeric value for that date (e.g., 27.5)
    const data = await response.json();

    // Call insertDataIntoExcel() to insert the fetched data into the active Excel worksheet
    // This function uses the Excel JavaScript API to write data into cells
    // Pass await so we wait for the Excel insertion to complete before showing the success message
    await insertDataIntoExcel(data);

    // Display a success message in the task pane after data insertion is complete
    document.getElementById("output").innerHTML =
      "Data inserted into Excel successfully.";

  // Catch any errors occurring during fetch, JSON parsing, or Excel insertion
  } catch (error) {
    // Display an error message in the task pane
    document.getElementById("output").innerHTML = "Error fetching data.";

    // Log the error to the console for debugging
    console.error(error);
  }
}

// ===========================

// insertDataIntoExcel(data):
//   Inserts time series data into the active Excel worksheet starting at cell A1.
//   Data is inserted in two columns: Column A (dates) and Column B (values).
//   The function uses the Excel JavaScript API to interact with Excel.
//
//   Parameters:
//     - data (array): An array of objects with the structure:
//       [
//         { date: "2023-01-01", value: 27.5 },
//         { date: "2023-02-01", value: 28.1 },
//         ...
//       ]
//
//   Return Value: None (void). Modifies Excel cells directly.
//
//   Process:
//     1. Uses Excel.run() to create a context for interacting with Excel
//     2. Gets the active worksheet
//     3. Converts the data array into a 2D array format (required by Excel)
//     4. Calculates the range needed (A1:Bn where n is the number of data rows)
//     5. Writes the data to the calculated range
//     6. Calls context.sync() to apply the changes
//
async function insertDataIntoExcel(data) {
  // Excel.run(callback):
  //   Creates a batching scope for executing Excel API operations.
  //   The callback receives a "context" object that represents the Excel application.
  //
  //   context: Provides access to the Excel object model:
  //     - context.workbook: The currently open workbook
  //     - context.workbook.worksheets: The sheets in the workbook
  //     - context.workbook.worksheets.getActiveWorksheet(): The sheet the user is currently viewing
  //
  //   All changes are batched and then sent to Excel together when context.sync() is called.
  //   This improves performance and prevents flickering in the UI.
  await Excel.run(async context => {
    // Get the active worksheet (the one currently displayed in Excel)
    // sheet: A reference to the Excel worksheet object
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Convert the data array into a 2D array format required by Excel
    // .map(item => [item.date, item.value]): Creates a new array where each element is a 2-element array [date, value]
    // This transforms the data from objects to nested arrays: [[date1, value1], [date2, value2], ...]
    // rows: A 2D array ready to be inserted into Excel cells
    const rows = data.map(item => [item.date, item.value]);

    // Calculate the range of cells where data will be inserted
    // Template literal creates a dynamic range string:
    //   - A1: Starting cell (column A, row 1)
    //   - B${rows.length}: Ending cell (column B, last row of data)
    //   Example: If there are 12 data points, range = "A1:B12"
    //
    // sheet.getRange(): Gets the specified range of cells from the worksheet
    // range: A reference to the Excel range object
    const range = sheet.getRange(`A1:B${rows.length}`);

    // Set the values of the range to the data from rows
    // range.values = rows: Assigns the 2D array of data to the Excel cells
    // This writes all the dates and values into columns A and B in one operation
    range.values = rows;

    // Send all batched operations to Excel and apply the changes
    // context.sync(): Communicates with Excel to execute all pending operations
    // await: Waits for the sync to complete before returning from the function
    // Without this call, the changes would not be reflected in Excel
    await context.sync();
  });
}
// ===========================
// SEARCH FUNCTIONALITY
// ===========================

// searchInput():
//   Performs a text search across FRED API series using the query entered by the user.
//   Fetches search results from the backend and displays them in the task pane.
//   This is an asynchronous function that handles network requests.
//
//   Parameters: None (retrieves search query from the HTML input field with id="searchInput")
//
//   Return Value: None (void). Results are displayed in the task pane via renderSeries().
//
//   Process:
//     1. Gets the search query from the text input field
//     2. Displays a loading message in the task pane
//     3. Makes an HTTP GET request to the backend: http://localhost:5000/search/{query}
//     4. Encodes the query for safe URL transmission using encodeURIComponent()
//     5. Parses the JSON response (array of matching series objects)
//     6. Calls renderSeries() to display the search results
//     7. Catches and logs any errors
//
// Called By: The "Search" button in the task pane (must be set up in HTML with onclick handler)
async function searchInput() {
  // Get the search query from the text input element
  // document.getElementById("searchInput"): Retrieves the input field from taskpane.html with id="searchInput"
  // .value: Gets the text that the user typed into the input field
  // query: Contains the search term (e.g., "unemployment", "GDP", etc.)
  const query = document.getElementById("searchInput").value;

  // Display a loading message showing what search is in progress
  // Template literal embeds the user's query for context
  document.getElementById("output").innerHTML = `Searching for "${query}"...`;

  try {
    // Make an HTTP GET request to the backend search endpoint
    // The backend endpoint "/search/{query}" handles this request and calls the FRED API search function
    //
    // encodeURIComponent(query): Encodes special characters in the search query to make it URL-safe
    //   Example: "US Economy" becomes "US%20Economy" (space becomes %20)
    //   This prevents URL parsing errors if the query contains special characters
    //
    // Template literal constructs the full URL with the encoded query
    const response = await fetch(
      `http://localhost:5000/search/${encodeURIComponent(query)}`
    );

    // Parse the JSON response into a JavaScript object/array
    // The parsed data should be an array of series objects matching the search query
    // Each object contains:
    //   - series_id: The unique identifier of the series
    //   - title: The series name or description
    //   - units: The measurement units
    //   - frequency: How often the data is updated
    const data = await response.json();

    // Pass the search results to renderSeries() for display
    // renderSeries() generates HTML to display the results as a clickable list
    // The user can then click on a result to select it for data fetching
    renderSeries(data);

  // Catch any errors occurring during fetch or JSON parsing
  } catch (error) {
    // Display an error message in the task pane
    document.getElementById("output").innerHTML = "Error performing search.";

    // Log the error to the console for debugging
    console.error(error);
  }
}