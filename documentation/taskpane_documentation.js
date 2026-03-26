/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */



/* global document, Office, Excel */

// Global State
// -----------------------------
let selectedCategoryId = null;
let selectedSeriesId = null;
// -----------------------------

// Office Initialization
// -----------------------------
Office.onReady(() => {
  // Wire the three buttons to their 3 functions

  // document.getElementById("btnGetCategories") is from the Document Object Model (DOM) API.
  // the DOM API allows JavaScript to interact with HTML elements on the task pane.
  // DOM API comes from webview which excel task panes are built on. the TP is basically a mini web page running inside excel.
  //
  // document.getElementById("btnGetCategories") gets the button element from the HTML page (taskpane.html)
  // by its ID and assigns it to the variable btnGetCategories.
  // = () => getSubcategories(0) sets the onclick event handler for the button. When the button is clicked
  // it will execute the function getSubcategories with an argument of 0.
  // 0 represents the root category in the FRED API, so this will fetch the top-level categories when the button is clicked.
  // it represnets a similar URL to this https://fred.stlouisfed.org/categories
  document.getElementById("btnGetCategories").onclick = () => {getSubcategories(0)};
  // .onclick listens for click events on the excel UI button. When the button is clicked, it calls the getCategories function defined below.
  document.getElementById("btnGetSeries").onclick = getSeries;
  document.getElementById("btnGetData").onclick = getData;
});
// -----------------------------

//  GET CATEGORIES
// -----------------------------
async function getSubcategories(categoryId) {
  // document.getElementById("output") gets the output element from the HTML page (taskpane.html)
  // by its ID and assigns it to the variable named "output".
  document.getElementById("output").innerHTML = "Fetching categories...";
  // .innerHTML = "Fetching categories..." sets the content of the output element to the string
  // This provides visual feedback to the user that the app is working on fetching the categories from the Fred API.

  try {
    // this line uses the Fetch API to make an HTTP GET request to the backend server at
    // "http://localhost:5000/categories/{categoryId}".
    // The backend server is expected to handle this request by calling the FRED API to get the list of categories and return
    // it as a JSON response.
    //
    // the backend server comes from the FastAPI found in main.py. The endpoint "/categories/{category_id}" is defined there
    // to handle this request and return the appropriate data.
    const response = await fetch(
      `http://localhost:5000/categories/${categoryId}`
    );
    // "await" is used to wait for the fetch request to complete and for the response to be received before moving on to the next line of code.

    // this line takes the response from the fetch request "response".
    // .json() is a method that parses the response's JSON content and converts it into a JavaScript object or array that can be used by the code.
    const data = await response.json();

    // After successfully fetching and parsing the categories data, the function "renderCategories" passes the parsed data to it.
    // The "renderCategories" now also receives the parent category ID so the UI knows where we are in the category tree.
    renderCategories(data, categoryId);

    // if theres an error during the fetch request or while parsing the response...
    // execute the code there.
  } catch (error) {
    // the follow line visually informs the user that there was an error fetching the categories.
    // it displays the message "Error fetching categories." in the output of the task pane.
    document.getElementById("output").innerHTML = "Error fetching categories.";

    // this line logs the error details to the console for debugging purposes.
    console.error(error);
  }
}

// this function takes the "categories" data (which should be an array of category objects) and generates HTML to display
// the categories in the task pane.
function renderCategories(categories, parentCategoryId) {
  // this line gets the output element from the file (taskpane.html) by its ID (also named output) and assigns it to the variable named "output"
  const output = document.getElementById("output");

  // "let" is used to declare a variable named "html" that will hold the HTML content to be displayed in the output element.
  // it ends with <ul> which is the opening tag for an unordered list in HTML.
  // The categories will be added as list items (<li>) inside this unordered list.
  let html = `<h3>Select a Category (Parent ID: ${parentCategoryId})</h3><ul>`;

  // cat = category object from the categories array
  // The forEach loop iterates through each category in the categories array (the parameter) then executes the provided function for each category.
  categories.forEach(cat => {
    // this ` is a template literal that allows for embedding expressions inside the string.
    // a literal is a string that can span multiple lines and include placeholders for variables or expressions, which are denoted by ${}.
    // It generates an HTML list item (<li>) for each category.
    html += `
      <li class="category-item" data-id="${cat.category_id}">
        ${cat.category} (ID: ${cat.category_id})
      </li>`;

    // html now looks like <h3>Select a Category</h3><ul> <li class="category-item" data-id="123">Category Name (ID: 123)</li> </ul>
    // notice how the category name and ID are added into the HTML using ${cat.category} and ${cat.category_id}.
  });

  // closes the unordered list by adding the closing </ul> tag to the html string.
  // This completes the HTML structure for displaying the list of categories in the task pane.
  html += "</ul>";

  // NEW: Add a button that lets the user view the series for the CURRENT category.
  // This allows the user to browse subcategories OR jump directly to the series.
  html += `
    <button id="btnViewSeries">View Series in Category ${parentCategoryId}</button>
  `;

  // this line sets the innerHTML of the output element to the generated HTML string.
  // .innerHTML is a property that allows you to set or get the HTML content of an element.
  // The list of categories will be displayed in the task pane for the user to see and interact with.
  output.innerHTML = html;

  // Add click handlers to each category item...
  // this function method querySelectorAll(".category-item") selects all elements in the document that have the class
  // "category-item" (which are the list items we just created for each category in the loop above).
  //
  // .forEach(item => { ... }) iterates through each of these selected elements (each category item) and executes
  // the function for each one.
  document.querySelectorAll(".category-item").forEach(item => {
    item.onclick = () => {
      // when clicked on a category item, this line retrieves the value of the "data-id" attribute from the clicked
      // item and assigns it to the global variable "selectedCategoryId".
      //
      // "item" is a function parameter that represents the current category item being processed in the forEach loop.
      // .getAttribute("data-id") is a method that retrieves the value of the "data-id" attribute from the HTML element.
      // the "data-id" attribute was set in the HTML generation step above to hold the category ID for each category item.
      selectedCategoryId = item.getAttribute("data-id");

      // NEW: Instead of just displaying the ID, we now automatically fetch the SUBCATEGORIES of the clicked category.
      // This enables full category-tree navigation.
      getSubcategories(selectedCategoryId);
    };
  });

  // NEW: Add click handler for the "View Series" button.
  // This allows the user to fetch the series for the CURRENT category level.
  document.getElementById("btnViewSeries").onclick = () => {
    // update the global selectedCategoryId so getSeries() knows which category to fetch series for
    selectedCategoryId = parentCategoryId;

    // call the getSeries function to fetch and display the series for this category
    getSeries();
  };
}
// -----------------------------

// GET SERIES FOR SELECTED CATEGORY
// -----------------------------

// this function is called when the user clicks the "Get Series in Category" button in the task pane.
// It fetches the series data for the selected category from the backend and displays it in the task pane.
async function getSeries() {
  // This function checks whether the user has selected a category yet.
  // selectedCategoryId is a global variable that gets set when the user clicks a category in renderCategories() located above.
  // If it is null or undefined, the user hasn't selected anything.
  // If the user hasn't selected a category, we want to show a message and not attempt to fetch series data from the backend.
  //
  // this if statement checks if selectedCategoryId would error (null, undefined, empty string, etc.). If it does, it means no
  // category has been selected.
  if (!selectedCategoryId) {

    // If no category is selected, the output tells the user what they need to do.
    document.getElementById("output").innerHTML = "Please select a category first.";

    // return stops the function immediately so no backend call is made.
    return;
  }

  // If a category IS selected then a loading message is displayed in the task pane.
  document.getElementById("output").innerHTML = "Fetching series for selected category...";

  try {
    // This makes an HTTP GET request to your backend.
    // It calls the /series/{category_id} endpoint using the selectedCategoryId.
    //
    // remember "await" is simply waiting for the fetch request to complete and for the response to
    // be received before moving on to the next line of code.
    const response = await fetch(
      `http://localhost:5000/series/${selectedCategoryId}`
    );

    // data is simply the parsed JSON response from the backend which should be an array of series objects for the user's selected category.
    const data = await response.json();

    // Pass the data to the renderSeries() function,
    // which will display the list of series in the task pane.
    renderSeries(data);

  // if there is an error during the fetch. display this error
  } catch (error) {
    // the html's output string becomes the string below
    document.getElementById("output").innerHTML = "Error fetching series.";

    // Log the actual error to the console for debugging
    console.error(error);
  }
}


// this function takes the list of series data (which should be an array of series objects) and generates HTML to display
// the series in the task pane. It also adds click handlers to each series item so that when a user clicks on a series
// it updates the selectedSeriesId variable and displays the selected series ID in the task pane.
function renderSeries(seriesList) {
  // this line gets the output element from the file (taskpane.html) by its ID (also named output) and assigns it to the variable named "output"
  const output = document.getElementById("output");

  
  // "let" is used to declare a variable named "html" that will hold the HTML content to be displayed in the output element.
  // it ends with <ul> which is the opening tag for an unordered list in HTML.
  // The categories will be added as list items (<li>) inside this unordered list.
  let html = "<h3>Select a Series</h3><ul>";

  seriesList.forEach(series => {
    html += `
      <li class="series-item" data-id="${series.series_id}">
        ${series.title} (ID: ${series.series_id})
      </li>`;
  });

  html += "</ul>";
  output.innerHTML = html;

  // Add click handlers to each series item
  document.querySelectorAll(".series-item").forEach(item => {
    item.onclick = () => {
      selectedSeriesId = item.getAttribute("data-id");
      output.innerHTML = `<p>Selected Series ID: ${selectedSeriesId}</p>`;
    };
  });
}

// -----------------------------

// GET DATA FOR SELECTED SERIES + INSERT INTO EXCEL
// -----------------------------
// this function is called when the user clicks the "Get Data for Series" button in the task pane.
// It fetches the data for the selected series from the backend and then calls another function to insert that data into Excel.
async function getData() {
  if (!selectedSeriesId) {
    document.getElementById("output").innerHTML =
      "Please select a series first.";
    return;
  }

  document.getElementById("output").innerHTML =
    "Fetching data for selected series...";

  try {
    const response = await fetch(
      `http://localhost:5000/data/${selectedSeriesId}`
    );
    const data = await response.json();

    // insertDataIntoExcel() is defined below and takes the data fetched from the backend and inserts it into the active Excel worksheet.
    // it goes into the first two columns and as many rows as needed depending on how much data there is.
    await insertDataIntoExcel(data);

    // display a success message in the task pane after the data has been loaded and inserted into Excel.
    document.getElementById("output").innerHTML =
      "Data inserted into Excel successfully.";
  } catch (error) {
    // display an error message in the task pane if there was an issue fetching the data for the selected series or inserting it into Excel.
    document.getElementById("output").innerHTML = "Error fetching data.";
    console.error(error);
  }
}
// -----------------------------

// -----------------------------
// this function takes the data for the selected series (which should be an array of objects with date and value properties) and inserts it into the active Excel worksheet.
async function insertDataIntoExcel(data) {
  // Excel.run is a method from the Excel's JavaScript API that allows interation with Excel
  // the .run method takes a function as an argument. The parameter "context" is an object which provides access to the Excel.
  // --- context.workbook gives access to the current workbook
  // --- context.workbook.worksheets gives access to the worksheets in the workbook
  // --- context.workbook.worksheets.getActiveWorksheet() gets the currently active worksheet that the user has open in Excel.
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Convert JSON array into 2D array for Excel
    const rows = data.map(item => [item.date, item.value]);

    // this line gets a range of cells in the active worksheet starting from A1 and extending down to as many rows as there are in the data, and 2 columns (A and B).
    // back ticks are used to create a literal string that allows embedding the number of rows dynamically into the range reference.
    // that way, if there are 10 data points, the range will be A1:B10; if there are 100 data points, the range will be A1:B100, etc.
    const range = sheet.getRange(`A1:B${rows.length}`);
    // this imports the 2D array of data (rows) into the specified range in Excel, inserting the date and value data into columns A and B of the active worksheet.
    range.values = rows;

    // allows and prevents Excel from trying to update the worksheet until all the changes have been made. This can improve performance and prevent flickering as the data is being inserted.
    await context.sync();
  });
}
// -----------------------------
// This function is called when the user clicks the "Search" button in the task pane. It takes the search query from an input field, sends it to the backend, and displays the search results in the task pane.
async function searchInput() {
  const query = document.getElementById("searchInput").value;
  document.getElementById("output").innerHTML = `Searching for "${query}"...`;

  try {
    const response = await fetch(
      `http://localhost:5000/search/${encodeURIComponent(query)}`
    );
    const data = await response.json();
    renderSeries(data);
  } catch (error) {
    document.getElementById("output").innerHTML = "Error performing search.";
    console.error(error);
  }
} 