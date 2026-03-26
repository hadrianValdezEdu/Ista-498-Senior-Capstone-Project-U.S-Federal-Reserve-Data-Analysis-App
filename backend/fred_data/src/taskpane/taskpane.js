/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Global State
// -----------------------------
let selectedCategoryId = null;
let selectedSeriesId = null;
// -----------------------------

// -----------------------------
Office.onReady(() => {
  document.getElementById("btnGetCategories").onclick = () => {getSubcategories(0)};
  document.getElementById("btnGetSeries").onclick = getSeries;
  document.getElementById("btnGetData").onclick = getData;
});
// -----------------------------

//  GET CATEGORIES
// -----------------------------
async function getSubcategories(categoryId) {
  document.getElementById("output").innerHTML = "Fetching categories...";

  try {
    const response = await fetch(
      `http://localhost:5000/categories/${categoryId}`
    );

    const data = await response.json();

    renderCategories(data, categoryId);
  } catch (error) {
    document.getElementById("output").innerHTML = "Error fetching categories.";
    console.error(error);
  }
}
// -----------------------------

function renderCategories(categories, parentCategoryId) {
  const output = document.getElementById("output");

  let html = `<h3>Select a Category (Parent ID: ${parentCategoryId})</h3><ul>`;

  categories.forEach(cat => {
    html += `
      <li class="category-item" data-id="${cat.category_id}">
        ${cat.category} (ID: ${cat.category_id})
      </li>`;
  });
  html += "</ul>";
  html += `
    <button id="btnViewSeries">View Series in Category ${parentCategoryId}</button>
  `;
  output.innerHTML = html;

  document.querySelectorAll(".category-item").forEach(item => {
    item.onclick = () => {
      selectedCategoryId = item.getAttribute("data-id");
      getSubcategories(selectedCategoryId);
    };
  });

  document.getElementById("btnViewSeries").onclick = () => {
    selectedCategoryId = parentCategoryId;
    getSeries();
  };
}
// -----------------------------

// GET SERIES FOR SELECTED CATEGORY
// -----------------------------

async function getSeries() {
  if (!selectedCategoryId) {
    document.getElementById("output").innerHTML = "Please select a category first.";
    return;
  }

  document.getElementById("output").innerHTML = "Fetching series for selected category...";

  try {
    const response = await fetch(
      `http://localhost:5000/series/${selectedCategoryId}`
    );
    const data = await response.json();
    renderSeries(data);
  } catch (error) {
    document.getElementById("output").innerHTML = "Error fetching series.";
    console.error(error);
  }
}

function renderSeries(seriesList) {
  const output = document.getElementById("output");

  let html = "<h3>Select a Series</h3><ul>";

  seriesList.forEach(series => {
    html += `
      <li class="series-item" data-id="${series.series_id}">
        ${series.title} (ID: ${series.series_id})
      </li>`;
  });

  html += "</ul>";
  output.innerHTML = html;

  document.querySelectorAll(".series-item").forEach(item => {
    item.onclick = () => {
      selectedSeriesId = item.getAttribute("data-id");
      output.innerHTML = `<p>Selected Series ID: ${selectedSeriesId}</p>`;
    };
  });
}

// -----------------------------

// -----------------------------
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

    await insertDataIntoExcel(data);

    document.getElementById("output").innerHTML =
      "Data inserted into Excel successfully.";
  } catch (error) {
    document.getElementById("output").innerHTML = "Error fetching data.";
    console.error(error);
  }
}
// -----------------------------

// -----------------------------
async function insertDataIntoExcel(data) {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const rows = data.map(item => [item.date, item.value]);

    const range = sheet.getRange(`A1:B${rows.length}`);
    range.values = rows;

    await context.sync();
  });
}
// -----------------------------
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