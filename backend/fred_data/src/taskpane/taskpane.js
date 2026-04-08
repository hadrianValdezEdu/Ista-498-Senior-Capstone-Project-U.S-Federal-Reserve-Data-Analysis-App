/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


/* --------------------------------------------------------------------------
   GLOBAL STATE & CACHES
-------------------------------------------------------------------------- */

let categoryCache = new Map();   // categoryId -> subcategories[]
let seriesCache = new Map();     // categoryId -> series[]
let dataCache = new Map();       // seriesId -> data[]
let navStack = [];            // Stack to keep track of navigation for "back" button feature

let currentSeriesId = null;
let currentData = null;

let outputEl = null;
let infoEl = null;
let btnLoadData = null;

/* --------------------------------------------------------------------------
   INITIALIZATION
-------------------------------------------------------------------------- */

Office.onReady(() => {
    outputEl = document.getElementById("output");
    infoEl = document.getElementById("info");
    btnLoadData = document.getElementById("btnLoadData");

    document.getElementById("btnGetCategories").onclick = loadRootCategories;
    document.getElementById("btnSearch").onclick = textSearch;
    document.getElementById("btnBack").onclick = goBack;
    btnLoadData.onclick = insertDataIntoExcel;

    outputEl.addEventListener("click", onOutputClick);
});

/* --------------------------------------------------------------------------
   HELPERS
-------------------------------------------------------------------------- */

async function fetchJSON(url) {
    const res = await fetch(url);
    return await res.json();
}

async function getSubcategories(categoryId) {
    if (categoryCache.has(categoryId)) {
        return categoryCache.get(categoryId);
    }
    const data = await fetchJSON(`http://127.0.0.1:5000/categories/${categoryId}`);
    categoryCache.set(categoryId, data);
    return data;
}

async function getSeries(categoryId) {
    if (seriesCache.has(categoryId)) {
        return seriesCache.get(categoryId);
    }
    const data = await fetchJSON(`http://127.0.0.1:5000/series/${categoryId}`);
    seriesCache.set(categoryId, data);
    return data;
}

async function getData(seriesId) {
    if (dataCache.has(seriesId)) {
        return dataCache.get(seriesId);
    }
    const data = await fetchJSON(`http://127.0.0.1:5000/data/${seriesId}`);
    dataCache.set(seriesId, data);
    return data;
}

function updateBackButton() {
    const btnBack = document.getElementById("btnBack");
    btnBack.disabled = navStack.length === 0;
}

/* --------------------------------------------------------------------------
   CATEGORY / SERIES BROWSING
-------------------------------------------------------------------------- */

async function loadRootCategories() {
    infoEl.innerHTML = "";
    const categories = await getSubcategories(0);
    renderCategories(categories);
}

async function onOutputClick(event) {
    const target = event.target.closest("[data-type]");
    if (!target) return;

    const type = target.dataset.type;

    if (type === "category") {
        const categoryId = parseInt(target.dataset.id, 10);
        await handleCategoryClick(categoryId);
    } else if (type === "series") {
        const seriesId = target.dataset.id;
        await handleSeriesClick(seriesId);
    }
}

async function handleCategoryClick(categoryId) {
    infoEl.innerHTML = "";

    navStack.push(categoryId);

    const subcats = await getSubcategories(categoryId);

    if (subcats && subcats.length > 0) {
        renderCategories(subcats);
    } else {
        const seriesList = await getSeries(categoryId);
        renderSeries(seriesList);
    }

    updateBackButton();
}

async function handleSeriesClick(seriesId) {
    infoEl.innerHTML = "Loading data for selected series...";
    const data = await getData(seriesId);

    currentSeriesId = seriesId;
    currentData = data;

    infoEl.innerHTML = `
        <p><strong>Series selected:</strong> ${seriesId}</p>
        <p>Data is ready to load into Excel.</p>
    `;

    btnLoadData.disabled = false;
}

async function goBack() {
    if (navStack.length <= 1) {
        navStack = [];
        const rootCats = await getSubcategories(0);
        renderCategories(rootCats);
        updateBackButton();
        return;
    }

    navStack.pop();

    const previousCategoryId = navStack[navStack.length - 1];

    const subcats = await getSubcategories(previousCategoryId);

    if (subcats.length > 0) {
        renderCategories(subcats);
    } else {
        const seriesList = await getSeries(previousCategoryId);
        renderSeries(seriesList);
    }

    updateBackButton();
}

/* --------------------------------------------------------------------------
   TEXT SEARCH LOGIC
-------------------------------------------------------------------------- */

async function textSearch() {
    const input = document.getElementById("searchInput").value.trim();
    if (!input) return;

    infoEl.innerHTML = "Searching series...";
    outputEl.innerHTML = "";

    const result = await fetchJSON(`http://127.0.0.1:5000/logic/search/${input}`);

    const infoList = result.info || [];
    const data = result.data || [];

    currentSeriesId = infoList.length > 0 ? infoList[0].series_id : input;
    currentData = data;

    renderSeriesInfo(infoList);

    infoEl.innerHTML = `
        <p><strong>Series selected via search:</strong> ${currentSeriesId}</p>
        <p>Data is ready to load into Excel.</p>
    `;

    btnLoadData.disabled = false;
}

/* --------------------------------------------------------------------------
   RENDERING FUNCTIONS
-------------------------------------------------------------------------- */

function renderCategories(categories) {
    outputEl.innerHTML = "<h3>Select a Category</h3>";

    const frag = document.createDocumentFragment();

    categories.forEach(cat => {
        const div = document.createElement("div");
        div.className = "category-item";
        div.dataset.type = "category";
        div.dataset.id = cat.category_id;
        div.textContent = cat.category;
        frag.appendChild(div);
    });

    outputEl.appendChild(frag);
}

function renderSeries(seriesList) {
    outputEl.innerHTML = "<h3>Select a Series</h3>";

    const frag = document.createDocumentFragment();

    seriesList.forEach(s => {
        const div = document.createElement("div");
        div.className = "series-item";
        div.dataset.type = "series";
        div.dataset.id = s.series_id;
        div.textContent = `${s.title} (${s.series_id})`;
        frag.appendChild(div);
    });

    outputEl.appendChild(frag);
}

function renderSeriesInfo(infoList) {
    outputEl.innerHTML = "<h3>Series Info</h3>";

    const frag = document.createDocumentFragment();

    infoList.forEach(info => {
        const div = document.createElement("div");
        div.className = "series-info";
        div.innerHTML = `
            <strong>${info.title}</strong><br>
            ID: ${info.series_id}<br>
            Frequency: ${info.frequency}<br>
            Units: ${info.units}<br>
            Seasonal Adjustment: ${info.seasonal_adjustment}
        `;
        frag.appendChild(div);
    });

    outputEl.appendChild(frag);
}

/* --------------------------------------------------------------------------
   EXCEL INSERTION
-------------------------------------------------------------------------- */

async function insertDataIntoExcel() {
    if (!currentData || currentData.length === 0) {
        infoEl.innerHTML = "<p>No data loaded yet.</p>";
        return;
    }

    try {
        await Excel.run(async context => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            const rowCount = currentData.length;
            const range = sheet.getRange(`A1:B${rowCount}`);

            const values = currentData.map(r => [r.date, r.value]);
            range.values = values;

            sheet.activate();
            await context.sync();
        });

        infoEl.innerHTML = `
            <p>Data for <strong>${currentSeriesId}</strong> loaded into the active sheet.</p>
        `;
    } catch (error) {
        infoEl.innerHTML = `<p>Error loading data into Excel: ${error}</p>`;
    }
}