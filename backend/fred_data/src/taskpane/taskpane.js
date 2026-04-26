/* --------------------------------------------------------------------------
   taskpane.js - Fixed version for category navigation freezing
-------------------------------------------------------------------------- */

const BACKEND_BASE_URL = "http://127.0.0.1:8080";

/* --------------------------------------------------------------------------
   GLOBAL STATE & CACHES
-------------------------------------------------------------------------- */

function loadCacheFromLocalStorage(key) {
    const cachedString = localStorage.getItem(key);
    if (cachedString) {
        try {
            const cachedArray = JSON.parse(cachedString);
            return new Map(cachedArray);
        } catch (e) {
            console.error(`Error parsing cache for ${key} from localStorage:`, e);
            return new Map();
        }
    }
    return new Map();
}

function saveCacheToLocalStorage(key, cacheMap) {
    const arrayToStore = Array.from(cacheMap.entries());
    localStorage.setItem(key, JSON.stringify(arrayToStore));
}

let categoryCache = loadCacheFromLocalStorage('categoryCache');
let seriesCache = loadCacheFromLocalStorage('seriesCache');
let dataCache = loadCacheFromLocalStorage('dataCache');
let navStack = [{ type: 'main' }];

let currentSeriesId = null;
let currentData = null;
let currentSeriesName = null;

let outputEl = null;
let infoEl = null;
let btnLoadDataCurrentSheet = null;
let btnLoadDataNewSheet = null;
let btnQuickRecall = null;
let btnViewBookmarks = null;
let lastLoadedSeries = null;
let sessionBookmarks = new Map();
let lastLoadedDataRange = null; // Stores range info for Pivot Table generation

/* --------------------------------------------------------------------------
   INITIALIZATION
-------------------------------------------------------------------------- */

Office.onReady(() => {
    outputEl = document.getElementById("output");
    infoEl = document.getElementById("info");
    btnLoadDataCurrentSheet = document.getElementById("btnLoadDataCurrentSheet");
    btnLoadDataNewSheet = document.getElementById("btnLoadDataNewSheet");
    btnQuickRecall = document.getElementById("btnQuickRecall");
    btnViewBookmarks = document.getElementById("btnViewBookmarks");

    document.getElementById("btnGetCategories").onclick = loadRootCategories;
    document.getElementById("btnStatisticalAnalysis").onclick = showStatisticalAnalysisMenu;
    document.getElementById("btnSearch").onclick = textSearch;
    document.getElementById("btnBack").onclick = goBack;
    btnLoadDataCurrentSheet.onclick = loadDataIntoCurrentSheet;
    btnLoadDataNewSheet.onclick = loadDataIntoNewSheet;
    btnQuickRecall.onclick = quickRecallLastSeries;
    btnViewBookmarks.onclick = viewSessionBookmarks;

    document.getElementById("btnResetCache").onclick = resetCache;
    document.getElementById("btnStatisticalAnalysis").disabled = true;

    // Use event delegation for dynamic buttons
    outputEl.addEventListener("click", onOutputClick);
});

/* --------------------------------------------------------------------------
   HELPERS
-------------------------------------------------------------------------- */

async function fetchJSON(url) {
    const res = await fetch(url);
    if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.detail || `HTTP Error: ${res.status} ${res.statusText}`);
    }
    return await res.json();
}

/**
 * Normalize API responses:
 * Accepts:
 *  - an array (direct)
 *  - an object with keys: categories, subcategories, results, data, series
 * Returns an array (possibly empty)
 */
function normalizeArrayResponse(resp, candidateKeys = []) {
    if (!resp) return [];
    if (Array.isArray(resp)) return resp;
    // If resp is an object, try common wrapper keys
    for (const k of candidateKeys) {
        if (Array.isArray(resp[k])) return resp[k];
    }
    // If resp has a top-level property that is an array, return it (fallback)
    for (const key of Object.keys(resp)) {
        if (Array.isArray(resp[key])) return resp[key];
    }
    return [];
}

async function getSubcategories(categoryId) {
    if (categoryCache.has(String(categoryId))) {
        return categoryCache.get(String(categoryId));
    }
    const raw = await fetchJSON(`${BACKEND_BASE_URL}/categories/${categoryId}`);
    const data = normalizeArrayResponse(raw, ['categories', 'subcategories', 'results']);
    categoryCache.set(String(categoryId), data);
    saveCacheToLocalStorage('categoryCache', categoryCache);
    return data;
}

async function getSeries(categoryId) {
    if (seriesCache.has(String(categoryId))) {
        return seriesCache.get(String(categoryId));
    }
    const raw = await fetchJSON(`${BACKEND_BASE_URL}/series/${categoryId}`);
    const data = normalizeArrayResponse(raw, ['series', 'results', 'data']);
    seriesCache.set(String(categoryId), data);
    saveCacheToLocalStorage('seriesCache', seriesCache);
    return data;
}

async function getData(seriesId) {
    if (dataCache.has(String(seriesId))) {
        return dataCache.get(String(seriesId));
    }
    const raw = await fetchJSON(`${BACKEND_BASE_URL}/data/${seriesId}`);
    const data = normalizeArrayResponse(raw, ['data', 'observations', 'results']);
    dataCache.set(String(seriesId), data);
    saveCacheToLocalStorage('dataCache', dataCache);
    return data;
}

function updateBookmarkButtons() {
    btnQuickRecall.disabled = !lastLoadedSeries;
    btnViewBookmarks.disabled = sessionBookmarks.size === 0;
}

function updateBackButton() {
    document.getElementById("btnBack").disabled = navStack.length <= 1;
}

function updateStatisticalAnalysisButton() {
    const isDisabled = !currentData || currentData.length === 0;
    document.getElementById("btnStatisticalAnalysis").disabled = isDisabled;
    btnLoadDataCurrentSheet.disabled = isDisabled;
    btnLoadDataNewSheet.disabled = isDisabled;
    updateBookmarkButtons();
}

function resetCache() {
    localStorage.removeItem('categoryCache');
    localStorage.removeItem('seriesCache');
    localStorage.removeItem('dataCache');

    categoryCache = new Map();
    seriesCache = new Map();
    dataCache = new Map();

    infoEl.innerHTML = "<p>Cache has been cleared. Please start a new search or browse categories.</p>";
    outputEl.innerHTML = "";
    navStack = [{ type: 'main' }];
    currentSeriesId = null;
    currentData = null;
    lastLoadedSeries = null;
    sessionBookmarks = new Map();
    btnLoadDataCurrentSheet.disabled = true;
    btnLoadDataNewSheet.disabled = true;
    updateBackButton();
    updateStatisticalAnalysisButton();
}

/* --------------------------------------------------------------------------
   CATEGORY / SERIES BROWSING
-------------------------------------------------------------------------- */

async function loadRootCategories() {
    infoEl.innerHTML = "Loading categories...";
    outputEl.innerHTML = "";
    try {
        const categories = await getSubcategories(0);
        // Reset nav stack to root + categoryList(0)
        navStack = [{ type: 'main' }, { type: 'categoryList', id: 0 }];
        if (categories && categories.length > 0) {
            renderCategories(categories);
            infoEl.innerHTML = "";
        } else {
            outputEl.innerHTML = "<p>No categories found. Please try again later.</p>";
            infoEl.innerHTML = "";
        }
    } catch (err) {
        console.error("loadRootCategories error:", err);
        infoEl.innerHTML = `<p>Error loading root categories: ${err.message || err}</p>`;
    } finally {
        updateBackButton();
        updateStatisticalAnalysisButton();
    }
}

async function onOutputClick(event) {
    const target = event.target.closest("[data-type]");
    if (!target) return;

    // Prevent default to avoid accidental form submits
    event.preventDefault();

    const type = target.dataset.type;

    try {
        if (type === "category") {
            const categoryId = parseInt(target.dataset.id, 10);
            if (Number.isNaN(categoryId)) {
                console.warn("Invalid category id:", target.dataset.id);
                return;
            }
            await handleCategoryClick(categoryId);
        } else if (type === "series") {
            const seriesId = target.dataset.id;
            if (!seriesId) {
                console.warn("Invalid series id:", target.dataset.id);
                return;
            }
            await handleSeriesClick(seriesId);
        } else if (type === "bookmarkedSeries") {
            const seriesId = target.dataset.id;
            await handleBookmarkedSeriesClick(seriesId);
        }
    } catch (err) {
        console.error("onOutputClick handler error:", err);
        infoEl.innerHTML = `<p>Error handling click: ${err.message || err}</p>`;
    }
}

async function handleCategoryClick(categoryId) {
    infoEl.innerHTML = "Loading...";
    outputEl.innerHTML = "";

    // Push state only if different from last
    const lastState = navStack[navStack.length - 1];
    if (!lastState || lastState.type !== 'categoryList' || lastState.id !== categoryId) {
        navStack.push({ type: 'categoryList', id: categoryId });
    }

    try {
        const subcats = await getSubcategories(categoryId);
        if (subcats && subcats.length > 0) {
            renderCategories(subcats);
            infoEl.innerHTML = "";
        } else {
            // No subcategories, try series
            const seriesList = await getSeries(categoryId);
            if (seriesList && seriesList.length > 0) {
                renderSeries(seriesList);
                infoEl.innerHTML = "";
            } else {
                outputEl.innerHTML = "<p>No subcategories or series were found for this category.</p>";
                infoEl.innerHTML = "";
            }
        }
    } catch (error) {
        console.error("Category click error:", error);
        infoEl.innerHTML = `<p>Could not load data for this category. It may be too large or temporarily unavailable.</p>`;
    } finally {
        updateBackButton();
        updateStatisticalAnalysisButton();
    }
}

async function handleSeriesClick(seriesId) {
    infoEl.innerHTML = "Loading data for selected series...";
    outputEl.innerHTML = "";

    const lastState = navStack[navStack.length - 1];
    if (lastState.type !== 'seriesInfo' || lastState.id !== seriesId) {
        navStack.push({ type: 'seriesInfo', id: seriesId });
    }

    try {
        const data = await getData(seriesId);
        if (!data || data.length === 0) {
            infoEl.innerHTML = `<p>No data found for series: ${seriesId}.</p>`;
            btnLoadDataCurrentSheet.disabled = true;
            btnLoadDataNewSheet.disabled = true;
            return;
        }
        currentSeriesId = seriesId;
        currentData = data;

        const infoList = await fetchJSON(`${BACKEND_BASE_URL}/info/${seriesId}`);
        const info = Array.isArray(infoList) && infoList.length > 0 ? infoList[0] : (infoList || {});
        currentSeriesName = info.title || seriesId;
        lastLoadedSeries = { seriesId: currentSeriesId, data: currentData, info: info };
        sessionBookmarks.set(currentSeriesId, { seriesId: currentSeriesId, title: info.title || currentSeriesId });

        infoEl.innerHTML = `<p><strong>Series selected:</strong> ${seriesId}</p><p>Data is ready to load into Excel. Choose a load option.</p>`;
        btnLoadDataCurrentSheet.disabled = false;
        btnLoadDataNewSheet.disabled = false;
    } catch (error) {
        console.error("Series click error:", error);
        infoEl.innerHTML = `<p>Could not load data for series '${seriesId}'. It may be discontinued or temporarily unavailable.</p>`;
        btnLoadDataCurrentSheet.disabled = true;
        btnLoadDataNewSheet.disabled = true;
    } finally {
        updateBackButton();
        updateStatisticalAnalysisButton();
    }
}

async function goBack() {
    if (navStack.length <= 1) return;

    navStack.pop();
    const previousState = navStack[navStack.length - 1];

    infoEl.innerHTML = "";
    outputEl.innerHTML = "";
    currentSeriesId = null;
    currentData = null;
    btnLoadDataCurrentSheet.disabled = true;
    btnLoadDataNewSheet.disabled = true;
    document.getElementById("btnStatisticalAnalysis").disabled = true;

    try {
        switch (previousState.type) {
            case 'main':
                // nothing to render
                break;
            case 'categoryList':
                const subcats = await getSubcategories(previousState.id);
                if (subcats && subcats.length > 0) {
                    renderCategories(subcats);
                } else {
                    const seriesList = await getSeries(previousState.id);
                    if (seriesList && seriesList.length > 0) {
                        renderSeries(seriesList);
                    } else {
                        outputEl.innerHTML = "<p>No results were found for this category.</p>";
                    }
                }
                break;
            // other states can be handled if needed
        }
    } catch (err) {
        console.error("goBack error:", err);
        infoEl.innerHTML = `<p>Error navigating back: ${err.message || err}</p>`;
    } finally {
        updateBackButton();
        updateStatisticalAnalysisButton();
    }
}

/* --------------------------------------------------------------------------
   TEXT SEARCH LOGIC
-------------------------------------------------------------------------- */

async function textSearch() {
    const input = document.getElementById("searchInput").value.trim();
    if (!input) return;

    infoEl.innerHTML = "Searching and loading data into memory...";
    outputEl.innerHTML = "";

    navStack = [{ type: 'main' }, { type: 'searchResult', query: input }];

    try {
        const encodedInput = encodeURIComponent(input);
        const result = await fetchJSON(`${BACKEND_BASE_URL}/logic/search?q=${encodedInput}`);

        if (result.type === "series") {
            const infoList = normalizeArrayResponse(result.info || result, ['info', 'results']);
            const data = normalizeArrayResponse(result.data || result, ['data', 'observations']);
            if (infoList.length > 0) {
                currentSeriesId = infoList[0].series_id;
                currentData = data;
                currentSeriesName = infoList[0].title || currentSeriesId;

                lastLoadedSeries = { seriesId: currentSeriesId, data: currentData, info: infoList[0] };
                sessionBookmarks.set(currentSeriesId, { seriesId: currentSeriesId, title: infoList[0].title });

                renderSeriesInfo(infoList);

                infoEl.innerHTML = `
                    <p><strong>Success!</strong> Series <strong>${currentSeriesId}</strong> has been loaded into memory.</p>
                    <p>You can now click "Load Data into Excel".</p>
                `;
                btnLoadDataCurrentSheet.disabled = false;
                btnLoadDataNewSheet.disabled = false;
                updateStatisticalAnalysisButton();
                updateBookmarkButtons();
            } else {
                infoEl.innerHTML = `<p>No data found for the given input.</p>`;
                btnLoadDataCurrentSheet.disabled = true;
                btnLoadDataNewSheet.disabled = true;
                updateStatisticalAnalysisButton();
            }
        } else {
            infoEl.innerHTML = `<p>No series result returned from search.</p>`;
            btnLoadDataCurrentSheet.disabled = true;
            btnLoadDataNewSheet.disabled = true;
        }
    } catch (error) {
        console.error("Search error:", error);
        infoEl.innerHTML = `<p>Error during search. Please ensure you entered a valid FRED URL or Series ID.</p>`;
        btnLoadDataCurrentSheet.disabled = true;
        btnLoadDataNewSheet.disabled = true;
    }
}

/* --------------------------------------------------------------------------
   RENDERING FUNCTIONS
-------------------------------------------------------------------------- */

function renderCategories(categories) {
    outputEl.innerHTML = "<h3>Select a Category</h3>";
    const frag = document.createDocumentFragment();

    categories.forEach(cat => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "buttonStyle";
        button.dataset.type = "category";
        button.dataset.id = String(cat.category_id ?? cat.id ?? cat.categoryId ?? cat.category_id);
        button.textContent = cat.category ?? cat.name ?? cat.title ?? `Category ${button.dataset.id}`;
        frag.appendChild(button);
    });

    outputEl.appendChild(frag);
}

function renderSeries(seriesList) {
    outputEl.innerHTML = "<h3>Select a Series</h3>";
    const frag = document.createDocumentFragment();

    seriesList.forEach(s => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "buttonStyle";
        button.dataset.type = "series";
        button.dataset.id = s.series_id ?? s.id ?? s.seriesId;
        button.textContent = `${s.title ?? s.name ?? 'Series'} (${button.dataset.id})`;
        frag.appendChild(button);
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
   STATISTICAL ANALYSIS (unchanged except for safety checks)
-------------------------------------------------------------------------- */

async function showStatisticalAnalysisMenu() {
    if (!currentData || currentData.length === 0) {
        infoEl.innerHTML = "<p>Please load data for a series first to enable statistical analysis.</p>";
        return;
    }

    infoEl.innerHTML = `<p>Statistical Analysis for Series: <strong>${currentSeriesId}</strong></p>`;
    outputEl.innerHTML = "<h3>Select a Statistical Tool</h3>";

    navStack.push({ type: 'statisticalMenu' });

    const frag = document.createDocumentFragment();

    const descriptiveStatsButton = document.createElement("button");
    descriptiveStatsButton.type = "button";
    descriptiveStatsButton.className = "buttonStyle";
    descriptiveStatsButton.textContent = "Descriptive Statistics";
    descriptiveStatsButton.onclick = calculateDescriptiveStatistics;
    frag.appendChild(descriptiveStatsButton);

    outputEl.appendChild(frag);
    updateBackButton();
}

async function calculateDescriptiveStatistics() {
    try {
        await Excel.run(async context => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load("values, rowIndex");
            await context.sync();

            let values = [];

            if (!Array.isArray(selectedRange.values) || selectedRange.values.length === 0) {
                infoEl.innerHTML = "<p>Please select a range of cells containing numeric data for statistical analysis.</p>";
                return;
            }

            selectedRange.values.forEach(row => {
                if (Array.isArray(row)) {
                    row.forEach(cellValue => {
                        if (typeof cellValue === 'number' && !isNaN(cellValue)) {
                            values.push(cellValue);
                        }
                    });
                }
            });

            if (values.length === 0) {
                infoEl.innerHTML = "<p>Please select a range of cells containing numeric data for statistical analysis.</p>";
                return;
            }

            const sum = values.reduce((a, b) => a + b, 0);
            const mean = sum / values.length;
            const sortedValues = [...values].sort((a, b) => a - b);
            const mid = Math.floor(sortedValues.length / 2);
            const median = sortedValues.length % 2 === 0 ? (sortedValues[mid - 1] + sortedValues[mid]) / 2 : sortedValues[mid];
            const min = Math.min(...values);
            const max = Math.max(...values);
            const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
            const stdDev = Math.sqrt(variance);
            const count = values.length;

            const results = [
                ["Count", count],
                ["Mean", mean.toFixed(4)],
                ["Median", median.toFixed(4)],
                ["Standard Deviation", stdDev.toFixed(4)],
                ["Min", min.toFixed(4)],
                ["Max", max.toFixed(4)]
            ];

            const outputExcelRow = selectedRange.rowIndex + 1;
            const outputRange = sheet.getRange(`C${outputExcelRow}:D${outputExcelRow + results.length - 1}`);
            outputRange.values = results;
            outputRange.format.autofitColumns();
            await context.sync();

            infoEl.innerHTML = `<p>Descriptive Statistics calculated and placed in Excel starting from C${outputExcelRow}.</p>`;
        });
    } catch (error) {
        console.error("Descriptive statistics error:", error);
        infoEl.innerHTML = `<p>Error performing descriptive statistics: ${error.message || error}</p>`;
    }
}

/* --------------------------------------------------------------------------
   EXCEL INSERTION FUNCTIONS (unchanged except for small safety checks)
-------------------------------------------------------------------------- */

async function loadDataIntoCurrentSheet() {
    if (!currentData || currentData.length === 0) {
        infoEl.innerHTML = "<p>No data loaded yet.</p>";
        return;
    }

    btnLoadDataCurrentSheet.disabled = true;
    btnLoadDataNewSheet.disabled = true;

    try {
        // Use the already-loaded metadata from memory instead of re-fetching.
        if (!lastLoadedSeries || !lastLoadedSeries.info) {
            throw new Error("Could not find cached metadata. Please search for the series again.");
        }
        const metadata = lastLoadedSeries.info;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("columnCount, isNullObject");
            sheet.load("name"); // Load the name property of the sheet object
            await context.sync();

            let startColumnIndex;
            if (usedRange.isNullObject) {
                startColumnIndex = 1;
            } else {
                startColumnIndex = usedRange.columnCount + 2;
            }

            const metadataLabels = [
                "Series ID:", "Title:", "Frequency:", "Units:", "Seasonal Adjustment:", "Observation Start:"
            ];
            const metadataValues = [
                metadata.series_id, metadata.title, metadata.frequency, metadata.units,
                metadata.seasonal_adjustment, metadata.observation_start
            ];
            const metadataArray = metadataLabels.map((label, index) => [label, metadataValues[index] || ""]);

            const dataValues = [["Date", currentSeriesId]];
            currentData.forEach(r => {
                dataValues.push([r.date, r.value]);
            });

            const metadataStartColZeroBased = startColumnIndex - 1;
            const metadataRange = sheet.getRangeByIndexes(0, metadataStartColZeroBased, metadataArray.length, 2);
            metadataRange.values = metadataArray;
            metadataRange.format.autofitColumns();
            metadataRange.getCell(0, 0).getResizedRange(metadataArray.length - 1, 0).format.font.bold = true;

            const dataStartColZeroBased = metadataStartColZeroBased + 3;
            const dataRange = sheet.getRangeByIndexes(0, dataStartColZeroBased, dataValues.length, 2); // 2 columns: Date, Value
            dataRange.values = dataValues;
            dataRange.format.autofitColumns();
            dataRange.getRow(0).format.font.bold = true;

            const safeTableName = "Table_" + currentSeriesId.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);
            const table = sheet.tables.add(dataRange, true);
            table.name = safeTableName;
            table.style = "TableStyleMedium2";

            // Define the target cell for the pivot table
            const pivotTargetCell = sheet.getCell(0, dataStartColZeroBased + 6); // Leave space for hidden grouping columns + 1 padding

            // Load the address properties before accessing them
            dataRange.load("address");
            pivotTargetCell.load("address");
            await context.sync(); // Synchronize to get the loaded addresses

            lastLoadedDataRange = {
                sheetName: sheet.name,
                tableName: safeTableName,
                dataRangeAddress: dataRange.address,
                dataStartColZeroBased: dataStartColZeroBased,
                rowCount: dataValues.length,
                pivotTargetAddress: pivotTargetCell.address,
                seriesId: currentSeriesId
            };
        });

        promptForPivotTable();
    } catch (error) {
        console.error("Error loading data into current sheet:", error);
        infoEl.innerHTML = `<p>Error loading data into current sheet: ${error.message || error}</p>`;
    } finally {
        btnLoadDataCurrentSheet.disabled = false;
        btnLoadDataNewSheet.disabled = false;
    }
    await createHistogram(currentSeriesName);
}

async function loadDataIntoNewSheet() {
    if (!currentData || currentData.length === 0) {
        infoEl.innerHTML = "<p>No data loaded yet.</p>";
        return;
    }

    btnLoadDataCurrentSheet.disabled = true;
    btnLoadDataNewSheet.disabled = true;

    try {
        // Use the already-loaded metadata from memory instead of re-fetching.
        if (!lastLoadedSeries || !lastLoadedSeries.info) {
            throw new Error("Could not find cached metadata. Please search for the series again.");
        }
        const metadata = lastLoadedSeries.info;

        await Excel.run(async (context) => {
            let sheetName = currentSeriesId.replace(/[^a-zA-Z0-9 ]/g, "").substring(0, 31).trim();
            if (sheetName.length === 0) sheetName = "FRED_Data";
            const newSheet = context.workbook.worksheets.add(sheetName);
            newSheet.activate(); // Activate the new sheet
            newSheet.load("name"); // Load the name property of the new sheet
            await context.sync(); // Sync to ensure newSheet.name is loaded

            const metadataLabels = [
                "Series ID:", "Title:", "Frequency:", "Units:", "Seasonal Adjustment:", "Observation Start:"
            ];
            const metadataValues = [
                metadata.series_id, metadata.title, metadata.frequency, metadata.units,
                metadata.seasonal_adjustment, metadata.observation_start
            ];
            const metadataArray = metadataLabels.map((label, index) => [label, metadataValues[index] || ""]);

            const dataValues = [["Date", currentSeriesId]];
            currentData.forEach(r => {
                dataValues.push([r.date, r.value]);
            });

            const metadataRange = newSheet.getRangeByIndexes(0, 0, metadataArray.length, 2);
            metadataRange.values = metadataArray;
            metadataRange.format.autofitColumns();
            metadataRange.getCell(0, 0).getResizedRange(metadataArray.length - 1, 0).format.font.bold = true;

            const dataRange = newSheet.getRangeByIndexes(0, 3, dataValues.length, 2); // 2 columns: Date, Value
            dataRange.values = dataValues;
            dataRange.format.autofitColumns();
            dataRange.getRow(0).format.font.bold = true;

            const safeTableName = "Table_" + currentSeriesId.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);
            const table = newSheet.tables.add(dataRange, true);
            table.name = safeTableName;
            table.style = "TableStyleMedium2";

            // Define the target cell for the pivot table
            const pivotTargetCell = newSheet.getCell(0, 9); // Col J (3 for metadata/pad + 2 for data + 3 for hidden columns + 1 for padding)

            // Load the address properties before accessing them
            dataRange.load("address");
            pivotTargetCell.load("address");
            await context.sync(); // Synchronize to get the loaded addresses

            lastLoadedDataRange = {
                sheetName: newSheet.name,
                tableName: safeTableName,
                dataRangeAddress: dataRange.address,
                dataStartColZeroBased: 3,
                rowCount: dataValues.length,
                pivotTargetAddress: pivotTargetCell.address,
                seriesId: currentSeriesId
            };
        });

        promptForPivotTable();
    } catch (error) {
        console.error("Error loading data into new sheet:", error);
        infoEl.innerHTML = `<p>Error loading data into new sheet: ${error.message || error}</p>`;
    } finally {
        btnLoadDataCurrentSheet.disabled = false;
        btnLoadDataNewSheet.disabled = false;
    }
    await createHistogram(currentSeriesName);
}

/* --------------------------------------------------------------------------
   PIVOT TABLE GENERATION
-------------------------------------------------------------------------- */

function promptForPivotTable() {
    navStack.push({ type: 'pivotPrompt' });
    infoEl.innerHTML = `<p>Data successfully loaded into Excel! Would you like to generate a Pivot Table?</p>`;
    
    // 1. Determine dataset frequency
    const freq = (lastLoadedSeries && lastLoadedSeries.info && lastLoadedSeries.info.frequency) 
        ? lastLoadedSeries.info.frequency.toLowerCase() 
        : "";
    
    // 2. Define logic for which options are available based on frequency
    const isAnnual = freq.includes("annual") && !freq.includes("semi") && !freq.includes("bi");
    const isQuarterly = freq.includes("semi") || freq.includes("quarter");
    
    const canUseYear = true; // Year is always an option
    const canUseQuarter = !isAnnual; 
    const canUseMonth = !isAnnual && !isQuarterly;

    // 3. Define default checked states (only the most granular option is checked by default)
    const checkYear = canUseYear && !canUseQuarter && !canUseMonth;
    const checkQuarter = canUseQuarter && !canUseMonth;
    const checkMonth = canUseMonth;

    // 4. Create a helper function to generate readable HTML for checkboxes
    const notApplicableHtml = `<span style="color: #d83b01; font-style: italic; font-size: 12px; margin-left: 5px;">(Not Applicable to Data Set)</span>`;

    function createCheckboxHtml(id, labelText, isEnabled, isChecked) {
        const disabledAttr = isEnabled ? "" : "disabled";
        const checkedAttr = isChecked ? "checked" : "";
        const cursorStyle = isEnabled ? "pointer" : "not-allowed";
        const extraText = isEnabled ? "" : notApplicableHtml;
        
        return `
            <div style="margin-bottom: 8px;">
                <label style="cursor: ${cursorStyle}; display: inline-flex; align-items: center;">
                    <input type="checkbox" id="${id}" value="${labelText}" ${disabledAttr} ${checkedAttr} style="margin-right: 5px;">
                    <span style="font-weight: bold;">${labelText}</span>
                </label>
                ${extraText}
            </div>
        `;
    }

    const yearHtml = createCheckboxHtml("ptYear", "Years", canUseYear, checkYear);
    const quarterHtml = createCheckboxHtml("ptQuarter", "Quarters", canUseQuarter, checkQuarter);
    const monthHtml = createCheckboxHtml("ptMonth", "Months", canUseMonth, checkMonth);

    // 5. Construct the final output HTML
    outputEl.innerHTML = `
        <div style="text-align: left; margin-top: 15px; padding: 15px; border: 1px solid #ccc; background-color: #f9f9f9; border-radius: 6px;">
            
            <p style="margin-top: 0; margin-bottom: 10px;"><strong>Group Dates By (Select at least one):</strong></p>
            ${yearHtml}
            ${quarterHtml}
            ${monthHtml}
            
            <hr style="border: 0; border-top: 1px solid #ddd; margin: 15px 0;">

            <p style="margin-top: 0; margin-bottom: 5px;"><strong>Summarize Value By:</strong></p>
            <select id="ptAggregation" style="width: 100%; padding: 6px; margin-bottom: 15px; border-radius: 4px; border: 1px solid #ccc;">
                <option value="Sum">Sum</option>
                <option value="Count">Count</option>
                <option value="Average">Average</option>
                <option value="Max">Max</option>
                <option value="Min">Min</option>
                <option value="Product">Product</option>
                <option value="CountNumbers">Count Numbers</option>
                <option value="StandardDeviation">StdDev</option>
                <option value="StandardDeviationP">StdDevP</option>
                <option value="Variance">Var</option>
                <option value="VarianceP">VarP</option>
            </select>

            <button id="btnGeneratePivot" class="buttonStyle" style="width: 100%; margin-bottom: 10px;">Generate Pivot Table</button>
            <button id="btnSkipPivot" class="buttonStyle" style="width: 100%; background-color: #666;">Skip</button>
        </div>
    `;

    document.getElementById("btnGeneratePivot").onclick = generatePivotTable;
    document.getElementById("btnSkipPivot").onclick = restoreSeriesView;
    updateBackButton();
}

async function generatePivotTable() {
    const aggFn = document.getElementById("ptAggregation").value;
    const groupFields = [];
    if (document.getElementById("ptYear") && document.getElementById("ptYear").checked) groupFields.push("Years");
    if (document.getElementById("ptQuarter") && document.getElementById("ptQuarter").checked) groupFields.push("Quarters");
    if (document.getElementById("ptMonth") && document.getElementById("ptMonth").checked) groupFields.push("Months");

    if (groupFields.length === 0) {
        alert("Please select at least one date grouping.");
        return;
    }

    infoEl.innerHTML = "<p>Generating Pivot Table...</p>";
    document.getElementById("btnGeneratePivot").disabled = true;
    document.getElementById("btnSkipPivot").disabled = true;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(lastLoadedDataRange.sheetName);
            const table = sheet.tables.getItem(lastLoadedDataRange.tableName);
            
            const colData = {};
            groupFields.forEach(field => {
                colData[field] = [];
            });

            // Use native month abbreviations so Excel natively sorts them chronologically
            const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

            currentData.forEach(r => {
                const parts = r.date.split("-");
                if(parts.length === 3) {
                    const year = parseInt(parts[0], 10);
                    const month = parseInt(parts[1], 10);
                    const quarter = "Qtr" + Math.ceil(month / 3);
                    const monthName = monthNames[month - 1]; 

                    groupFields.forEach(field => {
                        if (field === "Years") colData[field].push([year]);
                        else if (field === "Quarters") colData[field].push([quarter]);
                        else if (field === "Months") colData[field].push([monthName]);
                    });
                } else {
                    groupFields.forEach(field => colData[field].push([""]));
                }
            });

            groupFields.forEach(field => {
                let newCol = table.columns.add(null, null, field);
                newCol.getDataBodyRange().values = colData[field];
                newCol.getRange().columnHidden = true;
            });
            
            const targetCell = sheet.getRange(lastLoadedDataRange.pivotTargetAddress);
            await context.sync();

            const safeName = "Pivot_" + lastLoadedDataRange.seriesId.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);
            const pivotTable = sheet.pivotTables.add(safeName, lastLoadedDataRange.tableName, targetCell);

            groupFields.forEach(field => {
                pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
            });

            const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(lastLoadedDataRange.seriesId));
            dataHierarchy.summarizeBy = aggFn;
            pivotTable.layout.layoutType = "Compact";

            await context.sync();
        });

        restoreSeriesView();
        infoEl.innerHTML += `<p style="color: green;">Pivot Table successfully generated!</p>`;
    } catch (error) {
        console.error("Pivot Table generation error:", error);
        infoEl.innerHTML = `<p>Error generating Pivot Table: ${error.message || error}</p>`;
        document.getElementById("btnGeneratePivot").disabled = false;
        document.getElementById("btnSkipPivot").disabled = false;
    }
}

function restoreSeriesView() {
    if (!lastLoadedSeries) return;
    renderSeriesInfo([lastLoadedSeries.info]);
    infoEl.innerHTML = `<p><strong>Series selected:</strong> ${currentSeriesId}</p><p>Data is ready to load into Excel. Choose a load option.</p>`;
    
    if (navStack[navStack.length - 1].type === 'pivotPrompt') {
        navStack.pop();
    }
    updateBackButton();
    updateStatisticalAnalysisButton();
}

/* --------------------------------------------------------------------------
   BOOKMARK / QUICK RECALL STUBS (kept minimal)
-------------------------------------------------------------------------- */

function quickRecallLastSeries() {
    if (!lastLoadedSeries) {
        infoEl.innerHTML = "<p>No series to recall.</p>";
        return;
    }
    currentSeriesId = lastLoadedSeries.seriesId ?? lastLoadedSeries.seriesId;
    currentData = lastLoadedSeries.data;
    infoEl.innerHTML = `<p>Recalled series <strong>${currentSeriesId}</strong>.</p>`;
    btnLoadDataCurrentSheet.disabled = false;
    btnLoadDataNewSheet.disabled = false;
}

async function handleBookmarkedSeriesClick(seriesId) {
    infoEl.innerHTML = "Loading bookmarked series...";
    outputEl.innerHTML = "";

    // Fetch data and info (this will use cache if available)
    try {
        const data = await getData(seriesId);
        const infoList = await fetchJSON(`${BACKEND_BASE_URL}/info/${seriesId}`);

        if (!data || data.length === 0 || !infoList || infoList.length === 0) {
            infoEl.innerHTML = `<p>Could not load data for bookmarked series: ${seriesId}. It might be unavailable.</p>`;
            btnLoadDataCurrentSheet.disabled = true;
            btnLoadDataNewSheet.disabled = true;
            return;
        }

        currentSeriesId = seriesId;
        currentData = data;
        lastLoadedSeries = { seriesId, data, info: infoList[0] }; // Update last loaded for quick recall

        renderSeriesInfo(infoList);
        infoEl.innerHTML = `<p><strong>Bookmarked Series Selected:</strong> ${seriesId}</p><p>Data is ready to load into Excel. Choose a load option.</p>`;
        btnLoadDataCurrentSheet.disabled = false;
        btnLoadDataNewSheet.disabled = false;
    } catch (error) {
        infoEl.innerHTML = `<p>Error loading bookmarked series '${seriesId}': ${error.message || error}</p>`;
        btnLoadDataCurrentSheet.disabled = true;
        btnLoadDataNewSheet.disabled = true;
        console.error("Bookmarked series load error:", error);
    } finally {
        updateStatisticalAnalysisButton();
        updateBackButton();
    }
}

function viewSessionBookmarks() {
    if (sessionBookmarks.size === 0) {
        infoEl.innerHTML = "<p>No bookmarks in this session.</p>";
        return;
    }
    outputEl.innerHTML = "<h3>Session Bookmarks</h3>";
    const frag = document.createDocumentFragment();
    sessionBookmarks.forEach((val, key) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "buttonStyle";
        button.dataset.type = "bookmarkedSeries";
        button.dataset.id = key;
        button.textContent = `${val.title ?? key} (${key})`;
        frag.appendChild(button);
    });
    outputEl.appendChild(frag);
    infoEl.innerHTML = "";
}

async function createHistogram(currentSeriesName) {
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        // Deletes any existing charts to avoid clutter if user loads multiple series
        sheet.charts.load("items");
        await context.sync();
        sheet.charts.items.forEach(chart => {
            chart.delete();
        }); 
        
        if (!lastLoadedDataRange || !lastLoadedDataRange.tableName) return;
        
        let table = sheet.tables.getItem(lastLoadedDataRange.tableName);
        let valueColumn = table.columns.getItemAt(1).getDataBodyRange();
        
        // Create histogram
        let chart = sheet.charts.add(Excel.ChartType.histogram, valueColumn, Excel.ChartSeriesBy.columns);
        // Chart title using series name
        chart.title.text = `${currentSeriesName} Histogram`;
        // Cell position
        chart.setPosition("A15", "D30");
        // Y-axis title
        chart.axes.valueAxis.title.text = "Frequency";
        chart.axes.valueAxis.title.visible = true;
        chart.axes.categoryAxis.title.text = "Value Range";
        chart.axes.categoryAxis.title.visible = true;
        // Improve X-axis readability
        chart.axes.categoryAxis.format.font.size = 9;
        await context.sync();
    });
}
