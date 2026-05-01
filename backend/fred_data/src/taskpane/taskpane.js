/* --------------------------------------------------------------------------
   taskpane.js - Fixed version for category navigation freezing
-------------------------------------------------------------------------- */
const BACKEND_BASE_URL = "https://localhost:8080"; // Use localhost as the backend is mapped to host's 8080

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
    try {
        localStorage.setItem(key, JSON.stringify(arrayToStore));
    } catch (e) {
        console.warn(`Could not save ${key} to localStorage (quota may be exceeded):`, e);
    }
}

let categoryCache = loadCacheFromLocalStorage('categoryCache');
let seriesCache = loadCacheFromLocalStorage('seriesCache');
let dataCache = new Map(); // Keep data in memory only to prevent 5MB localStorage QuotaExceededError
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
    document.getElementById("btnSearch").onclick = textSearch;
    document.getElementById("btnBack").onclick = goBack;
    btnLoadDataCurrentSheet.onclick = loadDataIntoCurrentSheet;
    btnLoadDataNewSheet.onclick = loadDataIntoNewSheet;
    btnQuickRecall.onclick = quickRecallLastSeries;
    btnViewBookmarks.onclick = viewSessionBookmarks;

    document.getElementById("btnResetCache").onclick = resetCache;
    
    // Add listener for sort dropdown
    document.getElementById("seriesSort").onchange = () => {
        const lastState = navStack[navStack.length - 1];
        if (lastState && lastState.type === 'categoryList') {
            // Re-trigger the category click to refresh list with new sorting
            handleCategoryClick(lastState.id);
        }
    };

    // Use event delegation for dynamic buttons
    outputEl.addEventListener("click", onOutputClick);
    updateStatisticalAnalysisButton();
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

async function getSeries(categoryId, orderBy = "popularity") {
    const cacheKey = `${categoryId}_${orderBy}`;
    if (seriesCache.has(cacheKey)) {
        return seriesCache.get(cacheKey);
    }
    const raw = await fetchJSON(`${BACKEND_BASE_URL}/series/${categoryId}?order_by=${orderBy}`);
    const data = normalizeArrayResponse(raw, ['series', 'results', 'data']);
    seriesCache.set(cacheKey, data);
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
    const isDataReadyForLoad = currentData && currentData.length > 0;
    btnLoadDataCurrentSheet.disabled = !isDataReadyForLoad;
    btnLoadDataNewSheet.disabled = !isDataReadyForLoad;
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
            const orderBy = document.getElementById("seriesSort").value;

            const [subcats, seriesList] = await Promise.all([
                getSubcategories(categoryId).catch(err => { console.warn("Subcategories fetch error:", err); return []; }),
                getSeries(categoryId, orderBy).catch(err => { console.warn("Series fetch error:", err); return []; })
            ]);

            outputEl.innerHTML = ""; // Clear immediately before rendering to prevent double-render race condition

            let hasContent = false;
            if (subcats && subcats.length > 0) {
                renderCategories(subcats);
                hasContent = true;
            }
            if (seriesList && seriesList.length > 0) {
                renderSeries(seriesList);
                hasContent = true;
        }

            if (!hasContent) {
                outputEl.innerHTML = "<p>No subcategories or series were found for this category.</p>";
            }
            infoEl.innerHTML = "";
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

    const orderBy = document.getElementById("seriesSort").value;

    try {
        switch (previousState.type) {
            case 'main':
                // nothing to render
                break;
            case 'categoryList':
                const [subcats, seriesList] = await Promise.all([
                    getSubcategories(previousState.id).catch(err => { console.warn("Subcategories fetch error:", err); return []; }),
                    getSeries(previousState.id, orderBy).catch(err => { console.warn("Series fetch error:", err); return []; })
                ]);

                outputEl.innerHTML = ""; // Clear immediately before rendering to prevent double-render race condition

                let hasContent = false;
                if (subcats && subcats.length > 0) {
                    renderCategories(subcats);
                    hasContent = true;
                }
                if (seriesList && seriesList.length > 0) {
                    renderSeries(seriesList);
                    hasContent = true;
                }

                if (!hasContent) {
                    outputEl.innerHTML = "<p>No results were found for this category.</p>";
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

        outputEl.innerHTML = ""; // Clear immediately before rendering to prevent double-render race condition

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
    if (!categories || categories.length === 0) return;

    const heading = document.createElement("h3");
    heading.textContent = "Categories";
    outputEl.appendChild(heading);

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
    if (!seriesList || seriesList.length === 0) return;

    const heading = document.createElement("h3");
    heading.textContent = "Series";
    outputEl.appendChild(heading);

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

        let htmlContent = '';
        if (info.title) {
            htmlContent += `<strong>${info.title}</strong><br><br>`;
        }

        for (const key in info) {
            if (Object.prototype.hasOwnProperty.call(info, key) && key !== 'title' && info[key]) {
                const label = key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
                htmlContent += `<strong>${label}:</strong> ${info[key]}<br>`;
            }
        }
        div.innerHTML = htmlContent;
        frag.appendChild(div);
    });

    outputEl.appendChild(frag);
}

/* --------------------------------------------------------------------------
   STATISTICAL ANALYSIS
-------------------------------------------------------------------------- */

function promptForDecomposition() {
    navStack.push({ type: 'decompPrompt' });
    infoEl.innerHTML = `<p>Data successfully loaded into Excel! Would you like to perform a Time Series Decomposition?</p>`;
    
    const freq = (lastLoadedSeries && lastLoadedSeries.info && lastLoadedSeries.info.frequency) 
        ? lastLoadedSeries.info.frequency.toLowerCase() 
        : "";
    const isSA = (lastLoadedSeries && lastLoadedSeries.info && lastLoadedSeries.info.seasonal_adjustment)
        ? lastLoadedSeries.info.seasonal_adjustment.toLowerCase().includes("seasonally adjusted") &&
          !lastLoadedSeries.info.seasonal_adjustment.toLowerCase().includes("not seasonally adjusted")
        : false;

    let L = 1;
    if (freq.includes("month")) L = 12;
    else if (freq.includes("quarter")) L = 4;
    else if (freq.includes("week")) L = 52;
    else if (freq.includes("day") || freq.includes("daily")) L = 365;

    const canUseSeasonal = !isSA && L > 1;
    const seasonalExtraHtml = canUseSeasonal ? "" : `<span style="color: #d83b01; font-style: italic; font-size: 12px; margin-left: 5px;">(Not Applicable to Data Set)</span>`;

    outputEl.innerHTML = `
        <div style="text-align: left; margin-top: 15px; padding: 15px; border: 1px solid #ccc; background-color: #f9f9f9; border-radius: 6px;">
            <p style="margin-top: 0; margin-bottom: 10px;"><strong>Select Decomposition Components:</strong></p>
            <div style="margin-bottom: 8px;">
                <label style="cursor: pointer; display: inline-flex; align-items: center;">
                    <input type="checkbox" id="dcTrend" value="Trend" checked style="margin-right: 5px;">
                    <span style="font-weight: bold;">Trend</span>
                </label>
            </div>
            <div style="margin-bottom: 8px;">
                <label style="cursor: pointer; display: inline-flex; align-items: center;">
                    <input type="checkbox" id="dcCyclical" value="Cyclical" checked style="margin-right: 5px;">
                    <span style="font-weight: bold;">Cyclical</span>
                </label>
            </div>
            <div style="margin-bottom: 8px;">
                <label style="cursor: ${canUseSeasonal ? 'pointer' : 'not-allowed'}; display: inline-flex; align-items: center;">
                    <input type="checkbox" id="dcSeasonal" value="Seasonal" ${canUseSeasonal ? 'checked' : 'disabled'} style="margin-right: 5px;">
                    <span style="font-weight: bold;">Seasonal</span>
                </label>
                ${seasonalExtraHtml}
            </div>
            <div style="margin-bottom: 8px;">
                <label style="cursor: pointer; display: inline-flex; align-items: center;">
                    <input type="checkbox" id="dcResidual" value="Residual" checked style="margin-right: 5px;">
                    <span style="font-weight: bold;">Residual</span>
                </label>
            </div>
            
            <button id="btnGenerateDecomp" class="buttonStyle" style="width: 100%; margin-top: 10px; margin-bottom: 10px;">Generate Decomposition</button>
            <button id="btnSkipDecomp" class="buttonStyle" style="width: 100%; background-color: #666;">Skip</button>
        </div>
    `;

    document.getElementById("btnGenerateDecomp").onclick = generateDecompositionFromPrompt;
    document.getElementById("btnSkipDecomp").onclick = () => {
        if (navStack[navStack.length - 1].type === 'decompPrompt') navStack.pop();
        promptForPivotTable();
    };
    updateBackButton();
}

async function generateDecompositionFromPrompt() {
    const useTrend = document.getElementById("dcTrend").checked;
    const useCyclical = document.getElementById("dcCyclical").checked;
    const useSeasonal = document.getElementById("dcSeasonal") && !document.getElementById("dcSeasonal").disabled && document.getElementById("dcSeasonal").checked;
    const useResidual = document.getElementById("dcResidual").checked;

    if (!useTrend && !useCyclical && !useSeasonal && !useResidual) {
        alert("Please select at least one decomposition component.");
        return;
    }

    infoEl.innerHTML = `<p>Performing Time Series Decomposition...</p>`;
    document.getElementById("btnGenerateDecomp").disabled = true;
    document.getElementById("btnSkipDecomp").disabled = true;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(lastLoadedDataRange.sheetName);
            let targetTable = sheet.tables.getItem(lastLoadedDataRange.tableName);

            const fullTableRange = targetTable.getRange().load("values, rowCount, columnCount");
            targetTable.load("name, columns/items/name");
            await context.sync();

            const tableName = targetTable.name;
            const tableValues = fullTableRange.values;
            
            if (fullTableRange.columnCount < 2) {
                throw new Error("The loaded table does not have enough columns to perform decomposition.");
            }

            const dataRows = tableValues.slice(1);
            const info = lastLoadedSeries.info || {};
            const seriesIdFromHeader = lastLoadedSeries.seriesId;
            
            const freq = (info && info.frequency) ? info.frequency.toLowerCase() : "";

            let L = 1;
            if (freq.includes("month")) L = 12;
            else if (freq.includes("quarter")) L = 4;
            else if (freq.includes("week")) L = 52;
            else if (freq.includes("day") || freq.includes("daily")) L = 365;

            const values = dataRows.map(row => {
                const parsed = parseFloat(row[1]);
                return isNaN(parsed) ? 0 : parsed;
            });
            const n = values.length;

            function getCenteredMA(arr, span) {
                const result = new Array(arr.length).fill(null);
                if (span <= 1) return arr;
                let half = Math.floor(span / 2);
                for (let i = half; i < arr.length - half; i++) {
                    let sum = 0;
                    if (span % 2 === 0) {
                        for (let j = i - half + 1; j < i + half; j++) sum += arr[j];
                        sum += (arr[i - half] + arr[i + half]) / 2;
                        result[i] = sum / span;
                    } else {
                        for (let j = i - half; j <= i + half; j++) sum += arr[j];
                        result[i] = sum / span;
                    }
                }
                return result;
            }

            const maShort = getCenteredMA(values, Math.max(L, 3));
            const maLong = getCenteredMA(values, Math.max(L * 5, 5));

            const trend = new Array(n).fill(null);
            const cyclical = new Array(n).fill(null);
            const seasonal = new Array(n).fill(0);
            const residual = new Array(n).fill(null);

            for (let i = 0; i < n; i++) { if (maLong[i] !== null) trend[i] = maLong[i]; if (maShort[i] !== null && maLong[i] !== null) cyclical[i] = maShort[i] - maLong[i]; }

            if (useSeasonal) {
                const seasonSums = new Array(L).fill(0);
                const seasonCounts = new Array(L).fill(0);
                for (let i = 0; i < n; i++) { if (maShort[i] !== null) { seasonSums[i % L] += (values[i] - maShort[i]); seasonCounts[i % L]++; } }
                let seasonAvg = seasonSums.map((s, idx) => seasonCounts[idx] > 0 ? s / seasonCounts[idx] : 0);
                let meanSeason = seasonAvg.reduce((a,b)=>a+b, 0) / L;
                seasonAvg = seasonAvg.map(s => s - meanSeason);
                for (let i = 0; i < n; i++) seasonal[i] = seasonAvg[i % L];
            }

            for (let i = 0; i < n; i++) { if (maShort[i] !== null) { residual[i] = values[i] - maShort[i] - (useSeasonal ? seasonal[i] : 0); } }

            const existingCols = new Set(targetTable.columns.items.map(c => c.name));
                
            const trendVals = [["Trend"], ...trend.map(v => [v === null ? null : v])];
            const cyclicalVals = [["Cyclical"], ...cyclical.map(v => [v === null ? null : v])];
            const seasonalVals = [["Seasonal"], ...seasonal.map(v => [v === null ? null : v])];
            const residualVals = [["Residual"], ...residual.map(v => [v === null ? null : v])];
            
            let colsToAdd = [];
            if (useTrend && !existingCols.has("Trend")) colsToAdd.push({ name: "Trend", values: trendVals });
            if (useCyclical && !existingCols.has("Cyclical")) colsToAdd.push({ name: "Cyclical", values: cyclicalVals });
            if (useSeasonal && !existingCols.has("Seasonal")) colsToAdd.push({ name: "Seasonal", values: seasonalVals });
            if (useResidual && !existingCols.has("Residual")) colsToAdd.push({ name: "Residual", values: residualVals });

            const colsToAddCount = colsToAdd.length;

            if (colsToAddCount > 0) {
                // --- 1. Make Space by inserting entire columns ---
                // This is the most robust way to shift all content on the sheet and avoid breaking other tables.
                const tableRange = targetTable.getRange().load("columnIndex, columnCount, rowCount");
                await context.sync();
                const insertColIndex = tableRange.columnIndex + tableRange.columnCount;
                
                const columnsToInsert = sheet.getRangeByIndexes(0, insertColIndex, 1, colsToAddCount).getEntireColumn();
                columnsToInsert.insert(Excel.InsertShiftDirection.right);
                await context.sync();
                
                // --- 2. Re-fetch table reference as the insertion invalidates old ones ---
                targetTable = sheet.tables.getItem(tableName);

                // --- 3. Add empty columns first (structural change) ---
                for (const col of colsToAdd) {
                    targetTable.columns.add(null, null, col.name);
                }
                await context.sync();

                // --- 4. Populate the new columns with data ---
                for (const col of colsToAdd) {
                    const columnReference = targetTable.columns.getItem(col.name);
                    const bodyRange = columnReference.getDataBodyRange();
                    bodyRange.values = col.values.slice(1); // Data only, no header
                }
                await context.sync();
            }

            const newTableRange = targetTable.getRange();
            newTableRange.columnHidden = false; // Unhide columns 
            newTableRange.load("columnIndex, rowIndex, rowCount");
            await context.sync();

            const chartRange = newTableRange;
            const chart = sheet.charts.add(Excel.ChartType.line, chartRange, Excel.ChartSeriesBy.columns);
            chart.title.text = `${info.title || seriesIdFromHeader} - Decomposition`;

            // Position chart under the datasets' notes and aligned with the metadata
            const chartStartCol = lastLoadedDataRange.metadataStartColZeroBased;
            
            // Dynamically calculate position relative to previously created charts
            const metadataRowCount = lastLoadedDataRange.metadataRowCount;
            const histogramEndRow = metadataRowCount + 2 + 15 - 1;
            const lineChartEndRow = histogramEndRow + 2 + 15 - 1;
            const boxPlotEndRow = lineChartEndRow + 2 + 15 - 1;
            const textBoxEndRow = boxPlotEndRow + 2 + 10 - 1;
            const decompChartStartRow = textBoxEndRow + 4;
            
            chart.setPosition(sheet.getCell(decompChartStartRow, chartStartCol), sheet.getCell(decompChartStartRow + 15, chartStartCol + 3));

            await context.sync();
        });

        infoEl.innerHTML = `<p style="color: green;">Time Series Decomposition completed!</p>`;
    } catch (error) {
        console.error("Decomposition error:", error);
        infoEl.innerHTML = `<p>Error performing decomposition: ${error.message || error}</p>`;
    } finally {
        if (navStack[navStack.length - 1].type === 'decompPrompt') navStack.pop();
        promptForPivotTable(); // Proceed to the Pivot Table prompt unconditionally
    }
}

/* --------------------------------------------------------------------------
   EXCEL INSERTION FUNCTIONS
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

        // Extract notes content separately
        const notesContent = metadata.notes || null;

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("columnIndex, columnCount, isNullObject, width"); // Load width for dynamic sizing
            sheet.load("name"); // Load the name property of the sheet object
            await context.sync();

            let startColumnIndex;
            if (usedRange.isNullObject) {
                startColumnIndex = 1;
            } else {
                // Ensure enough space from the last used column
                startColumnIndex = usedRange.columnIndex + usedRange.columnCount + 2; 
            }

            const metadataArray = [];
            if (metadata) { // Only process if metadata exists
                // Define a preferred order for important fields to display them at the top
                const preferredOrder = [
                    'series_id', 'title', 'frequency', 'units', 'seasonal_adjustment', 
                    'observation_start', 'observation_end', 'last_updated', 'popularity'
                ];

                // Function to format key into a readable label
                const formatLabel = (key) => key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()) + ':';

                // Add preferred fields first, in order
                preferredOrder.forEach(key => {
                    if (Object.prototype.hasOwnProperty.call(metadata, key) && metadata[key]) {
                        metadataArray.push([formatLabel(key), String(metadata[key])]);
                    }
                });

                // Add remaining fields that are not in the preferred list
                for (const key in metadata) { // Iterate over all properties in metadata
                    if (Object.prototype.hasOwnProperty.call(metadata, key) && !preferredOrder.includes(key) && key !== 'notes' && metadata[key]) {
                        metadataArray.push([formatLabel(key), String(metadata[key])]);
                    }
                }
            }
            const dataValues = [["Date", currentSeriesId]];
            currentData.forEach(r => {
                dataValues.push([r.date, r.value]);
            });

            const metadataStartColZeroBased = startColumnIndex - 1;
            const metadataRange = sheet.getRangeByIndexes(0, metadataStartColZeroBased, metadataArray.length, 2);
            metadataRange.values = metadataArray;
            metadataRange.format.autofitColumns();
            metadataRange.getCell(0, 0).getResizedRange(metadataArray.length - 1, 0).format.font.bold = true;

            const dataStartColZeroBased = metadataStartColZeroBased + 4; // 2 columns for metadata + 2 empty columns
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
                seriesId: currentSeriesId,
                metadataStartColZeroBased: metadataStartColZeroBased, // Start column of the metadata table
                metadataRowCount: metadataArray.length, // Number of rows in the metadata table
                notesContent: notesContent // Store notes content for later use
            };
        });

        await createHistogram(currentSeriesName);
        await createLineChart(currentSeriesName);
        await createBoxPlot(currentSeriesName);
        await createNotesTextBox(notesContent); // Create the text box after all charts
        promptForDecomposition();
    } catch (error) {
        console.error("Error loading data into current sheet:", error);
        infoEl.innerHTML = `<p>Error loading data into current sheet: ${error.message || error}</p>`;
    } finally {
        btnLoadDataCurrentSheet.disabled = false;
        btnLoadDataNewSheet.disabled = false;
    }
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
        
        // Extract notes content separately
        const notesContent = metadata.notes || null;

        await Excel.run(async (context) => {
            let sheetName = currentSeriesId.replace(/[^a-zA-Z0-9 ]/g, "").substring(0, 31).trim();
            if (sheetName.length === 0) sheetName = "FRED_Data";
            
            let uniqueSheetName = sheetName;
            let counter = 1;
            while (true) {
                let existingSheet = context.workbook.worksheets.getItemOrNullObject(uniqueSheetName);
                await context.sync();
                if (existingSheet.isNullObject) {
                    break;
                }
                uniqueSheetName = sheetName.substring(0, 31 - String(counter).length - 1) + "_" + counter;
                counter++;
            }
            const newSheet = context.workbook.worksheets.add(uniqueSheetName);
            newSheet.activate(); // Activate the new sheet
            newSheet.load("name"); // Load the name property of the new sheet
            await context.sync(); // Sync to ensure newSheet.name is loaded

            const metadataArray = [];
            if (metadata) {
                // Define a preferred order for important fields to display them at the top
                const preferredOrder = [
                    'series_id', 'title', 'frequency', 'units', 'seasonal_adjustment', 
                    'observation_start', 'observation_end', 'last_updated', 'popularity', 'notes'
                ];

                // Function to format key into a readable label
                const formatLabel = (key) => key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()) + ':';

                // Add preferred fields first, in order
                preferredOrder.forEach(key => {
                    if (Object.prototype.hasOwnProperty.call(metadata, key) && metadata[key]) {
                        metadataArray.push([formatLabel(key), String(metadata[key])]);
                    }
                });

                // Add remaining fields that are not in the preferred list
                for (const key in metadata) {
                    if (Object.prototype.hasOwnProperty.call(metadata, key) && !preferredOrder.includes(key) && metadata[key]) {
                        metadataArray.push([formatLabel(key), String(metadata[key])]);
                    }
                }
            }
            const dataValues = [["Date", currentSeriesId]];
            currentData.forEach(r => {
                dataValues.push([r.date, r.value]);
            });

            const metadataRange = newSheet.getRangeByIndexes(0, 0, metadataArray.length, 2);
            metadataRange.values = metadataArray;
            metadataRange.format.autofitColumns();
            metadataRange.getCell(0, 0).getResizedRange(metadataArray.length - 1, 0).format.font.bold = true;

            const dataRange = newSheet.getRangeByIndexes(0, 4, dataValues.length, 2); // 2 columns for metadata + 2 empty columns
            dataRange.values = dataValues;
            dataRange.format.autofitColumns();
            dataRange.getRow(0).format.font.bold = true;

            const safeTableName = "Table_" + currentSeriesId.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);
            const table = newSheet.tables.add(dataRange, true);
            table.name = safeTableName;
            table.style = "TableStyleMedium2";

            // Define the target cell for the pivot table
            const pivotTargetCell = newSheet.getCell(0, 10); // Col K (4 for metadata/pad + 2 for data + 3 for hidden columns + 1 for padding)

            // Load the address properties before accessing them
            dataRange.load("address");
            pivotTargetCell.load("address");
            await context.sync(); // Synchronize to get the loaded addresses

            lastLoadedDataRange = {
                sheetName: newSheet.name,
                tableName: safeTableName,
                dataRangeAddress: dataRange.address, // This will now reflect the new starting column
                dataStartColZeroBased: 4,
                rowCount: dataValues.length,
                pivotTargetAddress: pivotTargetCell.address,
                seriesId: currentSeriesId,
                metadataStartColZeroBased: 0, // Metadata starts at column A (0-indexed) in a new sheet for new sheets
                metadataRowCount: metadataArray.length, // Number of rows in the metadata table
                notesContent: notesContent // Store notes content for later use
            };
        });

        await createHistogram(currentSeriesName);
        await createLineChart(currentSeriesName);
        await createBoxPlot(currentSeriesName);
        await createNotesTextBox(notesContent); // Create the text box after all charts
        promptForDecomposition();
    } catch (error) {
        console.error("Error loading data into new sheet:", error);
        infoEl.innerHTML = `<p>Error loading data into new sheet: ${error.message || error}</p>`;
    } finally {
        btnLoadDataCurrentSheet.disabled = false;
        btnLoadDataNewSheet.disabled = false;
    }
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

    const yearHtml = createCheckboxHtml("ptYear", "Year", canUseYear, checkYear);
    const quarterHtml = createCheckboxHtml("ptQuarter", "Quarter", canUseQuarter, checkQuarter);
    const monthHtml = createCheckboxHtml("ptMonth", "Month", canUseMonth, checkMonth);

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
    if (document.getElementById("ptYear") && document.getElementById("ptYear").checked) groupFields.push("Year");
    if (document.getElementById("ptQuarter") && document.getElementById("ptQuarter").checked) groupFields.push("Quarter");
    if (document.getElementById("ptMonth") && document.getElementById("ptMonth").checked) groupFields.push("Month");

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

            // Get the current range of the table, which might include decomposition columns
            const tableRange = table.getRange();
            tableRange.load("columnIndex, columnCount, rowCount");
            await context.sync();

            const dataStartCol = tableRange.columnIndex;
            const tableTotalCols = tableRange.columnCount;
            const rowCount = tableRange.rowCount; // This includes the header row

            const extendedCols = groupFields.length;

            // --- Create helper columns for pivot table ---
            const extraDataValues = [];
            extraDataValues.push(groupFields); // Write header

            // Pre-format month names to make PivotTable sorting alphabetical
            const monthNames = ["01 - Jan", "02 - Feb", "03 - Mar", "04 - Apr", "05 - May", "06 - Jun", "07 - Jul", "08 - Aug", "09 - Sep", "10 - Oct", "11 - Nov", "12 - Dec"];

            currentData.forEach(r => {
                const parts = r.date.split("-");
                const rowVals = [];
                if (parts.length === 3) {
                    const year = parseInt(parts[0], 10);
                    const month = parseInt(parts[1], 10);
                    const quarter = "Q" + Math.ceil(month / 3);
                    const monthName = monthNames[month - 1];

                    groupFields.forEach(field => {
                        if (field === "Year") rowVals.push(year);
                        else if (field === "Quarter") rowVals.push(quarter);
                        else if (field === "Month") rowVals.push(monthName);
                    });
                } else {
                    groupFields.forEach(() => rowVals.push(""));
                }
                extraDataValues.push(rowVals);
            });

            // Place helper columns immediately to the right of the existing table
            const helperColStart = dataStartCol + tableTotalCols;
            const extraRange = sheet.getRangeByIndexes(0, helperColStart, rowCount, extendedCols);
            extraRange.values = extraDataValues;
            extraRange.columnHidden = true;

            // The full range for the pivot table source now includes the original table AND the new helper columns
            const fullDataRange = sheet.getRangeByIndexes(0, dataStartCol, rowCount, tableTotalCols + extendedCols);

            // Place the pivot table to the right of the helper columns, with a gap
            const targetCell = sheet.getCell(0, helperColStart + extendedCols + 1);

            await context.sync();

            const safeName = "Pivot_" + lastLoadedDataRange.seriesId.replace(/[^a-zA-Z0-9]/g, "") + "_" + Math.floor(Math.random() * 1000);
            const pivotTable = sheet.pivotTables.add(safeName, fullDataRange, targetCell);

            groupFields.forEach(field => {
                pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
            });

            const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(lastLoadedDataRange.seriesId));
            dataHierarchy.summarizeBy = aggFn;
            pivotTable.layout.layoutType = Excel.PivotLayoutType.compact;

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
    currentSeriesId = lastLoadedSeries.seriesId;
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
        if (!lastLoadedDataRange) return; // Safety check
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const valueColumnIndex = lastLoadedDataRange.dataStartColZeroBased + 1; // Value column is 1 index after Date column
        const rowCount = lastLoadedDataRange.rowCount;
        const dataRange = sheet.getRangeByIndexes(1, valueColumnIndex, rowCount - 1, 1);

        const chart = sheet.charts.add(Excel.ChartType.histogram, dataRange, Excel.ChartSeriesBy.columns);
        chart.title.text = `${currentSeriesName || currentSeriesId} Histogram`;

        // Position the chart under its corresponding metadata.
        const chartStartCol = lastLoadedDataRange.metadataStartColZeroBased;
        const chartStartRow = lastLoadedDataRange.metadataRowCount + 2; // 2 rows below metadata
        const chartEndRow = chartStartRow + 15 - 1;   // 15 rows tall

        chart.setPosition(
            sheet.getCell(chartStartRow, chartStartCol),
            sheet.getCell(chartEndRow, chartStartCol + 3) // 4 columns wide (to match metadata width)
        );

        chart.axes.valueAxis.title.text = "Frequency";
        chart.axes.valueAxis.title.visible = true;
        chart.axes.categoryAxis.title.text = "Value Range";
        chart.axes.categoryAxis.title.visible = true;
        chart.axes.categoryAxis.format.font.size = 9;

        await context.sync();
    });
}

async function createNotesTextBox(notesContent) {
    if (!notesContent || !lastLoadedDataRange) return;

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const chartStartCol = lastLoadedDataRange.metadataStartColZeroBased;
        const metadataRowCount = lastLoadedDataRange.metadataRowCount;

        // Calculate end row of the last chart (Box Plot)
        const histogramEndRow = metadataRowCount + 2 + 15 - 1;
        const lineChartEndRow = histogramEndRow + 2 + 15 - 1;
        const boxPlotEndRow = lineChartEndRow + 2 + 15 - 1;

        const textBoxStartRow = boxPlotEndRow + 2; // 2 rows below the box plot

        const widthRange = sheet.getRangeByIndexes(0, chartStartCol, 1, 3);
        const startCell = sheet.getCell(textBoxStartRow, chartStartCol);

        widthRange.load("width");
        startCell.load("left, top");

        await context.sync();

        const shape = sheet.shapes.addTextBox(notesContent);
        shape.left = startCell.left;
        shape.top = startCell.top;
        shape.width = widthRange.width;
        shape.height = 150; // A reasonable default height, can be adjusted or made dynamic

        shape.textFrame.textRange.font.size = 9;
        shape.textFrame.textRange.font.name = "Calibri";
        shape.textFrame.wordWrap = true;
        shape.textFrame.verticalAlignment = Excel.VerticalAlignment.top;
        shape.textFrame.horizontalAlignment = Excel.HorizontalAlignment.left;

        await context.sync();
    });
}
async function createBoxPlot(currentSeriesName) {
    await Excel.run(async (context) => {
        if (!lastLoadedDataRange) return; // Safety check
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Box plots typically use the value column to show statistical distribution
        const valueColumnIndex = lastLoadedDataRange.dataStartColZeroBased + 1;
        const rowCount = lastLoadedDataRange.rowCount;
        const dataRange = sheet.getRangeByIndexes(1, valueColumnIndex, rowCount - 1, 1);

        const chart = sheet.charts.add(Excel.ChartType.boxwhisker, dataRange, Excel.ChartSeriesBy.columns);

        chart.title.text = `${currentSeriesName || "Series"} Distribution (Box Plot)`;

        // Position the chart under the line chart
        const chartStartCol = lastLoadedDataRange.metadataStartColZeroBased;
        const histogramEndRow = lastLoadedDataRange.metadataRowCount + 2 + 15 - 1;
        const lineChartEndRow = histogramEndRow + 2 + 15 - 1;
        const boxPlotStartRow = lineChartEndRow + 2; 
        const boxPlotEndRow = boxPlotStartRow + 15 - 1;

        chart.setPosition(
            sheet.getCell(boxPlotStartRow, chartStartCol),
            sheet.getCell(boxPlotEndRow, chartStartCol + 3) // 4 columns wide
        );

        await context.sync();
    });
}

async function createLineChart(currentSeriesName) {
    await Excel.run(async (context) => {
        if (!lastLoadedDataRange) return; // Safety check
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // 1. Identify the data (X = Dates, Y = Values)
        const startCol = lastLoadedDataRange.dataStartColZeroBased;
        const rowCount = lastLoadedDataRange.rowCount;
        
        // Line charts need two columns for time series: Date and Value.
        const dataRange = sheet.getRangeByIndexes(1, startCol, rowCount - 1, 2);

        // 2. Add the Chart (Line)
        const chart = sheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.columns);

        chart.title.text = `${currentSeriesName || "Series"} Trend (Line Chart)`;

        // 3. Position the chart under the histogram for the same dataset.
        const chartStartCol = lastLoadedDataRange.metadataStartColZeroBased;
        const histogramEndRow = lastLoadedDataRange.metadataRowCount + 2 + 15 - 1; // End row of the histogram
        const lineChartStartRow = histogramEndRow + 2; // 2 rows below the histogram
        const lineChartEndRow = lineChartStartRow + 15 - 1;   // 15 rows tall

        chart.setPosition(
            sheet.getCell(lineChartStartRow, chartStartCol),
            sheet.getCell(lineChartEndRow, chartStartCol + 3) // 4 columns wide, same as histogram
        );

        // 4. Formatting
        chart.axes.valueAxis.title.text = "Value";
        chart.axes.valueAxis.title.visible = true;
        chart.axes.categoryAxis.title.text = "Date";
        chart.axes.categoryAxis.title.visible = true;
        chart.legend.visible = false; // Only one series, so legend is not needed.

        await context.sync();
    });
}
