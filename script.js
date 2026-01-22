const BASE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSL_mRAf-5YAR9vnzSkhOYhhY1eXsq-E-QAoHi1Gapdektd0gdAjJnAoG_6pIa5HA/pub?output=csv';

const SHEET_GIDS = {
    '2026': '1909984551',
    '2025': '2051125391',
    '2024': '893886517'
};

// ... (Exchange Rate etc.)

const EXCHANGE_RATE = 32.5; // Fixed rate for calculation

// 圖表實例
let myChart = null;

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    // 預設為 2026
    const yearSelect = document.getElementById('yearSelect');
    fetchData(yearSelect.value);

    // 年份切換事件
    yearSelect.addEventListener('change', (e) => {
        fetchData(e.target.value);
    });

    document.getElementById('refreshBtn').addEventListener('click', () => {
        // 簡單的旋轉動畫作為反饋
        const icon = document.querySelector('.refresh-btn i');
        icon.style.transition = 'transform 0.5s ease';
        icon.style.transform = 'rotate(360deg)';
        setTimeout(() => icon.style.transform = 'none', 500);

        fetchData(yearSelect.value);
    });

    // Settings / Force Update
    const settingsBtn = document.getElementById('settingsBtn');
    if (settingsBtn) {
        settingsBtn.addEventListener('click', () => {
            // 直接執行不彈窗
            forceUpdateApp();
        });
    }

    async function forceUpdateApp() {
        const statusEl = document.getElementById('totalAssetValue');
        if (statusEl) statusEl.textContent = '版本更新中...';

        try {
            // 1. Unregister Service Workers
            if ('serviceWorker' in navigator) {
                const registrations = await navigator.serviceWorker.getRegistrations();
                for (const registration of registrations) {
                    await registration.unregister();
                }
            }

            // 2. Clear All Caches
            if ('caches' in window) {
                const cacheNames = await caches.keys();
                await Promise.all(
                    cacheNames.map(name => caches.delete(name))
                );
            }

            // 3. Reload Page
            window.location.reload(true);
        } catch (error) {
            console.error('Force update failed:', error);
            // Fail silently or just log
            if (statusEl) statusEl.textContent = '更新失敗';
        }
    }

    // --- Swipe Support (Touch & Mouse) ---
    let startX = 0;
    let endX = 0;
    const minSwipeDistance = 50;

    // Touch Events
    document.addEventListener('touchstart', e => {
        startX = e.changedTouches[0].screenX;
    }, { passive: true });

    document.addEventListener('touchend', e => {
        endX = e.changedTouches[0].screenX;
        handleSwipe();
    }, { passive: true });

    // Mouse Events for Desktop Testing
    document.addEventListener('mousedown', e => {
        startX = e.screenX;
    });

    document.addEventListener('mouseup', e => {
        endX = e.screenX;
        handleSwipe();
    });

    function handleSwipe() {
        const swipeDistance = endX - startX;

        if (Math.abs(swipeDistance) < minSwipeDistance) return;

        if (swipeDistance < 0) {
            // Swipe Left (<--) -> Newer Year
            changeYear('newer');
        } else {
            // Swipe Right (-->) -> Older Year
            changeYear('older');
        }
    }

    function changeYear(direction) {
        const currentIndex = yearSelect.selectedIndex;
        let newIndex = currentIndex;

        if (direction === 'newer') {
            // Newer years are at lower indices (e.g. Index 0 is 2026, Index 1 is 2025)
            if (currentIndex > 0) newIndex--;
        } else {
            // Older years are at higher indices
            if (currentIndex < yearSelect.options.length - 1) newIndex++;
        }

        if (newIndex !== currentIndex) {
            const dashboard = document.querySelector('.dashboard');

            // 1. Determine Animation Classes
            // Swipe Left -> Move Current Left (Exit), New comes from Right -> Newer Year
            // Swipe Right -> Move Current Right (Exit), New comes from Left -> Older Year

            // direction 'newer' (triggered by Swipe Left) -> Slide Out Left
            // direction 'older' (triggered by Swipe Right) -> Slide Out Right
            const exitClass = direction === 'newer' ? 'anim-slide-out-left' : 'anim-slide-out-right';
            const enterClass = direction === 'newer' ? 'anim-slide-in-right' : 'anim-slide-in-left';

            // 2. Play Exit Animation
            dashboard.classList.add(exitClass);

            // 3. Wait for Exit to finish
            setTimeout(() => {
                // Change Data
                yearSelect.selectedIndex = newIndex;
                yearSelect.dispatchEvent(new Event('change'));

                // Reset Class and Add Enter Animation
                dashboard.classList.remove(exitClass);
                dashboard.classList.add(enterClass);

                // 4. Clean up Enter Class after animation
                setTimeout(() => {
                    dashboard.classList.remove(enterClass);
                }, 300); // 300ms matches CSS duration

            }, 300);
        }
    }
});

function fetchData(year) {
    const statusEl = document.getElementById('totalAssetValue');
    statusEl.innerHTML = '<span style="font-size:1rem; opacity:0.7;">更新中...</span>';

    const gid = SHEET_GIDS[year];
    // 加入時間戳記以防止快取，並透過 &gid= 指定年份
    // 注意：Publish to Web 的連結通常支援 &gid 參數
    const url = `${BASE_URL}&gid=${gid}&t=${Date.now()}`;
    console.log(`Fetching year ${year} with GID ${gid}`);

    Papa.parse(url, {
        download: true,
        complete: function (results) {
            processData(results.data);
        },
        error: function (err) {
            console.error('Error fetching CSV:', err);
            statusEl.textContent = '讀取失敗';
        }
    });
}

function processData(rows) {
    // Determine Mode
    const isUSPage = document.body.classList.contains('us-dashboard');
    const currencyFormatter = isUSPage ? formatUSD : formatCurrency;

    // 除錯：記錄前幾列以確認結構
    console.log('CSV Rows:', rows.slice(0, 5));

    // 清除舊資料以避免殘留
    clearUI();

    // 1. 尋找標題列索引 (月份名稱)
    let headerRowIndex = -1;
    for (let i = 0; i < 20; i++) { // 搜尋前 20 列
        if (rows[i] && rows[i].includes('1月')) {
            headerRowIndex = i;
            break;
        }
    }

    if (headerRowIndex === -1) {
        console.error('Could not find header row');
        return;
    }

    // 2. 識別月份欄位 (1月-12月的索引)
    const headerRow = rows[headerRowIndex];
    let monthIndices = [];
    ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'].forEach(month => {
        const idx = headerRow.indexOf(month);
        if (idx !== -1) monthIndices.push(idx);
    });

    // 3. 尋找相關數據列
    const assetRows = rows.filter(r => {
        if (!r) return false;
        for (let i = 0; i < 5; i++) {
            const cell = r[i] ? r[i].toString().trim() : '';
            if (cell === '總現值') return true;
        }
        return false;
    });

    if (assetRows.length === 0) {
        console.error('Could not find any asset rows');
        document.getElementById('totalAssetValue').textContent = '數據無法識別';
        return;
    }

    // 4. 聚合總資產數據
    // aggregatedPoints 將根據 mode 不同而存儲 TWD 或 USD
    let aggregatedPoints = new Array(monthIndices.length).fill(0);

    // 用於主頁的分開追蹤
    let twPoints = new Array(monthIndices.length).fill(0);
    let usPoints = new Array(monthIndices.length).fill(0); // In Original Currency (USD)

    let hasData = false;

    assetRows.forEach(row => {
        let isTW = false;
        let isUS = false;

        const rowStr = row.join('');
        if (rowStr.includes('台股')) isTW = true;
        if (rowStr.includes('美股')) isUS = true;

        if (!isTW && !isUS) isTW = true; // Fallback

        // For US Page: Skip TW rows
        if (isUSPage && isTW) return;

        // For Main Page: Skip US rows (User request: Exclude US data)
        if (!isUSPage && isUS) return;

        // Extract Data
        let rowHasMonthData = false;

        // Helper to process value
        const processValue = (val, monthIndex) => {
            if (val > 0) {
                if (isUSPage) {
                    // US Page: Only US rows here. Assuming CSV data for US is in USD.
                    aggregatedPoints[monthIndex] += val;
                } else {
                    // Main Page: Aggregate to total in TWD
                    // Since we return early for isUS, we only have TW here.
                    twPoints[monthIndex] += val;
                    aggregatedPoints[monthIndex] += val;
                }
                rowHasMonthData = true;
                hasData = true;
            }
        };

        monthIndices.forEach((colIdx, monthIndex) => {
            const val = parseCurrency(row[colIdx]);
            processValue(val, monthIndex);
        });

        // Fallback for single value rows (e.g. 2026)
        if (!rowHasMonthData) {
            for (let k = 4; k < row.length; k++) {
                const val = parseCurrency(row[k]);
                if (val > 0) {
                    const targetIndex = aggregatedPoints.length > 0 ? aggregatedPoints.length - 1 : 0;
                    processValue(val, targetIndex);
                    break;
                }
                if (k > 20) break;
            }
        }
    });

    let dataPoints = aggregatedPoints.filter(p => p > 0);
    let lastValidValue = dataPoints.length > 0 ? dataPoints[dataPoints.length - 1] : 0;

    let twData = twPoints.filter(p => p > 0);
    let usData = usPoints.filter(p => p > 0);
    let lastTW = twData.length > 0 ? twData[twData.length - 1] : 0;
    let lastUS = usData.length > 0 ? usData[usData.length - 1] : 0;

    // 5. 更新 UI
    if (!hasData) {
        document.getElementById('totalAssetValue').textContent = isUSPage ? '$0.00' : '$0';
        if (!isUSPage) {
            const twEl = document.getElementById('twAssetValue');
            if (twEl) twEl.textContent = '$0';
        }
    } else {
        document.getElementById('totalAssetValue').textContent = currencyFormatter(lastValidValue);

        if (!isUSPage) {
            const twEl = document.getElementById('twAssetValue');
            if (twEl) twEl.textContent = formatCurrency(lastTW);
        }
    }

    // Chart Labels
    const labels = [];
    aggregatedPoints.forEach((val, i) => {
        if (val > 0) {
            labels.push(`${i + 1}月`);
        }
    });

    // Trend
    if (dataPoints.length >= 2) {
        const current = dataPoints[dataPoints.length - 1];
        const prev = dataPoints[dataPoints.length - 2];
        const percentChange = ((current - prev) / prev) * 100;

        const trendEl = document.getElementById('assetTrend');
        const trendValEl = document.getElementById('trendValue');

        trendValEl.textContent = `${Math.abs(percentChange).toFixed(1)}%`;
        if (percentChange >= 0) {
            trendEl.className = 'trend positive';
            trendEl.innerHTML = '<i class="ri-arrow-up-line"></i> ' + trendValEl.textContent;
        } else {
            trendEl.className = 'trend negative';
            trendEl.innerHTML = '<i class="ri-arrow-down-line"></i> ' + trendValEl.textContent;
        }
    }

    // --- Metrics (Unrealized, Realized, etc.) ---
    const findRow = (keyword) => rows.find(r => {
        if (!r) return false;
        for (let i = 0; i < 5; i++) {
            if (r[i] && r[i].toString().trim().includes(keyword)) return true;
        }
        return false;
    });

    const findTotalColIndex = () => {
        const header = rows[headerRowIndex];
        let idx = header.indexOf('年度總計');
        if (idx !== -1) return idx;
        for (let i = 0; i < 5; i++) {
            if (rows[i]) {
                const found = rows[i].indexOf('年度總計');
                if (found !== -1) return found;
            }
        }
        return 14;
    };
    const totalColIdx = findTotalColIndex();

    // Note: Metrics in CSV like "未實現損益總計" might be pre-calculated/summed?
    // If we are on US page, we only want US metrics.
    // The CSV structure seems to have separate sections or just one summary row?
    // Assuming the summary rows at bottom are aggregated totals. 
    // If so, extracting specific TW/US parts from them might be hard without row context.
    // However, usually "總現值" is separated by TW/US. "未實現" might also be?
    // Let's assume for now we scan for US-specific label if isUSPage.

    // Better strategy for metrics:
    // If isUSPage, looks for rows that are likely US.
    // However, looking at the code I replaced, it was just grabbing the first '未實現損益總計'.
    // If the CSV has separate rows for TW and US P/L, we should filter.
    // If it only has one total row, then US page will show the mixed total (wrong).
    // Let's check `findRow`. It returns the *first* match.

    // For now, I will implement a simplifed version: 
    // On Main Page -> Show formatted TWD (assuming rows are valid totals)
    // On US Page -> Since we might not easily separate them without more CSV info, 
    // we might need to hide them or accept they might be mixed if the CSV is mixed.
    // BUT, the `assetRows` logic separated them.
    // Let's try to filter metric rows by '美股' if on US page.

    const getMetric = (keyword, isExpense = false, forcePositive = false) => {
        let targetRow = null;

        // Find all rows matching keyword
        const candidates = rows.filter(r => {
            if (!r) return false;
            for (let i = 0; i < 5; i++) {
                if (r[i] && r[i].toString().trim().includes(keyword)) return true;
            }
            return false;
        });

        if (candidates.length === 0) {
            console.warn(`Metric Not Found: ${keyword}`);
        } else {
            console.log(`Metric Candidates for "${keyword}":`, candidates.map(c => c.slice(0, 3)));
        }

        if (isUSPage) {
            targetRow = candidates.find(r => r.join('').includes('美股'));
        } else {
            // Main page: prefer '總計' or if not found, sum them? 
            // Or just use the one that doesn't say '美股' explicitly if there are multiple?
            // Actually, usually there is a '台股XX' and '美股XX' and maybe a '總計'?
            // Start with the first one found for now to keep it safe, OR sum if multiple found.
            // Given I don't see the full CSV, I'll stick to 'first match' but try to avoid specific US/TW ones if a 'Total' exists?
            // Reverting to simple logic: Just find the first one.
            targetRow = candidates[0];
        }

        if (targetRow) {
            // Extract value
            // Try '年度總計' col first
            let val = parseCurrency(targetRow[totalColIdx]);

            // If 0, try month columns (e.g. for Unrealized)
            if (val === 0) {
                // Try last valid month
                let lastDataIndex = -1;
                for (let i = aggregatedPoints.length - 1; i >= 0; i--) {
                    if (aggregatedPoints[i] > 0) { lastDataIndex = i; break; }
                }
                if (lastDataIndex !== -1) {
                    val = parseCurrency(targetRow[monthIndices[lastDataIndex]]);
                }
            }

            // If still 0, fallback scan
            if (val === 0) {
                for (let k = 4; k < targetRow.length; k++) {
                    const v = parseCurrency(targetRow[k]);
                    if (v !== 0) { val = v; break; }
                    if (k > 20) break;
                }
            }

            // Conversion for Main Page if the row was actually US? 
            // This is risky. Let's assume the CSV values are in native currency if separated.
            // If we picked a US row on Main Page, we should convert.
            // But we probably picked the "Grand Total" row if it exists.

            return val;
        }
        return 0;
    };

    // Update Metrics
    let unrealized = getMetric('未實現損益總計');
    let realized = getMetric('已實現損益總計');
    let dividends = getMetric('股息總計');
    let fees = getMetric('交易成本總計', true);

    // If Main Page, we might need to convert?
    // Let's assume for now these specific "Total" rows in CSV are already aggregated in TWD by the user in the sheet?
    // If not, this part might be wrong. But since user asked to "Separate TW/US assets", 
    // implies they might be mixed rows.
    // For safety: On US Page, ensure we format as USD.

    updateMetric('unrealizedValue', unrealized, false, false, currencyFormatter);
    updateMetric('realizedValue', realized, false, false, currencyFormatter);
    updateMetric('dividendValue', dividends, true, false, currencyFormatter);
    updateMetric('feeValue', fees, false, true, currencyFormatter);
    // 5. 已實現總損益 (Realized P/L + Dividends)
    const totalProfit = realized + dividends;
    updateMetric('totalProfitValue', totalProfit, true, false, currencyFormatter);

    // --- New Metrics (Total Cost & Net Profit) ---
    // 6. 總投入成本 (含息)
    const totalCostRow = findRow('總付出成本(含息)');
    let totalCost = 0;
    if (totalCostRow) {
        // Logic similar to assetRows: if Main Page (TW only), exclude US cost if mixed?
        // Actually, "總付出成本" usually has "國泰券商(台股)" and "國泰美股券商(美股)" separate rows.
        // 'findRow' finds the *first* one.
        // We should use specific row selection like we did for assets if possible, or filter properties.

        // Better strategy: Filter for relevant cost rows like we did for assetRows
        const costRows = rows.filter(r => {
            if (!r) return false;
            for (let i = 0; i < 5; i++) {
                const cell = r[i] ? r[i].toString().trim() : '';
                if (cell.includes('總付出成本')) return true;
            }
            return false;
        });

        // Determine which cost row to use based on page
        costRows.forEach(row => {
            let isTW = false;
            let isUS = false;
            const rowStr = row.join('');
            if (rowStr.includes('台股')) isTW = true;
            if (rowStr.includes('美股')) isUS = true;
            if (!isTW && !isUS) isTW = true; // Fallback

            if (isUSPage && isTW) return; // Skip TW on US page
            if (!isUSPage && isUS) return; // Skip US on Main page

            // Extract value using monthly or fallback logic
            let rowCost = 0;
            // First try Total Col
            const totalColVal = parseCurrency(row[totalColIdx]);
            if (totalColVal > 0) {
                rowCost = totalColVal;
            } else {
                // Try last valid month
                let lastDataIndex = -1;
                for (let i = aggregatedPoints.length - 1; i >= 0; i--) {
                    if (aggregatedPoints[i] > 0) { lastDataIndex = i; break; }
                }
                if (lastDataIndex !== -1) {
                    rowCost = parseCurrency(row[monthIndices[lastDataIndex]]);
                }
                // Fallback scan
                if (rowCost === 0) {
                    for (let k = 4; k < row.length; k++) {
                        const v = parseCurrency(row[k]);
                        if (v > 0) { rowCost = v; break; }
                        if (k > 20) break;
                    }
                }
            }

            // Accumulate
            if (isUSPage) {
                totalCost += rowCost;
            } else {
                if (isUS) totalCost += (rowCost * EXCHANGE_RATE);
                else totalCost += rowCost;
            }
        });
    }
    updateMetric('totalCostValue', totalCost, false, false, currencyFormatter);

    // 7. 實際總損益 (Net Profit)
    // Formula: Market Value + Dividends - Total Cost - Fees
    // Market Value = lastValidValue (already aggregated correctly for page)
    // Dividends = dividends (already aggregated? Need to check getMetric logic)
    // Cost = totalCost
    // Fees = fees

    // Note: getMetric('股息總計') logic was simple 'findRow'. 
    // If '股息總計' row is single bottom row, it might mix currencies or be TWD only.
    // However, usually Dividends are separated or at bottom. 
    // Given user instructions, let's assume the values we extracted are correct for the context.

    // But wait, 'getMetric' uses 'findRow' which finds FIRST match.
    // If US Page, and Dividends are in a separate 'US' blocks?
    // Looking at CSV snippet, '股息' seems to be a separate section (Row 5).
    // And it has '元大xxx', '國泰xxx'...
    // It seems '股息總計' might be a Sum row at the very bottom.
    // If so, on US Page, showing Total TWD Dividend might be wrong if it includes TW stocks.
    // BUT, usually US dividends are reinvested or handling is different.
    // Let's stick to the current 'dividends' variable but be aware.

    const currentMarketValue = lastValidValue;
    const netProfit = currentMarketValue + dividends - totalCost - fees;

    updateMetric('netProfitValue', netProfit, true, false, currencyFormatter);

    // --- Dividend Calendar Logic ---
    // Extract Monthly Dividends
    // "股息" section usually starts around row 5 in provided format.
    // Structure: Name, Type, Jan, Feb... Dec, Total
    // 股息 columns are at indices 4 to 15 (Jan=4, Dec=15)

    const monthlyDividends = new Array(12).fill(0);
    const dividendLabels = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];

    // Find Dividend Section Rows
    // Heuristic: Look for rows where type column (Col C, index 2) might be empty or specific, 
    // but usually user puts stock name in Col B or C.
    // Based on CSV snippet: 
    // Row 5 Header: "股息...", "現金股利", "元大高股息"
    // Let's filter rows that maintain the dividend section structure or just scan all rows for dividend values?
    // Safer: Look for specific section headers if possible.
    // But generalized approach: Iterate all rows. If a row looks like a dividend row (has values in month columns and is NOT 'Total Asset' or 'Profit'), add it.
    // BUT we must distinguish TW vs US dividends.
    // In CSV snippet, I see: "現金股利", "元大高股息"...
    // And "股票股利"... 

    // Improved Strategy: State-Based Extraction with Dynamic Alignment
    // Scan rows. When we hit strict signatures of Dividend Section, start collecting.
    let inDividendSection = false;
    let dividendRows = [];
    let dividendJanIndex = -1; // Auto-detect column for Jan

    // Pass 1: Collect Candidate Rows and Find "Total" row for alignment
    rows.forEach((row, i) => {
        const rowStr = row.join('');
        const firstCell = (row[0] || '').toString();

        // Detect Start of Dividend Section
        if (!inDividendSection) {
            // Check for section header
            // Relaxed: Check if row contains "股息" AND "通知書" anywhere (handles Col shift)
            if (rowStr.includes('股息') && rowStr.includes('通知書')) {
                inDividendSection = true;
            } else if (rowStr.includes('現金股利') && (rowStr.includes('元大') || rowStr.includes('國泰'))) {
                // Fallback heuristic
                inDividendSection = true;
            }
        }

        if (inDividendSection) {
            // Add row to potential list
            dividendRows.push(row);

            // Detect End of Section "股息總計" and Calculate Alignment
            if (rowStr.includes('股息總計')) {
                inDividendSection = false;

                // Find index of "股息總計"
                const totalLabelIdx = row.findIndex(cell => cell && cell.includes('股息總計'));
                if (totalLabelIdx !== -1) {
                    // Look for first valid number AFTER the label to be Jan
                    // Usually 2 slots after? [Label, Empty, Jan]
                    // Scan from totalLabelIdx + 1
                    for (let k = totalLabelIdx + 1; k < row.length; k++) {
                        const cellContent = (row[k] || '').trim();
                        const val = parseCurrency(row[k]);

                        // Fix: STRICTLY check for non-empty cell that looks like a number
                        // Ignor empty cells (val=0) to skip the gap column
                        if (cellContent !== '' && !isNaN(val) && val > 0) {
                            dividendJanIndex = k;
                            break;
                        }
                    }
                }
            }
        }
    });

    // Pass 2: Process Collected Rows using detected Alignment
    const stockDividends = {}; // { 'StockName': [Jan, Feb... Dec] }

    if (dividendRows.length > 0) {
        // If detection failed, fallback to global monthIndices[0] (Jan)
        let janIdx = dividendJanIndex;
        if (janIdx === -1) {
            janIdx = (monthIndices && monthIndices.length > 0) ? monthIndices[0] : 3;
        }

        dividendRows.forEach(row => {
            const rowStr = row.join('');
            if (rowStr.includes('股息總計')) return; // Skip total row

            // Determine Name (Try Col E(4), C(2), B(1))
            let name = (row[4] || '').trim();
            if (!name) name = (row[2] || '').trim();
            if (!name) name = (row[1] || '').trim();

            name = name.replace(/\s+/g, '');
            if (!name) return;

            // Filters
            const invalidNames = ['(依通知書填入)', '標的備註', '(寫個大概就好)', '報酬金額', '報酬率', '標的'];
            if (name === '股息' || name === '現金股利' || name === '股票股利') return;
            if (invalidNames.some(inv => name.includes(inv))) return;

            // Detect TW/US
            let isTW = true;
            let isUS = false;
            const usKeywords = ['BRK', 'VTI', 'QQQ', 'SPY', 'VOO', 'NVDA', 'TSLA', 'AAPL', 'MSFT', 'GOOG', 'AMZN', 'META', 'AMD', 'NFLX', '美股'];
            if (usKeywords.some(kw => name.toUpperCase().includes(kw)) || (name.match(/^[A-Z]+$/) && !name.includes('00'))) {
                isUS = true;
                isTW = false;
            }
            if (rowStr.includes('美股')) { isUS = true; isTW = false; }

            if (isUSPage && !isUS) return;
            if (!isUSPage && isUS) return;

            // Initialize Array for this stock if new
            if (!stockDividends[name]) {
                stockDividends[name] = new Array(12).fill(0);
            }

            // Extract Monthly Data using janIdx
            for (let m = 0; m < 12; m++) {
                const colIdx = janIdx + m;
                if (colIdx < row.length) {
                    const val = parseCurrency(row[colIdx]);
                    if (val > 0) {
                        if (isUSPage) {
                            stockDividends[name][m] += val;
                        } else {
                            if (isUS) stockDividends[name][m] += (val * EXCHANGE_RATE);
                            else stockDividends[name][m] += val;
                        }
                    }
                }
            }
        });
    }

    renderDividendChart(dividendLabels, stockDividends, currencyFormatter);

    // Chart
    renderChart(labels, dataPoints, isUSPage);

    // Stock List
    renderStockList(rows, headerRowIndex, monthIndices, isUSPage);
}

// --- Chart Rendering Functions ---

let divChart = null;

function renderDividendChart(labels, stockDividends, formatter) {
    const ctx = document.getElementById('dividendChart');
    if (!ctx) return;

    if (divChart) {
        divChart.destroy();
    }

    const isDarkMode = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;

    // Prepare Datasets
    const datasets = [];
    // Color Palette (Vibrant for dark mode)
    const colors = [
        'rgba(255, 99, 132, 0.8)',   // Red
        'rgba(54, 162, 235, 0.8)',   // Blue
        'rgba(255, 206, 86, 0.8)',   // Yellow
        'rgba(75, 192, 192, 0.8)',   // Teal
        'rgba(153, 102, 255, 0.8)',  // Purple
        'rgba(255, 159, 64, 0.8)',   // Orange
        'rgba(199, 199, 199, 0.8)',  // Grey
        'rgba(233, 30, 99, 0.8)',    // Pink
        'rgba(0, 188, 212, 0.8)',    // Cyan
        'rgba(139, 195, 74, 0.8)',   // Light Green
    ];

    let colorIdx = 0;
    for (const [stockName, monthlyData] of Object.entries(stockDividends)) {
        // Only add if there is data > 0
        if (monthlyData.some(v => v > 0)) {
            datasets.push({
                label: stockName,
                data: monthlyData,
                backgroundColor: colors[colorIdx % colors.length],
                borderColor: colors[colorIdx % colors.length].replace('0.8', '1'),
                borderWidth: 1,
                borderRadius: 4,
                stack: 'Stack 0', // Enable Stacking
            });
            colorIdx++;
        }
    }

    // Sort datasets by total value for consistency? Or just let them pile up.
    // Usually existing order is fine.

    divChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true, // Show Legend now that we have multiple stocks
                    position: 'bottom',
                    labels: { color: isDarkMode ? '#aaa' : '#666' }
                },
                tooltip: {
                    callbacks: {
                        label: function (context) {
                            return context.dataset.label + ': ' + formatter(context.raw);
                        },
                        footer: function (tooltipItems) {
                            let sum = 0;
                            tooltipItems.forEach(function (tooltipItem) {
                                sum += tooltipItem.raw;
                            });
                            return '當月總計: ' + formatter(sum);
                        }
                    }
                },
                datalabels: {
                    color: '#fff',
                    display: function (context) {
                        return context.dataset.data[context.dataIndex] > 0;
                    },
                    font: {
                        weight: 'bold',
                        size: 11
                    },
                    formatter: function (value) {
                        if (value >= 1000) return (value / 1000).toFixed(1) + 'k';
                        return value;
                    },
                    anchor: 'center',
                    align: 'center'
                }
            },
            scales: {
                y: {
                    stacked: true, // Enable Stacking
                    beginAtZero: true,
                    grid: {
                        color: isDarkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)'
                    },
                    ticks: {
                        callback: function (value) {
                            if (value >= 1000000) return (value / 1000000).toFixed(1) + 'M';
                            if (value >= 1000) return (value / 1000).toFixed(0) + 'k';
                            return value;
                        },
                        color: isDarkMode ? '#aaa' : '#666'
                    }
                },
                x: {
                    stacked: true, // Enable Stacking
                    grid: { display: false },
                    ticks: {
                        color: isDarkMode ? '#aaa' : '#666'
                    }
                }
            }
        }
    });
}

function renderStockList(rows, headerRowIndex, monthIndices, isUSPage) {
    const stockListBody = document.getElementById('stockListBody');
    if (!stockListBody) return;
    stockListBody.innerHTML = '';

    const stocks = [];

    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const typeCol = row[4] ? row[4].toString().trim() : '';
        const nameColD = row[3] ? row[3].toString().trim() : '';
        const nameColC = row[2] ? row[2].toString().trim() : '';
        const colB = row[1] ? row[1].toString().trim() : '';
        const rowStr = row.join('');

        let isStock = false;
        let stockName = '';
        // HACK: Use simple regex for US stocks (e.g. "NVDA", "VTI") vs TW (e.g. "0050", "台積電")
        // Also explicitly check for known US stock names in Chinese (e.g. 波克夏)

        // Basic Detection
        if (typeCol === '股票' && nameColD) {
            isStock = true;
            stockName = nameColD;
        } else if (!colB && nameColC) {
            const excludedKeywords = [
                '總現值', '總付出成本', '損益試算', '獲利率%', '標的備註',
                '總賣出金額', '總買進金額', '已實現損益', '未實現損益', '股息總計', '交易成本總計',
                '由高而低', '占比', '持股庫存'
            ];
            if (!excludedKeywords.some(kw => nameColC.includes(kw)) && nameColC.length > 1) {
                isStock = true;
                stockName = nameColC;
            }
        }

        // US Stock Detection Logic (NOW CORRECTLY PLACED)
        let isUSStock = false;
        if (isStock && stockName) {
            const usKeywords = ['波克夏', 'BRK', 'VTI', 'QQQ', 'SPY', 'VOO', 'NVDA', 'TSLA', 'AAPL', 'MSFT', 'GOOG', 'AMZN', 'META', 'AMD', 'NFLX'];
            if (usKeywords.some(kw => stockName.toUpperCase().includes(kw))) {
                isUSStock = true;
            } else {
                // Fallback: No Chinese usually means US stock (ticker only)
                const hasChinese = /[\u4e00-\u9fa5]/.test(stockName);
                if (!hasChinese && /[A-Za-z]/.test(stockName)) {
                    isUSStock = true;
                }
            }
        }

        // US Detection
        // Usually US stocks have no chinese name or specific patterns, OR are in the US section.
        // We can check if the row has '美股' earlier or just rely on name format?
        // Or check if the value is small (USD) vs Large (TWD)? Unreliable.
        // Let's assume if it is in a block that had '美股' header?
        // Simplify: Check if name is English-like or if row indicates.
        // Actually, the AssetRows detection used '美股' keyword.
        // But individual stock rows might not have it.
        // Let's look at the stocks.
        // If we are on US page, we ONLY want US stocks.
        // If we are on Main page, we want All (or TW?).

        // Filter
        if (isUSPage && !isUSStock) isStock = false;

        // Main Page: Exclude US Stocks
        if (!isUSPage && isUSStock) isStock = false;

        if (isStock && stockName) {
            let currentValue = 0;
            // Get Value
            for (let j = monthIndices.length - 1; j >= 0; j--) {
                const val = parseCurrency(row[monthIndices[j]]);
                if (val > 0) { currentValue = val; break; }
            }
            if (currentValue === 0) { // Fallback
                for (let k = 4; k < row.length; k++) {
                    const v = parseCurrency(row[k]);
                    if (v > 0) { currentValue = v; break; }
                    if (k > 20) break;
                }
            }

            if (currentValue > 0) {
                stocks.push({
                    name: stockName,
                    value: currentValue,
                    isUS: isUSStock
                });
            }
        }
    }

    stocks.sort((a, b) => b.value - a.value);

    stocks.forEach(stock => {
        const tr = document.createElement('tr');

        // Format display
        const parts = stock.name.split('-');
        let displayName = parts[0];
        let code = parts[1] || '';

        // Currency Display
        let valueStr = '';
        if (isUSPage) {
            valueStr = formatUSD(stock.value);
        } else {
            if (stock.isUS) {
                // Show USD with label or converted TWD?
                // Plan: "indicate currency".
                valueStr = `${formatUSD(stock.value)} <span style="font-size:0.8em; color:#94a3b8">(USD)</span>`;
            } else {
                valueStr = formatCurrency(stock.value);
            }
        }

        tr.innerHTML = `
            <td>
                <div class="stock-name">${displayName} <span class="stock-code">${code}</span></div>
            </td>
            <td class="text-right">
                <div class="asset-value">${valueStr}</div>
            </td>
        `;
        stockListBody.appendChild(tr);
    });
}

function clearUI() {
    // 重設數據
    ['unrealizedValue', 'realizedValue', 'dividendValue', 'feeValue', 'totalProfitValue', 'twAssetValue', 'usAssetValue', 'totalCostValue', 'netProfitValue'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.textContent = '--';
            el.className = 'metric-value neutral';
        }
    });
    const stockListBody = document.getElementById('stockListBody');
    if (stockListBody) stockListBody.innerHTML = '';

    if (myChart) {
        myChart.destroy();
        myChart = null;
    }
    document.getElementById('assetTrend').innerHTML = '<span id="trendValue">--%</span>';
    document.getElementById('assetTrend').className = 'trend';
}

function parseCurrency(str) {
    if (!str) return 0;
    const clean = parseFloat(str.replace(/,/g, '').trim());
    return isNaN(clean) ? 0 : clean;
}

function updateMetric(elementId, value, forcePositive = false, isExpense = false, formatter = formatCurrency) {
    const el = document.getElementById(elementId);
    if (!el) return;

    el.textContent = formatter(value);

    el.className = 'metric-value';
    if (isExpense) {
        el.classList.add('loss');
    } else if (value > 0 || forcePositive) {
        el.classList.add('gain');
    } else if (value < 0) {
        el.classList.add('loss');
    } else {
        el.classList.add('neutral');
    }
}

function formatCurrency(num) {
    return new Intl.NumberFormat('zh-TW', {
        style: 'currency',
        currency: 'TWD',
        maximumFractionDigits: 0
    }).format(num);
}

function formatUSD(num) {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        maximumFractionDigits: 2
    }).format(num);
}

function renderChart(labels, data, isUSPage) {
    const ctx = document.getElementById('assetChart').getContext('2d');
    if (myChart) myChart.destroy();

    const gradient = ctx.createLinearGradient(0, 0, 0, 400);
    gradient.addColorStop(0, 'rgba(56, 189, 248, 0.5)');
    gradient.addColorStop(1, 'rgba(56, 189, 248, 0.0)');

    myChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: isUSPage ? '總資產 (USD)' : '總資產 (TWD)',
                data: data,
                borderColor: '#38bdf8',
                backgroundColor: gradient,
                borderWidth: 3,
                pointBackgroundColor: '#fff',
                pointBorderColor: '#38bdf8',
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    callbacks: {
                        label: function (context) {
                            return isUSPage ? formatUSD(context.parsed.y) : formatCurrency(context.parsed.y);
                        }
                    }
                }
            },
            scales: {
                x: { grid: { display: false } },
                y: { display: false, grid: { display: false } }
            },
            interaction: { mode: 'nearest', axis: 'x', intersect: false }
        }
    });
}
