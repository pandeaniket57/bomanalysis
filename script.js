let bomData = [];
const HIDDEN_KEYS = new Set([]);

function normalizeKey(k){ return String(k||"").trim(); }

function parseWorkbook(workbook){
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval: ""});
    if(rows.length === 0) return [];
    let headerRowIndex = 0;
    const limit = Math.min(rows.length, 10);
    for(let i=0;i<limit;i++){
        const rowText = rows[i].map(c=>String(c||"").toLowerCase()).join(' ');
        if((rowText.includes('part') || rowText.includes('mpn') || rowText.includes('description'))){
            headerRowIndex = i;
            break;
        }
    }
    const headers = rows[headerRowIndex].map((h,idx)=> normalizeKey(h) || `Column${idx+1}`);
    const data = rows.slice(headerRowIndex+1).map(r=>{
        const obj = {};
        headers.forEach((h,i)=> obj[h] = r[i] != null ? r[i] : "");
        return obj;
    });
    return data;
}

function loadFromFile(file){
    const reader = new FileReader();
    reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type:'array'});
        bomData = parseWorkbook(workbook);
        renderTable(bomData);
        analyzeAllDimensions();
        document.getElementById('exportPdfBtn').style.display = 'inline-flex';
    };
    reader.readAsArrayBuffer(file);
}

function renderTable(data){
    const wrap = document.getElementById('table-wrap');
    wrap.innerHTML = '';
    if(!data || data.length === 0){
        wrap.innerHTML = '<div class="empty">No data loaded</div>';
        return;
    }
    const headers = Object.keys(data[0]);
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    headers.forEach(h=>{ const th = document.createElement('th'); th.textContent = h; tr.appendChild(th); });
    thead.appendChild(tr);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    data.forEach(row=>{
        const tr = document.createElement('tr');
        headers.forEach(h=>{
            const td = document.createElement('td');
            td.textContent = row[h] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
}

// 20-Dimension Analysis Engine
function analyzeAllDimensions(){
    if(!bomData.length) return;
    
    const dimensions = {
        cost: analyzeCost(),  // Now includes both category cost AND ABC
        supplier: analyzeSupplier(),
        manufacturer: analyzeManufacturer(),
        category: analyzeCategory(),
        quantity: analyzeQuantity(),
        lifecycle: analyzeLifecycle(),
        leadtime: analyzeLeadTime(),
        compliance: analyzeCompliance(),
        risk: analyzeRisk(),
        abc: analyzeABC(),
        criticality: analyzeCriticality(),
        alternate: analyzeAlternate(),
        pcb: analyzePCB(),
        reliability: analyzeReliability(),
        emi: analyzeEMI(),
        power: analyzePower(),
        country: analyzeCountry(),
        sustainability: analyzeSustainability(),
        temperature: analyzeTemperature()
    };
    
    window.dimensionsData = dimensions;
    renderChartsForDimensions(dimensions);
    return dimensions;
}

// ENHANCED Cost Analysis with Category-first + ABC Pareto
function analyzeCost(){
    const costCol = findColumn(['extended cost', 'total cost', 'cost', 'price', 'unit price']);
    const qtyCol = findColumn(['quantity', 'qty', 'count']);
    const categoryCol = findColumn(['category', 'type', 'component type', 'part type', 'class']);
    
    if(!costCol && !qtyCol) return { available: false, message: 'No cost or quantity column found' };
    
    // Calculate extended cost for each component
    let items = [];
    bomData.forEach((r, idx) => {
        let cost = 0;
        const unitPrice = parseFloat(r[costCol]) || 0;
        const quantity = parseInt(r[qtyCol]) || 1;
        
        if(costCol && r[costCol]) {
            cost = unitPrice;
        } else if(qtyCol && unitPrice) {
            cost = unitPrice * quantity;
        }
        
        if(cost > 0) {
            items.push({
                id: idx,
                name: r['Part Number'] || r['MPN'] || r['Description'] || `Item ${idx+1}`,
                description: r['Description'] || '',
                category: r[categoryCol] || 'Uncategorized',
                unitPrice: unitPrice,
                quantity: quantity,
                extendedCost: cost
            });
        }
    });
    
    if(items.length === 0) return { available: false, message: 'No cost data found' };
    
    // Sort by extended cost descending
    items.sort((a,b) => b.extendedCost - a.extendedCost);
    
    // Calculate totals
    const totalCost = items.reduce((sum, i) => sum + i.extendedCost, 0);
    const avgCost = totalCost / items.length;
    const top10Expensive = items.slice(0, 10);
    
    // ========== 1. COST BY CATEGORY ==========
    const categoryCost = {};
    const categoryCount = {};
    items.forEach(item => {
        const cat = item.category;
        if(!categoryCost[cat]) {
            categoryCost[cat] = 0;
            categoryCount[cat] = 0;
        }
        categoryCost[cat] += item.extendedCost;
        categoryCount[cat] += 1;
    });
    
    // Convert to array and sort by cost descending
    const categoryCostArray = Object.entries(categoryCost)
        .map(([name, cost]) => ({ 
            name, 
            cost, 
            count: categoryCount[name],
            percentage: (cost / totalCost) * 100
        }))
        .sort((a,b) => b.cost - a.cost);
    
    // ========== 2. ABC PARETO CLASSIFICATION ==========
    // Sort items by cost descending for ABC analysis
    const sortedByCost = [...items].sort((a,b) => b.extendedCost - a.extendedCost);
    
    let cumulativeCost = 0;
    const abcResults = {
        A: [],  // Top 70-80% of total cost
        B: [],  // Next 15-20% of total cost
        C: []   // Remaining 5-10% of total cost
    };
    
    sortedByCost.forEach(item => {
        cumulativeCost += item.extendedCost;
        const cumulativePercentage = (cumulativeCost / totalCost) * 100;
        
        if(cumulativePercentage <= 70) {
            abcResults.A.push(item);
        } else if(cumulativePercentage <= 90) {
            abcResults.B.push(item);
        } else {
            abcResults.C.push(item);
        }
    });
    
    // Calculate ABC percentages
    const totalACost = abcResults.A.reduce((sum, i) => sum + i.extendedCost, 0);
    const totalBCost = abcResults.B.reduce((sum, i) => sum + i.extendedCost, 0);
    const totalCCost = abcResults.C.reduce((sum, i) => sum + i.extendedCost, 0);
    
    const abcSummary = {
        A: { 
            count: abcResults.A.length, 
            cost: totalACost, 
            percentage: (totalACost / totalCost) * 100,
            items: abcResults.A.slice(0, 5)  // Top 5 A items
        },
        B: { 
            count: abcResults.B.length, 
            cost: totalBCost, 
            percentage: (totalBCost / totalCost) * 100,
            items: abcResults.B.slice(0, 5)
        },
        C: { 
            count: abcResults.C.length, 
            cost: totalCCost, 
            percentage: (totalCCost / totalCost) * 100,
            items: abcResults.C.slice(0, 5)
        }
    };
    
    return { 
        available: true, 
        totalCost,
        avgCost,
        itemCount: items.length,
        top10Expensive,
        categoryCost: categoryCostArray,
        abc: abcSummary,
        allItems: items,
        costColumn: costCol,
        categoryColumn: categoryCol
    };
}

function analyzeSupplier(){
    const supplierCol = findColumn(['supplier', 'vendor', 'manufacturer', 'mfr']);
    if(!supplierCol) return { available: false };
    const counts = {};
    bomData.forEach(r => { const s = String(r[supplierCol]||'Unknown'); counts[s] = (counts[s]||0)+1; });
    const unique = Object.keys(counts).length;
    const top5 = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,5);
    const singleSource = Object.values(counts).filter(c => c === 1).length;
    return { available: true, unique, top5, singleSource, column: supplierCol };
}

function analyzeManufacturer(){
    const mfrCol = findColumn(['manufacturer', 'mfr', 'brand', 'make']);
    if(!mfrCol) return { available: false };
    const counts = {};
    bomData.forEach(r => { const m = String(r[mfrCol]||'Unknown'); counts[m] = (counts[m]||0)+1; });
    return { available: true, unique: Object.keys(counts).length, top5: Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,5), column: mfrCol };
}

function analyzeCategory(){
    const catCol = findColumn(['category', 'type', 'component type', 'part type', 'class']);
    if(!catCol) return { available: false };
    const counts = {};
    bomData.forEach(r => { const c = String(r[catCol]||'Uncategorized'); counts[c] = (counts[c]||0)+1; });
    return { available: true, distribution: counts, column: catCol };
}

function analyzeQuantity(){
    const qtyCol = findColumn(['quantity', 'qty', 'count', 'pcs']);
    if(!qtyCol) return { available: false };
    const qtys = bomData.map(r => parseInt(r[qtyCol]) || 0);
    const total = qtys.reduce((a,b)=>a+b,0);
    const avg = total / qtys.length;
    const max = Math.max(...qtys);
    return { available: true, total, avg, max, column: qtyCol };
}

function analyzeLifecycle(){
    const lcCol = findColumn(['lifecycle', 'status', 'life cycle', 'obsolete', 'eol', 'nrnd']);
    if(!lcCol) return { available: false };
    const counts = { Active: 0, NRND: 0, EOL: 0, Obsolete: 0, Unknown: 0 };
    bomData.forEach(r => {
        const val = String(r[lcCol]||'').toLowerCase();
        if(val.includes('active')) counts.Active++;
        else if(val.includes('nrnd')) counts.NRND++;
        else if(val.includes('eol') || val.includes('end of life')) counts.EOL++;
        else if(val.includes('obsolete')) counts.Obsolete++;
        else counts.Unknown++;
    });
    return { available: true, counts, column: lcCol };
}

function analyzeLeadTime(){
    const ltCol = findColumn(['lead time', 'leadtime', 'lt', 'delivery', 'procurement']);
    if(!ltCol) return { available: false };
    const times = bomData.map(r => parseFloat(r[ltCol]) || 0).filter(t => t > 0);
    if(times.length === 0) return { available: false };
    const avg = times.reduce((a,b)=>a+b,0) / times.length;
    const max = Math.max(...times);
    const critical = times.filter(t => t > 20).length;
    const distribution = { '0-2 weeks': 0, '3-6 weeks': 0, '7-12 weeks': 0, '13+ weeks': 0 };
    times.forEach(t => {
        if(t <= 2) distribution['0-2 weeks']++;
        else if(t <= 6) distribution['3-6 weeks']++;
        else if(t <= 12) distribution['7-12 weeks']++;
        else distribution['13+ weeks']++;
    });
    return { available: true, avg, max, critical, count: times.length, distribution, column: ltCol };
}

function analyzeCompliance(){
    const rohsCol = findColumn(['rohs']);
    const reachCol = findColumn(['reach']);
    const compliance = { RoHS: { Yes: 0, No: 0, Unknown: 0 }, REACH: { Yes: 0, No: 0, Unknown: 0 } };
    bomData.forEach(r => {
        const rohs = String(r[rohsCol] || '').toLowerCase();
        if(rohs.includes('yes') || rohs === '✔') compliance.RoHS.Yes++;
        else if(rohs.includes('no')) compliance.RoHS.No++;
        else compliance.RoHS.Unknown++;
        
        const reach = String(r[reachCol] || '').toLowerCase();
        if(reach.includes('yes') || reach === '✔') compliance.REACH.Yes++;
        else if(reach.includes('no')) compliance.REACH.No++;
        else compliance.REACH.Unknown++;
    });
    return { available: !!(rohsCol || reachCol), rohs: compliance.RoHS, reach: compliance.REACH, rohsCol, reachCol };
}

function analyzeRisk(){
    const riskCol = findColumn(['risk', 'single source', 'critical', 'remark']);
    if(!riskCol) return { available: false, risks: [] };
    const risks = [];
    const riskCounts = { 'Single Source': 0, 'EOL Risk': 0, 'Long Lead Time': 0, 'Geopolitical': 0, 'Low Stock': 0 };
    bomData.forEach((r,idx) => {
        const val = String(r[riskCol]||'').toLowerCase();
        if(val.includes('single') || val.includes('sole')) { risks.push({ index: idx, part: r['Part Number'] || r['MPN'] || `Item ${idx}`, type: 'Single Source' }); riskCounts['Single Source']++; }
        if(val.includes('eol') || val.includes('obsolete')) { risks.push({ index: idx, part: r['Part Number'] || r['MPN'] || `Item ${idx}`, type: 'EOL Risk' }); riskCounts['EOL Risk']++; }
        if(val.includes('long lead')) { risks.push({ index: idx, part: r['Part Number'] || r['MPN'] || `Item ${idx}`, type: 'Long Lead Time' }); riskCounts['Long Lead Time']++; }
    });
    return { available: true, risks, count: risks.length, riskCounts, column: riskCol };
}

function analyzeABC(){
    // Use the cost analysis data for ABC
    const costAnalysis = analyzeCost();
    if(!costAnalysis.available) return { available: false };
    return costAnalysis.abc;
}

function analyzeCriticality(){
    const critCol = findColumn(['criticality', 'critical', 'importance', 'safety']);
    if(!critCol) return { available: false };
    const counts = { 'Safety-Critical': 0, 'Performance-Critical': 0, 'Standard': 0, 'Optional': 0 };
    bomData.forEach(r => {
        const val = String(r[critCol]||'').toLowerCase();
        if(val.includes('safety')) counts['Safety-Critical']++;
        else if(val.includes('performance')) counts['Performance-Critical']++;
        else if(val.includes('optional')) counts['Optional']++;
        else counts['Standard']++;
    });
    return { available: true, counts, column: critCol };
}

function analyzeTemperature(){
    const tempCol = findColumn(['temperature', 'temp range', 'operating temperature', 'temperature range']);
    if(!tempCol) return { available: false };
    const ranges = { '-55°C to 125°C': 0, '-40°C to 85°C': 0, '-40°C to 125°C': 0, '-20°C to 70°C': 0, 'Other': 0 };
    bomData.forEach(r => {
        const val = String(r[tempCol]||'').toLowerCase();
        if(val.includes('-55') && val.includes('125')) ranges['-55°C to 125°C']++;
        else if(val.includes('-40') && val.includes('85')) ranges['-40°C to 85°C']++;
        else if(val.includes('-40') && val.includes('125')) ranges['-40°C to 125°C']++;
        else if(val.includes('-20') && val.includes('70')) ranges['-20°C to 70°C']++;
        else ranges['Other']++;
    });
    return { available: true, ranges, column: tempCol };
}

function analyzeAlternate(){
    const altCol = findColumn(['alternate', 'alternative', 'cross reference', 'substitute']);
    if(!altCol) return { available: false };
    let hasAlternate = 0;
    bomData.forEach(r => {
        const val = String(r[altCol]||'');
        if(val && val.length > 0 && val !== 'N/A' && val !== 'None') hasAlternate++;
    });
    return { available: true, hasAlternate, total: bomData.length, pct: (hasAlternate/bomData.length*100).toFixed(1), column: altCol };
}

function analyzePCB(){
    const pkgCol = findColumn(['package', 'case', 'footprint', 'pkg']);
    if(!pkgCol) return { available: false };
    let smt = 0, tht = 0;
    bomData.forEach(r => {
        const val = String(r[pkgCol]||'').toLowerCase();
        if(val.includes('smt') || val.includes('smd')) smt++;
        else if(val.includes('tht') || val.includes('through hole')) tht++;
    });
    return { available: true, smt, tht, other: bomData.length - smt - tht, column: pkgCol };
}

function analyzeReliability(){
    const relCol = findColumn(['mtbf', 'derating', 'temperature', 'reliability']);
    if(!relCol) return { available: false };
    const values = bomData.map(r => parseFloat(r[relCol]) || 0).filter(v => v > 0);
    return { available: values.length > 0, avg: values.reduce((a,b)=>a+b,0)/values.length || 0, count: values.length, column: relCol };
}

function analyzeEMI(){
    const emiCol = findColumn(['emi', 'emc', 'emi/emc', 'filter', 'shielding', 'ferrite']);
    if(!emiCol) return { available: false };
    const ratings = { 'Filtered': 0, 'Shielded': 0, 'Standard': 0, 'RF Module': 0 };
    bomData.forEach(r => {
        const val = String(r[emiCol]||'').toLowerCase();
        if(val.includes('filtered')) ratings['Filtered']++;
        else if(val.includes('shielded')) ratings['Shielded']++;
        else if(val.includes('rf') || val.includes('module')) ratings['RF Module']++;
        else ratings['Standard']++;
    });
    return { available: true, ratings, column: emiCol };
}

function analyzePower(){
    const powerCol = findColumn(['power', 'regulator', 'mosfet', 'current', 'voltage']);
    if(!powerCol) return { available: false };
    let powerParts = 0;
    bomData.forEach(r => {
        const val = String(r[powerCol]||'').toLowerCase();
        if(val.includes('regulator') || val.includes('mosfet') || val.includes('power') || val.includes('current')) powerParts++;
    });
    return { available: true, powerParts, pct: (powerParts/bomData.length*100).toFixed(1), column: powerCol };
}

function analyzeCountry(){
    const countryCol = findColumn(['country', 'origin', 'source country', 'vendor country', 'country of origin']);
    if(!countryCol) return { available: false };
    const counts = {};
    bomData.forEach(r => { const c = String(r[countryCol]||'Unknown'); counts[c] = (counts[c]||0)+1; });
    return { available: true, distribution: counts, column: countryCol };
}

function analyzeSustainability(){
    const susCol = findColumn(['hazardous', 'green', 'recyclable', 'environment']);
    if(!susCol) return { available: false };
    let green = 0, hazardous = 0;
    bomData.forEach(r => {
        const val = String(r[susCol]||'').toLowerCase();
        if(val.includes('green') || val.includes('rohs') || val.includes('lead-free')) green++;
        if(val.includes('hazardous') || val.includes('lead') || val.includes('hg')) hazardous++;
    });
    return { available: true, green, hazardous, greenPct: (green/bomData.length*100).toFixed(1), column: susCol };
}

function findColumn(possibleNames){
    if(!bomData.length) return null;
    const headers = Object.keys(bomData[0]).map(h => h.toLowerCase());
    for(const name of possibleNames){
        const found = headers.find(h => h.includes(name.toLowerCase()));
        if(found) return Object.keys(bomData[0])[headers.indexOf(found)];
    }
    return null;
}

// Chart Rendering
let chartInstances = {};

function renderChartsForDimensions(dimensions){
    const chartsDiv = document.getElementById('charts');
    chartsDiv.innerHTML = '';
    chartInstances = {};
    
    const dimList = [
        { key: 'cost', name: '💰 Cost Analysis (Category + ABC Pareto)', data: dimensions.cost },
        { key: 'supplier', name: '🏭 Supplier Distribution', data: dimensions.supplier },
        { key: 'manufacturer', name: '🔧 Top Manufacturers', data: dimensions.manufacturer },
        { key: 'category', name: '📦 Component Categories', data: dimensions.category },
        { key: 'lifecycle', name: '🔄 Lifecycle Status', data: dimensions.lifecycle },
        { key: 'leadtime', name: '⏱️ Lead Time Distribution', data: dimensions.leadtime },
        { key: 'compliance', name: '✅ Compliance (RoHS/REACH)', data: dimensions.compliance },
        { key: 'risk', name: '⚠️ Risk Analysis', data: dimensions.risk },
        { key: 'criticality', name: '🎯 Criticality Analysis', data: dimensions.criticality },
        { key: 'temperature', name: '🌡️ Temperature Range', data: dimensions.temperature },
        { key: 'emi', name: '📡 EMI/EMC Rating', data: dimensions.emi },
        { key: 'country', name: '🌍 Vendor Country', data: dimensions.country }
    ];
    
    for(const dim of dimList){
        if(dim.data && dim.data.available){
            const card = document.createElement('div');
            card.className = 'chart-card';
            const title = document.createElement('h3');
            title.textContent = dim.name;
            card.appendChild(title);
            const canvas = document.createElement('canvas');
            canvas.id = `chart_${dim.key}`;
            card.appendChild(canvas);
            chartsDiv.appendChild(card);
            
            if(dim.key === 'cost' && dim.data.categoryCost){
                // Show Category Cost Distribution (Pie Chart)
                const categories = dim.data.categoryCost.slice(0, 8);
                createPieChart(canvas.id, 'Cost by Category (USD)', 
                    categories.map(c => `${c.name} ($${c.cost.toFixed(0)})`), 
                    categories.map(c => c.cost));
            } else if(dim.key === 'supplier' && dim.data.top5){
                createPieChart(canvas.id, 'Supplier Share', dim.data.top5.map(s=>s[0]), dim.data.top5.map(s=>s[1]));
            } else if(dim.key === 'manufacturer' && dim.data.top5){
                createPieChart(canvas.id, 'Manufacturer Share', dim.data.top5.map(m=>m[0]), dim.data.top5.map(m=>m[1]));
            } else if(dim.key === 'category' && dim.data.distribution){
                const entries = Object.entries(dim.data.distribution).slice(0,8);
                createPieChart(canvas.id, 'Category Distribution', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'lifecycle' && dim.data.counts){
                const entries = Object.entries(dim.data.counts).filter(e=>e[1]>0);
                createPieChart(canvas.id, 'Lifecycle Status', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'leadtime' && dim.data.distribution){
                const entries = Object.entries(dim.data.distribution);
                createBarChart(canvas.id, 'Lead Time Distribution (weeks)', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'compliance' && dim.data.rohs){
                createBarChart(canvas.id, 'RoHS Compliance', ['Yes', 'No', 'Unknown'], [dim.data.rohs.Yes, dim.data.rohs.No, dim.data.rohs.Unknown]);
            } else if(dim.key === 'risk' && dim.data.riskCounts){
                const entries = Object.entries(dim.data.riskCounts).filter(e=>e[1]>0);
                createBarChart(canvas.id, 'Risk Categories', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'criticality' && dim.data.counts){
                const entries = Object.entries(dim.data.counts).filter(e=>e[1]>0);
                createPieChart(canvas.id, 'Criticality Distribution', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'temperature' && dim.data.ranges){
                const entries = Object.entries(dim.data.ranges).filter(e=>e[1]>0);
                createBarChart(canvas.id, 'Temperature Range Distribution', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'emi' && dim.data.ratings){
                const entries = Object.entries(dim.data.ratings);
                createPieChart(canvas.id, 'EMI/EMC Ratings', entries.map(e=>e[0]), entries.map(e=>e[1]));
            } else if(dim.key === 'country' && dim.data.distribution){
                const entries = Object.entries(dim.data.distribution).slice(0,6);
                createPieChart(canvas.id, 'Country Distribution', entries.map(e=>e[0]), entries.map(e=>e[1]));
            }
        }
    }
    
    // Add ABC Pareto chart as a separate card if cost data is available
    if(window.dimensionsData?.cost?.abc){
        const abcCard = document.createElement('div');
        abcCard.className = 'chart-card';
        const abcTitle = document.createElement('h3');
        abcTitle.textContent = '📊 ABC Pareto Classification (80/20 Rule)';
        abcCard.appendChild(abcTitle);
        const abcCanvas = document.createElement('canvas');
        abcCanvas.id = 'chart_abc_pareto';
        abcCard.appendChild(abcCanvas);
        chartsDiv.appendChild(abcCard);
        
        const abc = window.dimensionsData.cost.abc;
        createPieChart('chart_abc_pareto', 'ABC Classification by Cost', 
            [`A (70% Cost) - ${abc.A.count} parts`, `B (20% Cost) - ${abc.B.count} parts`, `C (10% Cost) - ${abc.C.count} parts`],
            [abc.A.cost, abc.B.cost, abc.C.cost]);
    }
}

function generateColors(n){
    const palette = ['#19196e', '#2e2e9a', '#cc3300', '#ff6644', '#5a5ab5', '#b4b4dc', '#1a6b3c', '#b58a00', '#00a86b', '#ff6b35'];
    return Array.from({length: n}, (_,i) => palette[i % palette.length]);
}

function createPieChart(canvasId, title, labels, data){
    const ctx = document.getElementById(canvasId);
    if(!ctx) return;
    if(chartInstances[canvasId]) chartInstances[canvasId].destroy();
    chartInstances[canvasId] = new Chart(ctx, {
        type: 'pie',
        data: { labels, datasets: [{ data, backgroundColor: generateColors(labels.length), borderWidth: 0 }] },
        options: { 
            responsive: true, 
            maintainAspectRatio: true, 
            plugins: { 
                legend: { position: 'right', labels: { font: { size: 10 } } }, 
                title: { display: true, text: title, font: { size: 12, weight: 'bold' } }
            } 
        }
    });
}

function createBarChart(canvasId, title, labels, data){
    const ctx = document.getElementById(canvasId);
    if(!ctx) return;
    if(chartInstances[canvasId]) chartInstances[canvasId].destroy();
    chartInstances[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets: [{ label: title, data, backgroundColor: '#19196e', borderRadius: 6 }] },
        options: { responsive: true, maintainAspectRatio: true, plugins: { legend: { display: false }, title: { display: true, text: title } } }
    });
}

// Tab navigation
document.querySelectorAll('.dim-tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.dim-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const dim = tab.getAttribute('data-dim');
        if(dim === 'all' && window.dimensionsData){
            renderChartsForDimensions(window.dimensionsData);
        } else if(window.dimensionsData && window.dimensionsData[dim]){
            const chartsDiv = document.getElementById('charts');
            chartsDiv.innerHTML = '';
            const dimData = window.dimensionsData[dim];
            if(dimData.available){
                const card = document.createElement('div');
                card.className = 'chart-card';
                const title = document.createElement('h3');
                title.textContent = tab.textContent;
                card.appendChild(title);
                const canvas = document.createElement('canvas');
                canvas.id = `chart_single`;
                card.appendChild(canvas);
                chartsDiv.appendChild(card);
                
                if(dim === 'leadtime' && dimData.distribution){
                    createBarChart('chart_single', 'Lead Time (weeks)', Object.keys(dimData.distribution), Object.values(dimData.distribution));
                } else if(dim === 'compliance' && dimData.rohs){
                    createBarChart('chart_single', 'RoHS Compliance', ['Yes', 'No', 'Unknown'], [dimData.rohs.Yes, dimData.rohs.No, dimData.rohs.Unknown]);
                } else if(dim === 'risk' && dimData.riskCounts){
                    createBarChart('chart_single', 'Risk Categories', Object.keys(dimData.riskCounts), Object.values(dimData.riskCounts));
                } else if(dim === 'temperature' && dimData.ranges){
                    createBarChart('chart_single', 'Temperature Range', Object.keys(dimData.ranges), Object.values(dimData.ranges));
                } else if(dim === 'criticality' && dimData.counts){
                    createPieChart('chart_single', 'Criticality', Object.keys(dimData.counts), Object.values(dimData.counts));
                } else if(dim === 'emi' && dimData.ratings){
                    createPieChart('chart_single', 'EMI/EMC', Object.keys(dimData.ratings), Object.values(dimData.ratings));
                } else if(dim === 'country' && dimData.distribution){
                    createPieChart('chart_single', 'Country', Object.keys(dimData.distribution), Object.values(dimData.distribution));
                } else if(dim === 'lifecycle' && dimData.counts){
                    createPieChart('chart_single', 'Lifecycle', Object.keys(dimData.counts), Object.values(dimData.counts));
                } else if(dim === 'cost' && dimData.categoryCost){
                    createPieChart('chart_single', 'Cost by Category', 
                        dimData.categoryCost.slice(0,6).map(c => `${c.name} ($${c.cost.toFixed(0)})`), 
                        dimData.categoryCost.slice(0,6).map(c => c.cost));
                }
            } else {
                chartsDiv.innerHTML = '<div class="empty">No data available for this dimension</div>';
            }
        }
    });
});

// Search & Drag Drop
document.addEventListener('input', e => { if(e.target.id === 'searchBox'){ const q = e.target.value.toLowerCase(); const filtered = bomData.filter(r => JSON.stringify(r).toLowerCase().includes(q)); renderTable(filtered); } });

const fileInput = document.getElementById('excelFile');
const dropZone = document.getElementById('drop-zone');
const loadBtn = document.getElementById('loadBtn');

if(fileInput) fileInput.addEventListener('change', ()=>{ if(fileInput.files[0]) loadFromFile(fileInput.files[0]); });
if(loadBtn) loadBtn.addEventListener('click', ()=>{ if(fileInput.files[0]) loadFromFile(fileInput.files[0]); else alert('Select a file first.'); });
if(dropZone){
    ['dragenter','dragover'].forEach(ev=> dropZone.addEventListener(ev, e=>{ e.preventDefault(); dropZone.classList.add('drag'); }));
    ['dragleave','drop'].forEach(ev=> dropZone.addEventListener(ev, e=>{ e.preventDefault(); dropZone.classList.remove('drag'); }));
    dropZone.addEventListener('drop', e=>{ const f = e.dataTransfer.files[0]; if(f) { fileInput.files = e.dataTransfer.files; loadFromFile(f); } });
}