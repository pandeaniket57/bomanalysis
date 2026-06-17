// pdf-export.js — Simple & Reliable BOM Report Generator
(function(){
    if(typeof window.jspdf === 'undefined'){
        console.error('jsPDF not loaded!');
        alert('jsPDF library not loaded. Please check your internet connection.');
        return;
    }
    
    const { jsPDF } = window.jspdf;

    // Default color palette
    const DEFAULT_PALETTE = {
        navy: [25, 25, 110],
        steel: [180, 180, 220],
        light: [245, 242, 230],
        white: [255, 255, 255],
        text: [40, 40, 40],
        muted: [100, 100, 100],
        accent: [204, 51, 0],
        statusActive: [26, 107, 60],
        statusWarn: [181, 138, 0],
        statusCrit: [204, 51, 0]
    };
    
    function loadPalette() {
        try {
            const saved = sessionStorage.getItem('bomPalette');
            if(saved) {
                const parsed = JSON.parse(saved);
                if(parsed && Object.keys(parsed).length > 0) return parsed;
            }
        } catch(e) {}
        return JSON.parse(JSON.stringify(DEFAULT_PALETTE));
    }
    
    let logoDataUrl = null;
    
    function formatDate(d) {
        if(!d) return new Date().toLocaleDateString('en-GB');
        const parts = d.split('-');
        if(parts.length === 3) {
            return new Date(parts[0], parts[1]-1, parts[2]).toLocaleDateString('en-GB', 
                { day: '2-digit', month: 'long', year: 'numeric' });
        }
        return d;
    }
    
    function setProgress(pct, msg) {
        const fill = document.getElementById('progressFill');
        const text = document.getElementById('progressText');
        if(fill) fill.style.width = pct + '%';
        if(text) text.textContent = msg;
        console.log(`${pct}%: ${msg}`);
    }
    
    function closeModal() {
        const modal = document.getElementById('pdf-modal');
        if(modal) modal.style.display = 'none';
    }
    
    async function generatePDF() {
        // Check if data exists
        if(typeof bomData === 'undefined' || !bomData || bomData.length === 0) {
            alert('No BOM data loaded. Please upload an Excel file first.');
            return;
        }
        
        console.log('Starting PDF generation with', bomData.length, 'items');
        
        const C = loadPalette();
        const company = document.getElementById('companyName')?.value.trim() || 'Power Electronical';
        const reportName = document.getElementById('reportName')?.value.trim() || 'BOM Analysis Report';
        const dateStr = document.getElementById('reportDate')?.value || new Date().toISOString().split('T')[0];
        
        // Get user custom text
        const customExecutiveSummary = document.getElementById('customExecutiveSummary')?.value.trim() || '';
        const customKeyFindings = document.getElementById('customKeyFindings')?.value.trim() || '';
        const customRecommendations = document.getElementById('customRecommendations')?.value.trim() || '';
        const customConclusion = document.getElementById('customConclusion')?.value.trim() || '';
        
        // Section checkboxes
        const sections = {
            cover: document.getElementById('incCover')?.checked !== false,
            summary: document.getElementById('incSummary')?.checked !== false,
            manufacturer: document.getElementById('incManufacturer')?.checked !== false,
            supplier: document.getElementById('incSupplier')?.checked !== false,
            cost: document.getElementById('incCost')?.checked !== false,
            category: document.getElementById('incCategory')?.checked !== false,
            lifecycle: document.getElementById('incLifecycle')?.checked !== false,
            compliance: document.getElementById('incCompliance')?.checked !== false,
            risk: document.getElementById('incRisk')?.checked !== false,
            table: document.getElementById('incTable')?.checked !== false,
            conclusion: document.getElementById('incConclusion')?.checked !== false
        };
        
        const prog = document.getElementById('pdfProgress');
        if(prog) prog.style.display = 'block';
        setProgress(5, 'Initializing...');
        
        // Create PDF in LANDSCAPE orientation for better table display
        const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
        const PW = doc.internal.pageSize.getWidth();
        const PH = doc.internal.pageSize.getHeight();
        const MARGIN = 48;
        const CW = PW - (MARGIN * 2);
        
        let pageNum = 1;
        
        function addPage() {
            doc.addPage();
            pageNum++;
            drawHeader();
            drawFooter();
        }
        
        function drawHeader() {
            doc.setFillColor(...C.navy);
            doc.rect(0, 0, PW, 38, 'F');
            doc.setFillColor(...C.accent);
            doc.rect(0, 38, PW, 3, 'F');
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(10);
            doc.setTextColor(...C.steel);
            doc.text(company.toUpperCase(), MARGIN, 24);
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(10);
            doc.setTextColor(...C.steel);
            doc.text(reportName, PW - MARGIN, 24, { align: 'right' });
        }
        
        function drawFooter() {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(...C.muted);
            doc.text(`${company} — BOM Report`, MARGIN, PH - 20);
            doc.text(`Page ${pageNum}`, PW - MARGIN, PH - 20, { align: 'right' });
            doc.setDrawColor(...C.light);
            doc.setLineWidth(0.5);
            doc.line(MARGIN, PH - 32, PW - MARGIN, PH - 32);
        }
        
        let currentY = MARGIN;
        
        function checkPage(needed) {
            if(currentY + needed > PH - 60) {
                addPage();
                currentY = MARGIN + 10;
                return true;
            }
            return false;
        }
        
        function addText(text, size = 10, isBold = false, color = 'text', indent = 0) {
            if(!text) return;
            const lines = doc.splitTextToSize(text, CW - indent);
            checkPage(lines.length * 16);
            doc.setFont('helvetica', isBold ? 'bold' : 'normal');
            doc.setFontSize(size);
            doc.setTextColor(...C[color]);
            doc.text(lines, MARGIN + indent, currentY);
            currentY += (lines.length * 16) + 4;
        }
        
        function addSectionTitle(title) {
            checkPage(50);
            doc.setFillColor(...C.accent);
            doc.rect(MARGIN, currentY, 5, 28, 'F');
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(16);
            doc.setTextColor(...C.navy);
            doc.text(title, MARGIN + 16, currentY + 22);
            doc.setDrawColor(...C.light);
            doc.setLineWidth(1);
            doc.line(MARGIN, currentY + 32, MARGIN + CW, currentY + 32);
            currentY += 46;
        }
        
        function addStatRow(label, value, unit = '', color = 'text') {
            checkPage(24);
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(11);
            doc.setTextColor(...C[color]);
            doc.text(label, MARGIN + 10, currentY);
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(...C[color]);
            doc.text(`${value}${unit}`, PW - MARGIN - 10, currentY, { align: 'right' });
            currentY += 24;
        }
        
        // ========== COVER PAGE ==========
        if(sections.cover) {
            setProgress(10, 'Creating cover page...');
            doc.deletePage(1);
            doc.addPage();
            
            const topH = PH * 0.55;
            doc.setFillColor(...C.navy);
            doc.rect(0, 0, PW, topH, 'F');
            doc.setFillColor(...C.light);
            doc.rect(0, topH, PW, PH * 0.45, 'F');
            doc.setFillColor(...C.accent);
            doc.rect(0, topH - 5, PW, 10, 'F');
            doc.setFillColor(...C.accent);
            doc.rect(0, 0, 6, topH, 'F');
            
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(12);
            doc.setTextColor(...C.steel);
            doc.text(company.toUpperCase(), MARGIN, PH * 0.22);
            
            doc.setFontSize(36);
            doc.setTextColor(...C.white);
            const titleLines = doc.splitTextToSize(reportName, CW - 20);
            doc.text(titleLines, MARGIN, PH * 0.30);
            
            doc.setFontSize(12);
            doc.setTextColor(...C.steel);
            doc.text('Bill of Materials Analysis Report', MARGIN, PH * 0.30 + (titleLines.length * 42));
            
            const cardW = (CW - 30) / 3;
            const cardY = PH * 0.58;
            
            const cards = [
                { label: 'TOTAL PARTS', value: bomData.length },
                { label: 'REPORT DATE', value: formatDate(dateStr).split(',')[0] }
            ];
            cards.forEach((card, i) => {
                const cardX = MARGIN + i * (cardW + 15);
                doc.setFillColor(...C.navy);
                doc.roundedRect(cardX, cardY, cardW, 70, 6, 6, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(24);
                doc.setTextColor(...C.white);
                doc.text(String(card.value), cardX + cardW/2, cardY + 38, { align: 'center' });
                doc.setFontSize(9);
                doc.setTextColor(...C.steel);
                doc.text(card.label, cardX + cardW/2, cardY + 58, { align: 'center' });
            });
            
            pageNum = 1;
        }
        
        // ========== EXECUTIVE SUMMARY ==========
        if(sections.summary && customExecutiveSummary) {
            setProgress(18, 'Adding Executive Summary...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Executive Summary');
            addText(customExecutiveSummary, 11, false, 'text');
            drawFooter();
        }
        
        // ========== KEY FINDINGS ==========
        if(sections.summary && customKeyFindings) {
            setProgress(25, 'Adding Key Findings...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Key Findings');
            const findings = customKeyFindings.split('\n');
            findings.forEach(f => {
                if(f.trim()) addText('• ' + f.trim(), 11, false, 'text');
            });
            drawFooter();
        } else if(sections.summary && !customKeyFindings && window.dimensionsData) {
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Key Findings');
            const riskCount = window.dimensionsData?.risk?.count || 0;
            const eolCount = window.dimensionsData?.lifecycle?.counts?.EOL || 0;
            const rohsCompliant = window.dimensionsData?.compliance?.rohs?.Yes || 0;
            addText(`• Total Components Analyzed: ${bomData.length} parts`, 11, false, 'text');
            addText(`• Risk Items Identified: ${riskCount} components`, 11, false, riskCount > 0 ? 'statusWarn' : 'text');
            addText(`• End of Life Components: ${eolCount} parts requiring action`, 11, false, eolCount > 0 ? 'statusCrit' : 'text');
            addText(`• RoHS Compliance: ${Math.round(rohsCompliant/bomData.length*100)}% of parts compliant`, 11, false, 'text');
            drawFooter();
        }
        
        // ========== RECOMMENDATIONS ==========
        if(sections.summary && customRecommendations) {
            setProgress(32, 'Adding Recommendations...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Recommendations');
            const recs = customRecommendations.split('\n');
            recs.forEach(r => {
                if(r.trim()) addText('• ' + r.trim(), 11, false, 'text');
            });
            drawFooter();
        }
        
        // ========== MANUFACTURER ANALYSIS ==========
        if(sections.manufacturer && window.dimensionsData?.manufacturer?.available) {
            setProgress(40, 'Manufacturer analysis...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Manufacturer Analysis');
            
            const mfr = window.dimensionsData.manufacturer;
            addStatRow('Total Manufacturers', mfr.unique || 0, '');
            
            if(mfr.top5 && mfr.top5.length) {
                currentY += 10;
                const headers = ['Manufacturer', 'Components'];
                const data = mfr.top5.map(m => [m[0], m[1]]);
                const colWidth = CW / headers.length;
                const rowH = 22;
                const headerH = 26;
                
                doc.setFillColor(...C.navy);
                doc.rect(MARGIN, currentY, CW, headerH, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.setTextColor(...C.white);
                let x = MARGIN;
                headers.forEach(h => {
                    doc.text(h, x + 6, currentY + 17);
                    x += colWidth;
                });
                currentY += headerH;
                
                data.forEach((row, idx) => {
                    const bg = (idx % 2 === 0) ? C.white : C.light;
                    doc.setFillColor(...bg);
                    doc.rect(MARGIN, currentY, CW, rowH, 'F');
                    doc.setDrawColor(...C.light);
                    doc.line(MARGIN, currentY + rowH, MARGIN + CW, currentY + rowH);
                    
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(8);
                    doc.setTextColor(...C.text);
                    x = MARGIN;
                    row.forEach(cell => {
                        doc.text(String(cell || '-'), x + 6, currentY + 15);
                        x += colWidth;
                    });
                    currentY += rowH;
                });
                currentY += 8;
            }
            drawFooter();
        }
        
        // ========== SUPPLIER ANALYSIS ==========
        if(sections.supplier && window.dimensionsData?.supplier?.available) {
            setProgress(50, 'Supplier analysis...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Supplier Analysis');
            
            const sup = window.dimensionsData.supplier;
            addStatRow('Total Suppliers', sup.unique || 0, '');
            addStatRow('Single-Source Components', sup.singleSource || 0, '', sup.singleSource > 0 ? 'statusWarn' : 'text');
            
            if(sup.top5 && sup.top5.length) {
                currentY += 10;
                const headers = ['Supplier', 'Components'];
                const data = sup.top5.map(s => [s[0], s[1]]);
                const colWidth = CW / headers.length;
                const rowH = 22;
                const headerH = 26;
                
                doc.setFillColor(...C.navy);
                doc.rect(MARGIN, currentY, CW, headerH, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.setTextColor(...C.white);
                let x = MARGIN;
                headers.forEach(h => {
                    doc.text(h, x + 6, currentY + 17);
                    x += colWidth;
                });
                currentY += headerH;
                
                data.forEach((row, idx) => {
                    const bg = (idx % 2 === 0) ? C.white : C.light;
                    doc.setFillColor(...bg);
                    doc.rect(MARGIN, currentY, CW, rowH, 'F');
                    doc.setDrawColor(...C.light);
                    doc.line(MARGIN, currentY + rowH, MARGIN + CW, currentY + rowH);
                    
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(8);
                    doc.setTextColor(...C.text);
                    x = MARGIN;
                    row.forEach(cell => {
                        doc.text(String(cell || '-'), x + 6, currentY + 15);
                        x += colWidth;
                    });
                    currentY += rowH;
                });
                currentY += 8;
            }
            drawFooter();
        }
        
        // ========== COST ANALYSIS ==========
        if(sections.cost && window.dimensionsData?.cost?.available) {
            setProgress(60, 'Cost analysis...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Cost Analysis');
            
            const cost = window.dimensionsData.cost;
            addStatRow('Total Material Cost', `$${(cost.totalCost || 0).toLocaleString()}`, ' USD', 'statusActive');
            addStatRow('Average Component Cost', `$${(cost.avgCost || 0).toFixed(2)}`, ' USD');
            
            if(cost.categoryCost && cost.categoryCost.length) {
                currentY += 10;
                const headers = ['Category', 'Cost (USD)', 'Share'];
                const data = cost.categoryCost.slice(0, 6).map(c => [c.name, `$${c.cost.toFixed(2)}`, `${c.percentage.toFixed(1)}%`]);
                const colWidth = CW / headers.length;
                const rowH = 22;
                const headerH = 26;
                
                doc.setFillColor(...C.navy);
                doc.rect(MARGIN, currentY, CW, headerH, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.setTextColor(...C.white);
                let x = MARGIN;
                headers.forEach(h => {
                    doc.text(h, x + 6, currentY + 17);
                    x += colWidth;
                });
                currentY += headerH;
                
                data.forEach((row, idx) => {
                    const bg = (idx % 2 === 0) ? C.white : C.light;
                    doc.setFillColor(...bg);
                    doc.rect(MARGIN, currentY, CW, rowH, 'F');
                    doc.setDrawColor(...C.light);
                    doc.line(MARGIN, currentY + rowH, MARGIN + CW, currentY + rowH);
                    
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(8);
                    doc.setTextColor(...C.text);
                    x = MARGIN;
                    row.forEach(cell => {
                        doc.text(String(cell || '-'), x + 6, currentY + 15);
                        x += colWidth;
                    });
                    currentY += rowH;
                });
                currentY += 8;
            }
            drawFooter();
        }
        
        // ========== LIFECYCLE ==========
        if(sections.lifecycle && window.dimensionsData?.lifecycle?.available) {
            setProgress(70, 'Lifecycle analysis...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Lifecycle Status');
            
            const lc = window.dimensionsData.lifecycle;
            addStatRow('Active Components', lc.counts?.Active || 0, '', 'statusActive');
            addStatRow('NRND Components', lc.counts?.NRND || 0, '', lc.counts?.NRND > 0 ? 'statusWarn' : 'text');
            addStatRow('EOL Components', lc.counts?.EOL || 0, '', lc.counts?.EOL > 0 ? 'statusCrit' : 'text');
            drawFooter();
        }
        
        // ========== COMPLIANCE ==========
        if(sections.compliance && window.dimensionsData?.compliance?.available) {
            setProgress(78, 'Compliance analysis...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Compliance Status');
            
            const comp = window.dimensionsData.compliance;
            const compliant = comp.rohs?.Yes || 0;
            const nonComp = comp.rohs?.No || 0;
            addStatRow('RoHS Compliant', compliant, ` (${Math.round(compliant/bomData.length*100)}%)`, 'statusActive');
            addStatRow('Non-Compliant', nonComp, ` (${Math.round(nonComp/bomData.length*100)}%)`, nonComp > 0 ? 'statusCrit' : 'text');
            drawFooter();
        }
        
        // ========== RISK ==========
        if(sections.risk && window.dimensionsData?.risk?.available) {
            setProgress(85, 'Risk analysis...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Risk Assessment');
            
            const risk = window.dimensionsData.risk;
            addStatRow('Total Risk Items', risk.count || 0, '', risk.count > 0 ? 'statusCrit' : 'statusActive');
            
            if(risk.riskCounts && Object.keys(risk.riskCounts).length > 0) {
                currentY += 10;
                const headers = ['Risk Category', 'Count'];
                const data = Object.entries(risk.riskCounts).map(([cat, cnt]) => [cat, cnt]);
                const colWidth = CW / headers.length;
                const rowH = 22;
                const headerH = 26;
                
                doc.setFillColor(...C.navy);
                doc.rect(MARGIN, currentY, CW, headerH, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.setTextColor(...C.white);
                let x = MARGIN;
                headers.forEach(h => {
                    doc.text(h, x + 6, currentY + 17);
                    x += colWidth;
                });
                currentY += headerH;
                
                data.forEach((row, idx) => {
                    const bg = (idx % 2 === 0) ? C.white : C.light;
                    doc.setFillColor(...bg);
                    doc.rect(MARGIN, currentY, CW, rowH, 'F');
                    doc.setDrawColor(...C.light);
                    doc.line(MARGIN, currentY + rowH, MARGIN + CW, currentY + rowH);
                    
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(8);
                    doc.setTextColor(...C.text);
                    x = MARGIN;
                    row.forEach(cell => {
                        doc.text(String(cell || '-'), x + 6, currentY + 15);
                        x += colWidth;
                    });
                    currentY += rowH;
                });
                currentY += 8;
            }
            drawFooter();
        }
        
        // ========== BOM TABLE - CLEAN AND READABLE ==========
        if(sections.table) {
            setProgress(90, 'Building BOM table...');

            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Bill of Materials');

            const importantColumns = [
                'Part Number', 'MPN', 'Description',
                'Manufacturer', 'Supplier', 'Cost', 'Unit Price',
                'Risk', 'Lead Time', 'Total Cost'
            ];

            const availableColumns = [];
            const firstRow = bomData[0] || {};
            const columnKeys = Object.keys(firstRow);

            for (const importantCol of importantColumns) {
                const found = columnKeys.find(key =>
                    key.toLowerCase() === importantCol.toLowerCase() ||
                    key.toLowerCase().includes(importantCol.toLowerCase())
                );
                if (found && !availableColumns.includes(found)) {
                    availableColumns.push(found);
                }
            }

            if (availableColumns.length === 0) {
                availableColumns.push(...columnKeys.slice(0, 6));
            }

            const columns = availableColumns.slice(0, 7);
            const headers = columns.map(h => h);

            const widthWeights = columns.map(column => {
                const lower = column.toLowerCase();
                if (lower.includes('description')) return 2.5;
                if (lower.includes('part') || lower.includes('mpn')) return 1.5;
                if (lower.includes('manufacturer') || lower.includes('supplier')) return 1.3;
                if (lower.includes('quantity') || lower.includes('qty')) return 1.0;
                if (lower.includes('cost') || lower.includes('price')) return 1.2;
                return 1.0;
            });

            const totalWeight = widthWeights.reduce((sum, w) => sum + w, 0);
            const colWidths = widthWeights.map(w => (w / totalWeight) * CW);

            const tableRows = bomData.map(row => {
                return columns.map(col => {
                    let value = row[col];
                    if (value === undefined || value === null) value = '';
                    return String(value);
                });
            });

            const headerHeight = 30;

            function drawTableHeader() {
                doc.setFillColor(...C.navy);
                doc.rect(MARGIN, currentY, CW, headerHeight, 'F');
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(9);
                doc.setTextColor(...C.white);

                let x = MARGIN;
                headers.forEach((label, idx) => {
                    let displayHeader = label;
                    if (displayHeader.length > 22) {
                        displayHeader = displayHeader.substring(0, 19) + '...';
                    }
                    doc.text(displayHeader, x + 6, currentY + 19);
                    x += colWidths[idx];
                });
                currentY += headerHeight;
            }

            function drawRow(cells, rowIndex) {
                const cellLines = cells.map((cell, idx) => {
                    const maxWidth = colWidths[idx] - 10;
                    return doc.splitTextToSize(cell || '-', maxWidth);
                });
                const rowHeight = Math.max(...cellLines.map(lines => lines.length)) * 12 + 8;

                if (currentY + rowHeight > PH - 60) {
                    addPage();
                    currentY = MARGIN + 10;
                    drawTableHeader();
                }

                const bgColor = (rowIndex % 2 === 0) ? C.white : C.light;
                doc.setFillColor(...bgColor);
                doc.rect(MARGIN, currentY, CW, rowHeight, 'F');
                doc.setDrawColor(...C.light);
                doc.setLineWidth(0.5);

                let x = MARGIN;
                cellLines.forEach((lines, colIdx) => {
                    doc.line(x, currentY, x, currentY + rowHeight);
                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(8);
                    doc.setTextColor(...C.text);
                    doc.text(lines, x + 6, currentY + 14);
                    x += colWidths[colIdx];
                });
                doc.line(MARGIN + CW, currentY, MARGIN + CW, currentY + rowHeight);
                doc.line(MARGIN, currentY + rowHeight, MARGIN + CW, currentY + rowHeight);
                currentY += rowHeight;
            }

            drawTableHeader();
            tableRows.forEach((cells, idx) => drawRow(cells, idx));

            currentY += 14;
            addText(`Total ${bomData.length} components listed above`, 9, false, 'muted');
            drawFooter();
        }

        // ========== CONCLUSION ==========
        if(sections.conclusion) {
            setProgress(95, 'Writing conclusion...');
            addPage();
            currentY = MARGIN + 10;
            addSectionTitle('Conclusion');
            
            if(customConclusion) {
                addText(customConclusion, 11, false, 'text');
            } else {
                const riskCount = window.dimensionsData?.risk?.count || 0;
                const eolCount = window.dimensionsData?.lifecycle?.counts?.EOL || 0;
                const rohsCompliant = window.dimensionsData?.compliance?.rohs?.Yes || 0;
                const conclusionText = `This report analyzed ${bomData.length} components. ${riskCount} risk items were identified, with ${eolCount} components at End-of-Life. RoHS compliance is at ${Math.round(rohsCompliant/bomData.length*100)}%. Regular BOM reviews are recommended to maintain supply chain health.`;
                addText(conclusionText, 11, false, 'text');
            }
            
            currentY += 20;
            addText(`Report generated on ${formatDate(dateStr)}`, 9, false, 'muted');
            addText(`Data processed locally — no external transmission`, 9, false, 'muted');
            drawFooter();
        }
        
        // Save PDF
        setProgress(100, 'Saving PDF...');
        const filename = `${company.replace(/\s+/g, '_')}_BOM_Report_${dateStr.replace(/-/g, '')}.pdf`;
        doc.save(filename);
        
        setTimeout(() => {
            closeModal();
            setProgress(0, '');
        }, 800);
    }
    
    // Event Listeners
    function init() {
        const exportBtn = document.getElementById('exportPdfBtn');
        if(exportBtn) {
            exportBtn.addEventListener('click', () => {
                const dateInput = document.getElementById('reportDate');
                if(dateInput) dateInput.value = new Date().toISOString().split('T')[0];
                const modal = document.getElementById('pdf-modal');
                if(modal) modal.style.display = 'flex';
            });
        }
        
        const modalClose = document.getElementById('modalClose');
        if(modalClose) modalClose.addEventListener('click', closeModal);
        
        const cancelBtn = document.getElementById('cancelPdfBtn');
        if(cancelBtn) cancelBtn.addEventListener('click', closeModal);
        
        const modal = document.getElementById('pdf-modal');
        if(modal) modal.addEventListener('click', (e) => {
            if(e.target.id === 'pdf-modal') closeModal();
        });
        
        const logoInput = document.getElementById('logoInput');
        if(logoInput) {
            logoInput.addEventListener('change', function() {
                const file = this.files[0];
                if(!file) return;
                const reader = new FileReader();
                reader.onload = (ev) => {
                    logoDataUrl = ev.target.result;
                    const preview = document.getElementById('logoPreview');
                    if(preview) {
                        preview.className = 'logo-preview-img';
                        preview.innerHTML = `<img src="${logoDataUrl}" alt="Logo" style="max-height:50px"><button onclick="window.clearLogo && window.clearLogo()" style="margin-left:10px;background:#eee;color:#333;padding:4px 8px;border-radius:4px;">Remove</button>`;
                    }
                };
                reader.readAsDataURL(file);
            });
        }
        
        const generateBtn = document.getElementById('generatePdfBtn');
        if(generateBtn) {
            generateBtn.addEventListener('click', generatePDF);
        }
    }
    
    window.clearLogo = function() {
        logoDataUrl = null;
        const preview = document.getElementById('logoPreview');
        if(preview) {
            preview.className = 'logo-preview-empty';
            preview.innerHTML = '<span>Click to upload logo</span>';
            preview.onclick = () => document.getElementById('logoInput').click();
        }
        const logoInput = document.getElementById('logoInput');
        if(logoInput) logoInput.value = '';
    };
    
    if(document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();