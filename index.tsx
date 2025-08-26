import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

// --- TYPE DEFINITIONS ---
interface Equipment {
    name: string;
    diversity: { [standard: string]: number };
}
interface ApplianceData {
    [category: string]: Equipment[];
}
interface ApplianceState {
    id: string;
    category: string;
    equipment: string;
    rating: number;
    units: 'BTU/hr' | 'MMBTU/hr' | 'CFH';
    diversity: number;
    diversitySource?: string;
}
interface ProjectData {
    projectInfo: { [key: string]: any };
    appliances: ApplianceState[];
    gasEnergyContent: number;
}

// --- CONSTANTS ---
const DIVERSITY_STANDARDS: { [key: string]: string } = {
    default: "Utility Standard (Default)",
    ashrae: "ASHRAE (Conservative)",
    iapmo: "IAPMO (Lenient)"
};

const APPLIANCE_DATA: ApplianceData = {
    "Heating": [
        { name: "Forced Air Furnace (<225k BTU/hr)", diversity: { default: 75, ashrae: 90, iapmo: 70 } },
        { name: "Forced Air Furnace (>225k BTU/hr)", diversity: { default: 70, ashrae: 85, iapmo: 65 } },
        { name: "Hydronic Boiler (<300k BTU/hr)", diversity: { default: 75, ashrae: 90, iapmo: 70 } },
        { name: "Hydronic Boiler (>300k BTU/hr)", diversity: { default: 70, ashrae: 85, iapmo: 65 } },
        { name: "Combination Boiler (Combi)", diversity: { default: 80, ashrae: 85, iapmo: 75 } },
        { name: "Unit Heater (e.g., Garage)", diversity: { default: 75, ashrae: 80, iapmo: 70 } },
        { name: "Direct-Vent Wall Furnace", diversity: { default: 75, ashrae: 85, iapmo: 70 } },
        { name: "Radiant Tube Heater", diversity: { default: 90, ashrae: 95, iapmo: 85 } },
        { name: "Rooftop Unit (RTU)", diversity: { default: 90, ashrae: 95, iapmo: 85 } },
        { name: "Makeup Air Unit (MUA)", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
        { name: "Driveway Boiler", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
    ],
    "Hot Water": [
        { name: "Tank Water Heater (<50 gal)", diversity: { default: 35, ashrae: 50, iapmo: 30 } },
        { name: "Tank Water Heater (50-100 gal)", diversity: { default: 30, ashrae: 45, iapmo: 25 } },
        { name: "Tank Water Heater (>100 gal)", diversity: { default: 25, ashrae: 40, iapmo: 20 } },
        { name: "Power-Vent Water Heater", diversity: { default: 35, ashrae: 50, iapmo: 30 } },
        { name: "On-Demand Water Heater", diversity: { default: 25, ashrae: 30, iapmo: 20 } },
        { name: "Indirect Water Heater (from Boiler)", diversity: { default: 20, ashrae: 30, iapmo: 15 } },
        { name: "Booster Heater", diversity: { default: 20, ashrae: 25, iapmo: 15 } },
    ],
    "Cooking": [
        { name: "Range/Oven", diversity: { default: 15, ashrae: 20, iapmo: 10 } },
        { name: "Cooktop", diversity: { default: 10, ashrae: 15, iapmo: 8 } },
        { name: "Double Wall Oven", diversity: { default: 15, ashrae: 20, iapmo: 10 } },
        { name: "Commercial-Style Range", diversity: { default: 20, ashrae: 25, iapmo: 15 } },
        { name: "Grill (Built-in)", diversity: { default: 5, ashrae: 10, iapmo: 3 } },
        { name: "Outdoor Kitchen (Grill + Burners)", diversity: { default: 8, ashrae: 12, iapmo: 5 } },
        { name: "Deep Fryer", diversity: { default: 10, ashrae: 15, iapmo: 8 } },
    ],
    "Clothes Dryers": [
        { name: "Standard Dryer", diversity: { default: 25, ashrae: 30, iapmo: 20 } },
        { name: "High-Capacity Dryer", diversity: { default: 25, ashrae: 30, iapmo: 20 } },
        { name: "Gas Dryer/Steamer Combo", diversity: { default: 25, ashrae: 30, iapmo: 20 } },
    ],
    "Fireplaces & Stoves": [
        { name: "Gas Fireplace Insert", diversity: { default: 20, ashrae: 25, iapmo: 15 } },
        { name: "Direct-Vent Fireplace", diversity: { default: 20, ashrae: 25, iapmo: 15 } },
        { name: "Gas Stove", diversity: { default: 20, ashrae: 25, iapmo: 15 } },
        { name: "Gas Log Set", diversity: { default: 15, ashrae: 20, iapmo: 10 } },
        { name: "Outdoor Fire Pit", diversity: { default: 10, ashrae: 15, iapmo: 5 } },
        { name: "Patio Heater", diversity: { default: 50, ashrae: 60, iapmo: 40 } },
    ],
    "Pool & Spa Heaters": [
        { name: "Pool Heater", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
        { name: "Spa Heater", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
        { name: "Combined Pool/Spa Heater", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
    ],
    "Home Generators": [
        { name: "Air-cooled (7-22kW)", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
        { name: "Liquid-cooled (>22kW)", diversity: { default: 100, ashrae: 100, iapmo: 100 } },
    ],
    "Other Appliances": [
        { name: "Gas Lighting (e.g., Lanterns)", diversity: { default: 40, ashrae: 50, iapmo: 30 } },
        { name: "Incinerator", diversity: { default: 5, ashrae: 10, iapmo: 3 } },
        { name: "Gas Refrigerator", diversity: { default: 50, ashrae: 60, iapmo: 40 } },
    ],
    "Custom": [
        { name: "Custom Appliance", diversity: { default: 0, ashrae: 0, iapmo: 0 } },
    ]
};

const CATEGORY_DIVERSITY_DEFAULTS: { [category: string]: { [standard: string]: number } } = {
    "Heating": { default: 75, ashrae: 90, iapmo: 70 },
    "Hot Water": { default: 35, ashrae: 50, iapmo: 30 },
    "Cooking": { default: 15, ashrae: 20, iapmo: 10 },
    "Clothes Dryers": { default: 25, ashrae: 30, iapmo: 20 },
    "Fireplaces & Stoves": { default: 20, ashrae: 25, iapmo: 15 },
    "Pool & Spa Heaters": { default: 100, ashrae: 100, iapmo: 100 },
    "Home Generators": { default: 100, ashrae: 100, iapmo: 100 },
    "Other Appliances": { default: 50, ashrae: 60, iapmo: 40 },
};

// --- DOM ELEMENT SELECTORS ---
const applianceList = document.getElementById('appliance-list')!;
const addApplianceBtn = document.getElementById('add-appliance-btn')!;
const calculateBtn = document.getElementById('calculate-btn')!;
const resultsUnitsSelect = document.getElementById('results-units') as HTMLSelectElement;
const gasEnergyContentInput = document.getElementById('gas-energy-content') as HTMLInputElement;
const diversityStandardSelect = document.getElementById('diversity-standard') as HTMLSelectElement;
const totalConnectedLoadEl = document.getElementById('total-connected-load')!;
const totalDiversifiedLoadEl = document.getElementById('total-diversified-load')!;
const saveBtn = document.getElementById('save-btn')!;
const loadBtn = document.getElementById('load-btn')!;
const loadInput = document.getElementById('load-input') as HTMLInputElement;
const exportPdfBtn = document.getElementById('export-pdf-btn')!;
const form = document.getElementById('calculator-form') as HTMLFormElement;

// --- STATE ---
let calculatedResults = { connectedBtu: 0, diversifiedBtu: 0 };

// --- UTILITY FUNCTIONS ---
const generateId = () => `appliance-${Date.now()}-${Math.random()}`;

const convertToBtu = (value: number, unit: string, gasEnergy: number): number => {
    switch (unit) {
        case 'MMBTU/hr': return value * 1_000_000;
        case 'CFH': return value * gasEnergy;
        case 'BTU/hr':
        default: return value;
    }
};

const convertFromBtu = (value: number, unit: string, gasEnergy: number): number => {
    switch (unit) {
        case 'MMBTU/hr': return value / 1_000_000;
        case 'CFH': return value / gasEnergy;
        case 'BTU/hr':
        default: return value;
    }
};

// --- CORE FUNCTIONS ---
function addApplianceRow(appliance?: ApplianceState) {
    const rowId = appliance?.id || generateId();
    const row = document.createElement('div');
    row.className = 'appliance-row';
    row.id = rowId;

    const categories = Object.keys(APPLIANCE_DATA);
    const categoryOptions = categories.map(cat => `<option value="${cat}" ${appliance?.category === cat ? 'selected' : ''}>${cat}</option>`).join('');
    const diversitySourceOptions = Object.keys(CATEGORY_DIVERSITY_DEFAULTS).map(cat => `<option value="${cat}" ${appliance?.diversitySource === cat ? 'selected' : ''}>${cat}</option>`).join('');

    row.innerHTML = `
        <select class="category-select" aria-label="Appliance Category"><option value="" selected disabled>Select Category...</option>${categoryOptions}</select>
        <div class="equipment-container">
            <select class="equipment-select" aria-label="Equipment Type"></select>
            <input type="text" class="custom-equipment-input" style="display: none;" placeholder="Appliance Name" aria-label="Custom Equipment Name" value="${(appliance?.category === 'Custom' ? appliance.equipment : '')}">
        </div>
        <input type="number" class="rating-input" aria-label="Max Input Rating" value="${appliance?.rating ?? ''}" min="0" placeholder="e.g., 120000">
        <div class="radio-group">
            <label><input type="radio" name="units-${rowId}" value="BTU/hr" ${appliance?.units === 'BTU/hr' ? 'checked' : ''}> BTU/hr</label>
            <label><input type="radio" name="units-${rowId}" value="MMBTU/hr" ${appliance?.units === 'MMBTU/hr' ? 'checked' : ''}> MMBTU/hr</label>
            <label><input type="radio" name="units-${rowId}" value="CFH" ${appliance?.units === 'CFH' || !appliance ? 'checked' : ''}> CFH</label>
        </div>
        <div class="diversity-container">
            <input type="number" class="diversity-input" aria-label="Diversity Factor" value="${appliance?.diversity ?? 0}" min="0" max="100">
            <select class="diversity-source-select" style="display: none;" aria-label="Diversity Source Category" required>
               <option value="" disabled ${!appliance?.diversitySource ? 'selected' : ''}>Select category...</option>
               ${diversitySourceOptions}
            </select>
        </div>
        <button type="button" class="remove-btn button-danger">Remove</button>
    `;

    applianceList.appendChild(row);

    const categorySelect = row.querySelector('.category-select') as HTMLSelectElement;
    if (appliance?.category) {
        updateEquipmentOptions(categorySelect, appliance?.equipment);
        updateRowUIForCategory(row);
    }
}

function updateRowUIForCategory(row: HTMLElement) {
    const category = (row.querySelector('.category-select') as HTMLSelectElement).value;
    const equipmentSelect = row.querySelector('.equipment-select') as HTMLSelectElement;
    const customEquipmentInput = row.querySelector('.custom-equipment-input') as HTMLInputElement;
    const diversitySourceSelect = row.querySelector('.diversity-source-select') as HTMLSelectElement;
    const diversityInput = row.querySelector('.diversity-input') as HTMLInputElement;

    if (category === 'Custom') {
        equipmentSelect.style.display = 'none';
        customEquipmentInput.style.display = 'block';
        diversitySourceSelect.style.display = 'block';
        diversityInput.readOnly = true;
    } else {
        equipmentSelect.style.display = 'block';
        customEquipmentInput.style.display = 'none';
        diversitySourceSelect.style.display = 'none';
        diversityInput.readOnly = false;
    }
}

function updateEquipmentOptions(categorySelect: HTMLSelectElement, selectedEquipment?: string) {
    const row = categorySelect.closest('.appliance-row')!;
    const equipmentSelect = row.querySelector('.equipment-select') as HTMLSelectElement;
    const category = categorySelect.value;
    
    if (category === 'Custom') {
        equipmentSelect.innerHTML = '';
        return;
    }

    const equipmentList = APPLIANCE_DATA[category];
    equipmentSelect.innerHTML = `<option value="" selected disabled>Select Equipment...</option>` + equipmentList.map(eq => `<option value="${eq.name}" ${selectedEquipment === eq.name ? 'selected' : ''}>${eq.name}</option>`).join('');
}


function updateDiversityFromEquipment(equipmentSelect: HTMLSelectElement) {
    const row = equipmentSelect.closest('.appliance-row')!;
    const diversityInput = row.querySelector('.diversity-input') as HTMLInputElement;
    const category = (row.querySelector('.category-select') as HTMLSelectElement).value;
    const standard = diversityStandardSelect.value;

    const equipmentList = APPLIANCE_DATA[category];
    if (!equipmentList) return;

    const selected = equipmentList.find(eq => eq.name === equipmentSelect.value);
    if (selected) {
        diversityInput.value = selected.diversity[standard].toString();
    }
}

function updateAllDiversityFactors() {
    const newStandard = diversityStandardSelect.value;
    document.querySelectorAll('.appliance-row').forEach(row => {
        if (row.classList.contains('header')) return;
        
        const diversityInput = row.querySelector('.diversity-input') as HTMLInputElement;
        const category = (row.querySelector('.category-select') as HTMLSelectElement).value;

        if (category === 'Custom') {
            const sourceCategory = (row.querySelector('.diversity-source-select') as HTMLSelectElement).value as keyof typeof CATEGORY_DIVERSITY_DEFAULTS;
            if (sourceCategory && CATEGORY_DIVERSITY_DEFAULTS[sourceCategory]) {
                diversityInput.value = CATEGORY_DIVERSITY_DEFAULTS[sourceCategory][newStandard].toString();
            }
        } else {
            const equipmentName = (row.querySelector('.equipment-select') as HTMLSelectElement).value;
            const equipmentData = APPLIANCE_DATA[category]?.find(eq => eq.name === equipmentName);
            if (equipmentData) {
                diversityInput.value = equipmentData.diversity[newStandard].toString();
            }
        }
    });
}


function calculate() {
    const gasEnergy = parseFloat(gasEnergyContentInput.value) || 1040;
    let totalConnected = 0;
    let totalDiversified = 0;

    document.querySelectorAll('.appliance-row').forEach(row => {
        if (row.classList.contains('header')) return;
        
        const rating = parseFloat((row.querySelector('.rating-input') as HTMLInputElement).value) || 0;
        const units = (row.querySelector('input[type="radio"]:checked') as HTMLInputElement).value;
        const diversity = parseFloat((row.querySelector('.diversity-input') as HTMLInputElement).value) || 0;

        const ratingInBtu = convertToBtu(rating, units, gasEnergy);
        totalConnected += ratingInBtu;
        totalDiversified += ratingInBtu * (diversity / 100);
    });

    calculatedResults = { connectedBtu: totalConnected, diversifiedBtu: totalDiversified };
    updateResultDisplay();
}

function updateResultDisplay() {
    const displayUnit = resultsUnitsSelect.value;
    const gasEnergy = parseFloat(gasEnergyContentInput.value) || 1040;

    const connected = convertFromBtu(calculatedResults.connectedBtu, displayUnit, gasEnergy);
    const diversified = convertFromBtu(calculatedResults.diversifiedBtu, displayUnit, gasEnergy);

    totalConnectedLoadEl.textContent = `${connected.toLocaleString(undefined, {maximumFractionDigits: 2})} ${displayUnit}`;
    totalDiversifiedLoadEl.textContent = `${diversified.toLocaleString(undefined, {maximumFractionDigits: 2})} ${displayUnit}`;
}

function getFormData(): ProjectData {
    const projectInfo = Object.fromEntries(new FormData(form).entries());
    const appliances: ApplianceState[] = [];
    document.querySelectorAll('.appliance-row').forEach(row => {
        if (row.classList.contains('header')) return;
        const category = (row.querySelector('.category-select') as HTMLSelectElement).value;
        const equipment = category === 'Custom'
            ? (row.querySelector('.custom-equipment-input') as HTMLInputElement).value
            : (row.querySelector('.equipment-select') as HTMLSelectElement).value;
        const diversitySource = category === 'Custom'
            ? (row.querySelector('.diversity-source-select') as HTMLSelectElement).value
            : undefined;
            
        appliances.push({
            id: row.id,
            category: category,
            equipment: equipment,
            rating: parseFloat((row.querySelector('.rating-input') as HTMLInputElement).value) || 0,
            units: (row.querySelector('input[type="radio"]:checked') as HTMLInputElement).value as any,
            diversity: parseFloat((row.querySelector('.diversity-input') as HTMLInputElement).value) || 0,
            diversitySource: diversitySource,
        });
    });
    return { 
      projectInfo, 
      appliances, 
      gasEnergyContent: parseFloat(gasEnergyContentInput.value) 
    };
}

function loadFormData(data: ProjectData) {
    Object.entries(data.projectInfo).forEach(([key, value]) => {
        const el = form.elements.namedItem(key) as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement;
        if (el) el.value = value;
    });

    applianceList.innerHTML = '';
    data.appliances.forEach(app => addApplianceRow(app));
    gasEnergyContentInput.value = data.gasEnergyContent.toString();
    calculate();
}

function save() {
    const data = getFormData();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${data.projectInfo.projectName || 'load-assessment'}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function load() {
    const file = loadInput.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = JSON.parse(e.target?.result as string);
            loadFormData(data);
        } catch (error) {
            alert('Error: Could not parse file. Please ensure it is a valid JSON file.');
            console.error(error);
        }
    };
    reader.readAsText(file);
}

function exportToPdf() {
    const doc = new jsPDF();
    const data = getFormData();
    calculate(); // Ensure calculations are fresh
    
    const pageHeight = doc.internal.pageSize.height;
    let y = 15;

    // Header
    doc.setFontSize(18);
    doc.setTextColor(40);
    doc.text("Residential Load Diversity Report", 105, y, { align: 'center' });
    y += 10;

    // Project Details
    doc.setFontSize(14);
    doc.text("Project Details", 14, y);
    y += 2;
    autoTable(doc, {
        startY: y,
        body: [
            ['Project Name', data.projectInfo.projectName],
            ['Address', `${data.projectInfo.streetAddress}, ${data.projectInfo.town}, ${data.projectInfo.state}`],
            ['DOC', data.projectInfo.doc],
            ['G-Intake #', data.projectInfo.gIntake],
            ['Assessed by', data.projectInfo.assessedBy],
            ['Assessment Date', data.projectInfo.assessmentDate],
            ['Revision #', `${data.projectInfo.revisionNumber} (on ${data.projectInfo.revisionDate})`],
        ],
        theme: 'striped',
        styles: { fontSize: 10 }
    });
    // @ts-ignore
    y = doc.lastAutoTable.finalY + 10;

    // Executive Summary
    doc.setFontSize(12);
    doc.text("Executive Summary", 14, y);
    y += 6;
    doc.setFontSize(10);
    const summaryText = `This report documents the calculated natural gas load for the residential property at ${data.projectInfo.streetAddress}, ${data.projectInfo.town}, ${data.projectInfo.state}. It outlines the total connected load and the diversified load based on the listed appliances and standard diversity methodologies.`;
    const splitSummary = doc.splitTextToSize(summaryText, 180);
    doc.text(splitSummary, 14, y);
    y += splitSummary.length * 4 + 10;
    
    // Appliance Table
    doc.setFontSize(12);
    doc.text("Appliance Load & Diversity Calculation", 14, y);
    y += 2;
    const applianceBody = data.appliances.map(app => {
      const ratingBtu = convertToBtu(app.rating, app.units, data.gasEnergyContent);
      return [
        app.equipment,
        ratingBtu.toLocaleString(),
        `${app.diversity}%`,
        (ratingBtu * app.diversity / 100).toLocaleString()
      ]
    });
    autoTable(doc, {
      startY: y,
      head: [['Equipment Type', 'Max Input (BTU/hr)', 'Diversity Factor', 'Diversified Load (BTU/hr)']],
      body: applianceBody,
      theme: 'grid',
      styles: { fontSize: 9 }
    });
    // @ts-ignore
    y = doc.lastAutoTable.finalY + 10;
    
    if (y > pageHeight - 60) {
      doc.addPage();
      y = 20;
    }

    // Calculation Methodology
    const standardKey = diversityStandardSelect.value;
    const standardName = DIVERSITY_STANDARDS[standardKey] || 'the selected';

    doc.setFontSize(12);
    doc.text("Calculation Methodology & Sources", 14, y);
    y += 6;
    doc.setFontSize(10);
    const methText = `The Diversified Load for each appliance is calculated as: Max Input Rating × (Diversity Factor / 100). The Total Diversified Load is the sum of these individual values. The default diversity factors used are representative values derived from an analysis of common North American gas utility engineering practices and are intended for estimation. Specifically, these values are consistent with methodologies found in sources such as the ASHRAE Handbook—Fundamentals. The diversity factors used in this calculation are based on the ${standardName} standard. For code-compliant sizing, factors should be verified with the local utility or a qualified engineer. Gas energy content was assumed to be ${data.gasEnergyContent} BTU/SCF.`;
    const splitMeth = doc.splitTextToSize(methText, 180);
    doc.text(splitMeth, 14, y);
    y += splitMeth.length * 4 + 10;
    
    // Final Summary
    doc.setFontSize(12);
    doc.text("Final Summary", 14, y);
    y += 2;
    autoTable(doc, {
        startY: y,
        body: [
            [{ content: 'Total Connected Load (Without Diversity)', styles: { fontStyle: 'bold' } }, `${calculatedResults.connectedBtu.toLocaleString()} BTU/hr`],
            [{ content: 'Total Diversified Load (With Diversity)', styles: { fontStyle: 'bold' } }, `${calculatedResults.diversifiedBtu.toLocaleString()} BTU/hr`],
        ],
        theme: 'grid',
        styles: { fontSize: 10 }
    });
    // @ts-ignore
    y = doc.lastAutoTable.finalY + 10;


    // Footer on all pages
    const pageCount = (doc as any).internal.getNumberOfPages();
    for(let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text(`Report generated on ${new Date().toLocaleDateString()}`, 14, pageHeight - 10);
        doc.text(`Page ${i} of ${pageCount}`, doc.internal.pageSize.width - 35, pageHeight - 10);
    }
    
    doc.save(`${data.projectInfo.projectName || 'load-report'}.pdf`);
}

function createDiversityTable() {
    const tableContainer = document.getElementById('table-tab');
    if (!tableContainer) return;

    const mainHeading = document.createElement('h2');
    mainHeading.textContent = 'Appliance Diversity Factor Reference Tables';
    tableContainer.appendChild(mainHeading);

    const categories = Object.keys(APPLIANCE_DATA);

    for (const category of categories) {
        if (category === 'Custom') continue;

        const categoryData = APPLIANCE_DATA[category];
        if (categoryData.length === 0) continue;

        const categoryHeading = document.createElement('h3');
        categoryHeading.textContent = category;
        tableContainer.appendChild(categoryHeading);

        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        const headerRow = document.createElement('tr');
        const headers = ['Equipment Type', DIVERSITY_STANDARDS.default, DIVERSITY_STANDARDS.ashrae, DIVERSITY_STANDARDS.iapmo];
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        categoryData.forEach(equipment => {
            const row = document.createElement('tr');
            
            const cellName = document.createElement('td');
            cellName.textContent = equipment.name;
            row.appendChild(cellName);

            const cellDefault = document.createElement('td');
            cellDefault.textContent = `${equipment.diversity.default}%`;
            row.appendChild(cellDefault);

            const cellAshrae = document.createElement('td');
            cellAshrae.textContent = `${equipment.diversity.ashrae}%`;
            row.appendChild(cellAshrae);
            
            const cellIapmo = document.createElement('td');
            cellIapmo.textContent = `${equipment.diversity.iapmo}%`;
            row.appendChild(cellIapmo);

            tbody.appendChild(row);
        });

        table.appendChild(thead);
        table.appendChild(tbody);
        tableContainer.appendChild(table);
    }
}


// --- EVENT LISTENERS ---
function setupEventListeners() {
    // Tab switching logic
    const tabsContainer = document.querySelector('.tabs');
    tabsContainer?.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        if (target.classList.contains('tab-link')) {
            const tabId = target.dataset.tab;
            if (tabId) {
                document.querySelectorAll('.tab-link').forEach(link => link.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
                
                target.classList.add('active');
                document.getElementById(tabId)?.classList.add('active');
            }
        }
    });

    addApplianceBtn.addEventListener('click', () => addApplianceRow());
    
    applianceList.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        if (target.classList.contains('remove-btn')) {
            target.closest('.appliance-row')?.remove();
        }
    });

    applianceList.addEventListener('change', (e) => {
        const target = e.target as HTMLElement;
        const row = target.closest('.appliance-row');
        if (!row) return;

        if (target.classList.contains('category-select')) {
            const categorySelect = target as HTMLSelectElement;
            const category = categorySelect.value;
            
            updateRowUIForCategory(row as HTMLElement);
            updateEquipmentOptions(categorySelect);

            if (category === 'Custom') {
                (row.querySelector('.rating-input') as HTMLInputElement).value = '0';
                (row.querySelector('.diversity-input') as HTMLInputElement).value = '0';
                (row.querySelector('.custom-equipment-input') as HTMLInputElement).value = '';
                (row.querySelector('.diversity-source-select') as HTMLSelectElement).value = '';
            } else {
                const equipmentSelect = row.querySelector('.equipment-select') as HTMLSelectElement;
                updateDiversityFromEquipment(equipmentSelect);
            }
        }

        if (target.classList.contains('equipment-select')) {
            updateDiversityFromEquipment(target as HTMLSelectElement);
        }

        if (target.classList.contains('diversity-source-select')) {
            const standard = diversityStandardSelect.value;
            const sourceCategory = (target as HTMLSelectElement).value as keyof typeof CATEGORY_DIVERSITY_DEFAULTS;
            if (sourceCategory && CATEGORY_DIVERSITY_DEFAULTS[sourceCategory]) {
                const diversityInput = row.querySelector('.diversity-input') as HTMLInputElement;
                diversityInput.value = CATEGORY_DIVERSITY_DEFAULTS[sourceCategory][standard].toString();
            }
        }
    });

    calculateBtn.addEventListener('click', calculate);
    resultsUnitsSelect.addEventListener('change', updateResultDisplay);
    gasEnergyContentInput.addEventListener('change', updateResultDisplay);
    diversityStandardSelect.addEventListener('change', updateAllDiversityFactors);
    
    saveBtn.addEventListener('click', save);
    loadBtn.addEventListener('click', () => loadInput.click());
    loadInput.addEventListener('change', load);
    
    exportPdfBtn.addEventListener('click', exportToPdf);
}

// --- INITIALIZATION ---
function init() {
    const today = new Date().toISOString().split('T')[0];
    (document.getElementById('assessment-date') as HTMLInputElement).value = today;
    (document.getElementById('revision-date') as HTMLInputElement).value = today;
    
    createDiversityTable();
    setupEventListeners();
}

init();