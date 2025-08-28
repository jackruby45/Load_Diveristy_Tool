// FIX: Changed import to use default export for jsPDF
import jsPDF from 'jspdf';
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

const SEASONAL_USAGE: { [category: string]: 'winter' | 'summer' | 'year-round' } = {
    "Heating": 'winter',
    "Hot Water": 'year-round',
    "Cooking": 'year-round',
    "Clothes Dryers": 'year-round',
    "Fireplaces & Stoves": 'winter',
    "Pool & Spa Heaters": 'summer',
    "Home Generators": 'year-round', // Critical load, can run anytime
    "Other Appliances": 'year-round',
    "Custom": 'year-round'
};

// --- DOM ELEMENT SELECTORS ---
const applianceList = document.getElementById('appliance-list')!;
const addApplianceBtn = document.getElementById('add-appliance-btn')!;
const resultsUnitsSelect = document.getElementById('results-units') as HTMLSelectElement;
const gasEnergyContentInput = document.getElementById('gas-energy-content') as HTMLInputElement;
const diversityStandardSelect = document.getElementById('diversity-standard') as HTMLSelectElement;
const totalConnectedLoadEl = document.getElementById('total-connected-load')!;
const winterPeakLoadEl = document.getElementById('winter-peak-load')!;
const summerPeakLoadEl = document.getElementById('summer-peak-load')!;
const totalDiversifiedLoadEl = document.getElementById('total-diversified-load')!;
const saveBtn = document.getElementById('save-btn')!;
const loadBtn = document.getElementById('load-btn')!;
const loadInput = document.getElementById('load-input') as HTMLInputElement;
const exportPdfBtn = document.getElementById('export-pdf-btn')!;
const form = document.getElementById('calculator-form') as HTMLFormElement;
const methodologyDisplayEl = document.getElementById('methodology-display')!;
const showDetailsBtn = document.getElementById('show-details-btn') as HTMLButtonElement;
const calculationDetailsEl = document.getElementById('calculation-details')!;

// --- STATE ---
let calculatedResults = { connectedBtu: 0, diversifiedBtu: 0, winterPeakBtu: 0, summerPeakBtu: 0 };

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

function showToast(message: string, duration = 5000, type: 'error' | 'success' = 'error') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;

    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('show');
    }, 100);

    setTimeout(() => {
        toast.classList.remove('show');
        toast.addEventListener('transitionend', () => toast.remove());
    }, duration);
}


// --- CORE FUNCTIONS ---
function updateMethodologyDisplay() {
    if (!methodologyDisplayEl) return;
    const standardKey = diversityStandardSelect.value;
    const standardName = DIVERSITY_STANDARDS[standardKey] || 'the selected';

    methodologyDisplayEl.innerHTML = `
        <p>Calculations are based on the principle of load diversity, where the diversified load for each appliance is determined by applying a specific factor to its maximum input rating.</p>
        <div class="formula-box">
            Diversified Load = Max Input Rating × (Diversity Factor / 100)
        </div>
        <p>The <strong>Design Load</strong> is the greater of the calculated <strong>Winter Peak</strong> and <strong>Summer Peak</strong> diversified loads. The diversity factors currently applied are from the <strong>${standardName}</strong> standard.</p>
    `;
}


function updateEmptyState() {
    const emptyState = document.getElementById('empty-state');
    if (!emptyState) return;
    const hasAppliances = applianceList.querySelector('.appliance-row:not(.header)');
    emptyState.style.display = hasAppliances ? 'none' : 'block';
}

function addApplianceRow(appliance?: ApplianceState) {
    const rowId = appliance?.id || generateId();
    const row = document.createElement('div');
    row.className = 'appliance-row';
    row.id = rowId;

    const categories = Object.keys(APPLIANCE_DATA);
    const categoryOptions = categories.map(cat => `<option value="${cat}" ${appliance?.category === cat ? 'selected' : ''}>${cat}</option>`).join('');
    const diversitySourceOptions = Object.keys(CATEGORY_DIVERSITY_DEFAULTS).map(cat => `<option value="${cat}" ${appliance?.diversitySource === cat ? 'selected' : ''}>${cat}</option>`).join('');

    row.innerHTML = `
        <div class="form-group">
            <label for="category-${rowId}">Category</label>
            <select id="category-${rowId}" class="category-select"><option value="" selected disabled>Select Category...</option>${categoryOptions}</select>
        </div>
        <div class="form-group">
            <label for="equipment-select-${rowId}">Equipment Type</label>
            <div class="equipment-container">
                <select id="equipment-select-${rowId}" class="equipment-select"></select>
                <input type="text" class="custom-equipment-input" style="display: none;" placeholder="Appliance Name" value="${(appliance?.category === 'Custom' ? appliance.equipment : '')}">
            </div>
        </div>
        <div class="form-group">
            <label for="rating-${rowId}">Max Input Rating</label>
            <input type="number" id="rating-${rowId}" class="rating-input" value="${appliance?.rating ?? ''}" min="0" placeholder="e.g., 120000">
        </div>
        <div class="form-group">
            <label for="units-${rowId}">Units</label>
            <select id="units-${rowId}" class="units-select">
                <option value="CFH" ${!appliance || appliance.units === 'CFH' ? 'selected' : ''}>CFH</option>
                <option value="BTU/hr" ${appliance?.units === 'BTU/hr' ? 'selected' : ''}>BTU/hr</option>
                <option value="MMBTU/hr" ${appliance?.units === 'MMBTU/hr' ? 'selected' : ''}>MMBTU/hr</option>
            </select>
        </div>
        <div class="form-group">
            <label for="diversity-${rowId}">Diversity Factor (%)</label>
            <div class="diversity-container">
                <input type="number" id="diversity-${rowId}" class="diversity-input" value="${appliance?.diversity ?? 0}" min="0" max="100">
                <select class="diversity-source-select" style="display: none;" required>
                   <option value="" disabled ${!appliance?.diversitySource ? 'selected' : ''}>Select category...</option>
                   ${diversitySourceOptions}
                </select>
            </div>
        </div>
        <div class="form-group action-group">
            <label class="action-label">&nbsp;</label>
            <button type="button" class="remove-btn button-danger">Remove</button>
        </div>
    `;

    // Add row to DOM and animate it in
    applianceList.appendChild(row);
    
    requestAnimationFrame(() => {
      row.classList.add('show');
    });


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
    const gasEnergy = parseFloat(gasEnergyContentInput.value) || 1036;
    let totalConnected = 0;
    let winterPeak = 0;
    let summerPeak = 0;

    document.querySelectorAll('.appliance-row').forEach(row => {
        if (row.classList.contains('header')) return;
        
        const rating = parseFloat((row.querySelector('.rating-input') as HTMLInputElement).value) || 0;
        const units = (row.querySelector('.units-select') as HTMLSelectElement).value;
        const diversity = parseFloat((row.querySelector('.diversity-input') as HTMLInputElement).value) || 0;
        const category = (row.querySelector('.category-select') as HTMLSelectElement).value;

        const ratingInBtu = convertToBtu(rating, units, gasEnergy);
        totalConnected += ratingInBtu;
        const diversifiedBtu = ratingInBtu * (diversity / 100);

        const season = SEASONAL_USAGE[category] || 'year-round';

        if (season === 'winter' || season === 'year-round') {
            winterPeak += diversifiedBtu;
        }
        if (season === 'summer' || season === 'year-round') {
            summerPeak += diversifiedBtu;
        }
    });

    const designLoad = Math.max(winterPeak, summerPeak);

    calculatedResults = { 
        connectedBtu: totalConnected, 
        diversifiedBtu: designLoad,
        winterPeakBtu: winterPeak,
        summerPeakBtu: summerPeak,
    };
    updateResultDisplay();
    updateCalculationDetails();
}

function updateResultDisplay() {
    const displayUnit = resultsUnitsSelect.value;
    const gasEnergy = parseFloat(gasEnergyContentInput.value) || 1036;

    const connected = convertFromBtu(calculatedResults.connectedBtu, displayUnit, gasEnergy);
    const winter = convertFromBtu(calculatedResults.winterPeakBtu, displayUnit, gasEnergy);
    const summer = convertFromBtu(calculatedResults.summerPeakBtu, displayUnit, gasEnergy);
    const diversified = convertFromBtu(calculatedResults.diversifiedBtu, displayUnit, gasEnergy);

    totalConnectedLoadEl.textContent = `${connected.toLocaleString(undefined, {maximumFractionDigits: 2})} ${displayUnit}`;
    winterPeakLoadEl.textContent = `${winter.toLocaleString(undefined, {maximumFractionDigits: 2})} ${displayUnit}`;
    summerPeakLoadEl.textContent = `${summer.toLocaleString(undefined, {maximumFractionDigits: 2})} ${displayUnit}`;
    totalDiversifiedLoadEl.textContent = `${diversified.toLocaleString(undefined, {maximumFractionDigits: 2})} ${displayUnit}`;
}

function updateCalculationDetails() {
    const displayUnit = resultsUnitsSelect.value;
    const gasEnergy = parseFloat(gasEnergyContentInput.value) || 1036;

    const seasonalAppliances: { [key: string]: { name: string, rating: number, diversity: number, diversifiedValue: number }[] } = {
        winter: [],
        summer: [],
        'year-round': []
    };
    
    let subtotalWinter = 0;
    let subtotalSummer = 0;
    let subtotalYearRound = 0;

    document.querySelectorAll('.appliance-row').forEach(row => {
        if (row.classList.contains('header')) return;
        const rating = parseFloat((row.querySelector('.rating-input') as HTMLInputElement).value) || 0;
        const units = (row.querySelector('.units-select') as HTMLSelectElement).value;
        const diversity = parseFloat((row.querySelector('.diversity-input') as HTMLInputElement).value) || 0;
        const category = (row.querySelector('.category-select') as HTMLSelectElement).value;
        const equipment = (row.querySelector('.equipment-select') as HTMLSelectElement).value || (row.querySelector('.custom-equipment-input') as HTMLInputElement).value;
        
        if (!category || !equipment) return;
        
        const ratingInBtu = convertToBtu(rating, units, gasEnergy);
        const diversifiedBtu = ratingInBtu * (diversity / 100);
        
        const ratingDisplay = convertFromBtu(ratingInBtu, displayUnit, gasEnergy);
        const diversifiedDisplay = convertFromBtu(diversifiedBtu, displayUnit, gasEnergy);
        
        const season = SEASONAL_USAGE[category] || 'year-round';
        seasonalAppliances[season].push({ name: equipment, rating: ratingDisplay, diversity: diversity, diversifiedValue: diversifiedDisplay });
        
        if (season === 'winter') subtotalWinter += diversifiedDisplay;
        if (season === 'summer') subtotalSummer += diversifiedDisplay;
        if (season === 'year-round') subtotalYearRound += diversifiedDisplay;
    });
    
    const formatValue = (val: number) => val.toLocaleString(undefined, { maximumFractionDigits: 2 });
    
    const renderList = (items: {name: string, rating: number, diversity: number, diversifiedValue: number}[]) => {
        if (items.length === 0) return '<li class="no-items">None</li>';
        return items.map(item => `
            <li class="calculation-line">
                <div class="item-name">${item.name}</div>
                <div class="item-formula">
                    <span class="formula-part formula-rating">${formatValue(item.rating)} ${displayUnit}</span>
                    <span class="formula-part formula-operator">×</span>
                    <span class="formula-part formula-diversity">${item.diversity}%</span>
                    <span class="formula-part formula-operator">=</span>
                    <span class="formula-part formula-result">${formatValue(item.diversifiedValue)} ${displayUnit}</span>
                </div>
            </li>`).join('');
    };

    const winterTotal = subtotalWinter + subtotalYearRound;
    const summerTotal = subtotalSummer + subtotalYearRound;

    calculationDetailsEl.innerHTML = `
        <div class="details-column">
            <h4>Winter Peak Calculation</h4>
            <ul class="details-list">
                <li class="category-title"><strong>Winter Loads</strong></li>
                ${renderList(seasonalAppliances.winter)}
                <li class="subtotal"><strong>Subtotal: ${formatValue(subtotalWinter)} ${displayUnit}</strong></li>
                
                <li class="category-title"><strong>Year-Round Loads</strong></li>
                ${renderList(seasonalAppliances['year-round'])}
                <li class="subtotal"><strong>Subtotal: ${formatValue(subtotalYearRound)} ${displayUnit}</strong></li>
            </ul>
            <div class="final-calculation">
                <strong>Total Winter Peak:</strong>
                <span>${formatValue(subtotalWinter)} + ${formatValue(subtotalYearRound)} = <strong>${formatValue(winterTotal)} ${displayUnit}</strong></span>
            </div>
        </div>
        <div class="details-column">
            <h4>Summer Peak Calculation</h4>
            <ul class="details-list">
                <li class="category-title"><strong>Summer Loads</strong></li>
                ${renderList(seasonalAppliances.summer)}
                <li class="subtotal"><strong>Subtotal: ${formatValue(subtotalSummer)} ${displayUnit}</strong></li>

                <li class="category-title"><strong>Year-Round Loads</strong></li>
                ${renderList(seasonalAppliances['year-round'])}
                <li class="subtotal"><strong>Subtotal: ${formatValue(subtotalYearRound)} ${displayUnit}</strong></li>
            </ul>
            <div class="final-calculation">
                <strong>Total Summer Peak:</strong>
                <span>${formatValue(subtotalSummer)} + ${formatValue(subtotalYearRound)} = <strong>${formatValue(summerTotal)} ${displayUnit}</strong></span>
            </div>
        </div>
    `;
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
            units: (row.querySelector('.units-select') as HTMLSelectElement).value as any,
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
    updateEmptyState();
    calculate();
}

function save() {
    try {
        const data = getFormData();
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const suggestedName = `${data.projectInfo.projectName || 'load-assessment'}.json`;

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = suggestedName;
        document.body.appendChild(a);
        a.click();
        
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (error) {
        console.error('Error preparing file for save:', error);
        showToast('Error: Could not prepare the file for download.', 5000, 'error');
    }
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
            showToast('Error: Could not parse file. Please ensure it is a valid JSON file.', 5000, 'error');
            console.error(error);
        }
    };
    reader.readAsText(file);
}

function exportToPdf() {
    const btn = exportPdfBtn as HTMLButtonElement;
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.classList.add('loading');
    btn.textContent = 'Generating...';

    setTimeout(() => {
        try {
            const doc = new jsPDF();
            const data = getFormData();
            calculate(); 
            
            const displayUnit = resultsUnitsSelect.value;
            const gasEnergy = data.gasEnergyContent;
            const pageHeight = doc.internal.pageSize.height;
            const pageWidth = doc.internal.pageSize.width;
            let y = 0;

            const addHeader = (docInstance: jsPDF, pageData: ProjectData) => {
                docInstance.setFontSize(10);
                docInstance.setTextColor(40);
                docInstance.setFont(undefined, 'bold');
                docInstance.text('Residential Load Diversity Report', 14, 15);
                docInstance.setFont(undefined, 'normal');
                docInstance.text(`${pageData.projectInfo.projectName || 'N/A'}`, pageWidth / 2, 15, { align: 'center' });
                docInstance.text(`${pageData.projectInfo.streetAddress || 'N/A'}`, pageWidth - 14, 15, { align: 'right' });
                docInstance.setDrawColor(150);
                docInstance.line(14, 18, pageWidth - 14, 18);
            };

            const checkPageBreak = (docInstance: jsPDF, currentY: number, requiredSpace: number) => {
                if (currentY + requiredSpace > pageHeight - 20) {
                    docInstance.addPage();
                    return 25;
                }
                return currentY;
            };

            // Page 1: Project Details & Summary
            addHeader(doc, data);
            y = 25;

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
                theme: 'striped', styles: { fontSize: 10 }
            });
            // @ts-ignore
            y = doc.lastAutoTable.finalY + 10;

            // Appliance List
            y = checkPageBreak(doc, y, 60);
            doc.setFontSize(12);
            doc.text("Appliance List & Individual Loads", 14, y);
            y += 6;
            const applianceBody = data.appliances.map(app => {
                const ratingBtu = convertToBtu(app.rating, app.units, gasEnergy);
                const diversifiedBtu = ratingBtu * app.diversity / 100;
                const displayRating = convertFromBtu(ratingBtu, displayUnit, gasEnergy);
                const displayDiversified = convertFromBtu(diversifiedBtu, displayUnit, gasEnergy);
                return [
                    app.equipment,
                    displayRating.toLocaleString(undefined, { maximumFractionDigits: 2 }),
                    `${app.diversity}%`,
                    displayDiversified.toLocaleString(undefined, { maximumFractionDigits: 2 })
                ];
            });
            autoTable(doc, {
                startY: y,
                head: [['Equipment Type', `Max Input (${displayUnit})`, 'Diversity Factor', `Diversified Load (${displayUnit})`]],
                body: applianceBody,
                theme: 'grid', styles: { fontSize: 9 }
            });
            // @ts-ignore
            y = doc.lastAutoTable.finalY + 10;
            
            // Detailed Calculation Analysis
            y = checkPageBreak(doc, y, 100);
            if(y === 25) doc.addPage();
            doc.setFontSize(14);
            doc.text("Detailed Load Calculation Analysis", 14, y);
            y += 10;

            const formatVal = (val: number) => val.toLocaleString(undefined, {maximumFractionDigits: 2});
            const winterApps = data.appliances.filter(a => SEASONAL_USAGE[a.category] === 'winter');
            const summerApps = data.appliances.filter(a => SEASONAL_USAGE[a.category] === 'summer');
            const yearRoundApps = data.appliances.filter(a => SEASONAL_USAGE[a.category] === 'year-round' || !SEASONAL_USAGE[a.category]);

            const generateSeasonalTable = (title: string, apps: ApplianceState[]) => {
                let subtotal = 0;
                const body = apps.map(app => {
                    const ratingBtu = convertToBtu(app.rating, app.units, gasEnergy);
                    const diversifiedBtu = ratingBtu * (app.diversity / 100);
                    const displayDiv = convertFromBtu(diversifiedBtu, displayUnit, gasEnergy);
                    subtotal += displayDiv;
                    return [
                        app.equipment,
                        `${formatVal(convertFromBtu(ratingBtu, displayUnit, gasEnergy))} ${displayUnit}`,
                        `${app.diversity}%`,
                        `${formatVal(displayDiv)} ${displayUnit}`
                    ];
                });
                autoTable(doc, {
                    startY: y,
                    head: [[{ content: title, colSpan: 4, styles: { fontStyle: 'bold', fillColor: '#e9ecef', textColor: '#212529' } }]],
                    body: body.length > 0 ? body : [['No appliances in this category', '', '', '']],
                    foot: [[{ content: 'Subtotal', colSpan: 3, styles: { halign: 'right', fontStyle: 'bold' } }, { content: `${formatVal(subtotal)} ${displayUnit}`, styles: { fontStyle: 'bold' } }]],
                    theme: 'grid', styles: { fontSize: 9 }, headStyles: { halign: 'center' }
                });
                // @ts-ignore
                y = doc.lastAutoTable.finalY;
                return subtotal;
            };

            doc.setFontSize(12);
            doc.text("Winter Peak Calculation", 14, y);
            y += 6;
            const winterSub = generateSeasonalTable('Winter Loads', winterApps);
            y = checkPageBreak(doc, y, 40) + 2;
            const yrSub1 = generateSeasonalTable('Year-Round Loads', yearRoundApps);
            const winterTotal = winterSub + yrSub1;
            y = checkPageBreak(doc, y, 15) + 8;
            doc.setFontSize(10);
            doc.setFillColor('#e7f1ff');
            doc.rect(14, y - 4.5, pageWidth - 28, 12, 'F');
            doc.text(`Total Winter Peak = ${formatVal(winterSub)} (${'Winter'}) + ${formatVal(yrSub1)} (${'Year-Round'}) = `, 16, y);
            doc.setFont(undefined, 'bold');
            doc.text(`${formatVal(winterTotal)} ${displayUnit}`, doc.getTextWidth(`Total Winter Peak = ${formatVal(winterSub)} (Winter) + ${formatVal(yrSub1)} (Year-Round) = `) + 16, y);
            doc.setFont(undefined, 'normal');
            y += 15;

            y = checkPageBreak(doc, y, 100);
            doc.setFontSize(12);
            doc.text("Summer Peak Calculation", 14, y);
            y += 6;
            const summerSub = generateSeasonalTable('Summer Loads', summerApps);
            y = checkPageBreak(doc, y, 40) + 2;
            const yrSub2 = generateSeasonalTable('Year-Round Loads', yearRoundApps);
            const summerTotal = summerSub + yrSub2;
            y = checkPageBreak(doc, y, 15) + 8;
            doc.setFontSize(10);
            doc.setFillColor('#e7f1ff');
            doc.rect(14, y - 4.5, pageWidth - 28, 12, 'F');
            doc.text(`Total Summer Peak = ${formatVal(summerSub)} (${'Summer'}) + ${formatVal(yrSub2)} (${'Year-Round'}) = `, 16, y);
            doc.setFont(undefined, 'bold');
            doc.text(`${formatVal(summerTotal)} ${displayUnit}`, doc.getTextWidth(`Total Summer Peak = ${formatVal(summerSub)} (Summer) + ${formatVal(yrSub2)} (Year-Round) = `) + 16, y);
            doc.setFont(undefined, 'normal');
            y += 15;


            // Load Summary
            y = checkPageBreak(doc, y, 60);
            doc.setFontSize(14);
            doc.text("Final Load Summary", 14, y);
            y += 6;
            autoTable(doc, {
                startY: y,
                body: [
                    [{ content: 'Total Connected Load', styles: { fontStyle: 'bold' } }, `${formatVal(convertFromBtu(calculatedResults.connectedBtu, displayUnit, gasEnergy))} ${displayUnit}`],
                    ['Winter Peak (Diversified)', `${formatVal(winterTotal)} ${displayUnit}`],
                    ['Summer Peak (Diversified)', `${formatVal(summerTotal)} ${displayUnit}`],
                    [{ content: 'Design Load (Worst Case)', styles: { fontStyle: 'bold', fillColor: '#d3e5ff' } }, { content: `${formatVal(convertFromBtu(calculatedResults.diversifiedBtu, displayUnit, gasEnergy))} ${displayUnit}`, styles: { fontStyle: 'bold', fillColor: '#d3e5ff' }}],
                ],
                theme: 'grid', styles: { fontSize: 10 }
            });
            // @ts-ignore
            y = doc.lastAutoTable.finalY + 10;

            // Methodology & Disclaimer
            y = checkPageBreak(doc, y, 150);
            const standardKey = diversityStandardSelect.value;
            const standardName = DIVERSITY_STANDARDS[standardKey] || 'the selected';
            doc.setFontSize(12);
            doc.text("Calculation Methodology & Sources", 14, y);
            y += 6;
            doc.setFontSize(10);
            const methText = `The Diversified Load for each appliance is calculated as: Max Input Rating × (Diversity Factor / 100). The Design Load is the greater of the Winter and Summer peak loads. The diversity factors used in this calculation are based on the ${standardName} standard. For code-compliant sizing, factors should be verified with the local utility or a qualified engineer. Gas energy content was assumed to be ${data.gasEnergyContent} BTU/ft³.`;
            doc.text(doc.splitTextToSize(methText, 180), 14, y);
            y += doc.splitTextToSize(methText, 180).length * 4 + 10;

            y = checkPageBreak(doc, y, 70);
            doc.setFontSize(12);
            doc.text("Important Disclaimer", 14, y);
            y += 6;
            doc.setFontSize(9);
            const disclaimerText = `This report is an estimation tool intended for sizing of UTILITY-OWNED ASSETS (service line, meter, regulator) only. The diversity factors used are standard engineering practice for this purpose. Sizing of CUSTOMER-OWNED interior gas piping is governed by NFPA 54 (National Fuel Gas Code), as adopted by the local Authority Having Jurisdiction (AHJ). NFPA 54 generally requires piping to be sized for the maximum connected load of all appliances. The diversity factors in this report MUST NOT be used for interior piping design unless explicitly permitted by the AHJ. All final designs must be verified by a qualified professional and approved by the local gas utility.`;
            const splitDisclaimer = doc.splitTextToSize(disclaimerText, 180);
            doc.setDrawColor(150);
            doc.rect(12, y - 2, 186, splitDisclaimer.length * 3.5 + 4);
            doc.setTextColor(80);
            doc.text(splitDisclaimer, 14, y);

            // Add Headers and Footers to all pages
            const pageCount = (doc as any).internal.getNumberOfPages();
            for(let i = 1; i <= pageCount; i++) {
                doc.setPage(i);
                addHeader(doc, data);
                doc.setFontSize(8);
                doc.setTextColor(150);
                doc.text(`Report generated on ${new Date().toLocaleDateString()}`, 14, pageHeight - 10);
                doc.text(`Page ${i} of ${pageCount}`, pageWidth - 35, pageHeight - 10);
            }
            
            doc.save(`${data.projectInfo.projectName || 'load-report'}.pdf`);
        } catch (error) {
            console.error("Failed to generate PDF:", error);
            showToast('There was an error generating the PDF.', 5000, 'error');
        } finally {
            btn.disabled = false;
            btn.classList.remove('loading');
            btn.textContent = originalText;
        }
    }, 50);
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

    // Accordion logic
    const accordionHeader = document.querySelector('#project-info .accordion-header');
    accordionHeader?.addEventListener('click', () => {
        const accordion = accordionHeader.closest('.accordion');
        if (accordion) {
            const isExpanded = accordion.classList.contains('open');
            accordionHeader.setAttribute('aria-expanded', String(!isExpanded));
            accordion.classList.toggle('open');
        }
    });

    // Quick Start guide dismiss logic
    const dismissBtn = document.getElementById('dismiss-guide-btn');
    const guide = document.getElementById('quick-start-guide');
    dismissBtn?.addEventListener('click', () => {
        if (guide) {
            guide.style.display = 'none';
        }
    });

    // Debounce function for performance on frequent inputs
    const debounce = (func: (...args: any[]) => void, delay: number) => {
        let timeoutId: number;
        return (...args: any[]) => {
            clearTimeout(timeoutId);
            timeoutId = window.setTimeout(() => func(...args), delay);
        };
    };
    const debouncedCalculate = debounce(calculate, 300);

    addApplianceBtn.addEventListener('click', () => {
        addApplianceRow();
        updateEmptyState();
        calculate();
    });
    
    applianceList.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        if (target.classList.contains('remove-btn')) {
            const rowToRemove = target.closest('.appliance-row');
            if(rowToRemove) {
                rowToRemove.classList.remove('show');
                rowToRemove.addEventListener('transitionend', () => {
                    rowToRemove.remove();
                    updateEmptyState();
                    calculate();
                }, { once: true });
            }
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
        calculate(); // Recalculate after any select/radio change
    });

    applianceList.addEventListener('input', (e) => {
        const target = e.target as HTMLElement;
        if (target.classList.contains('rating-input') || 
            target.classList.contains('diversity-input') || 
            target.classList.contains('custom-equipment-input')) 
        {
            debouncedCalculate();
        }
    });

    resultsUnitsSelect.addEventListener('change', () => {
        updateResultDisplay();
        updateCalculationDetails();
    });
    gasEnergyContentInput.addEventListener('input', debouncedCalculate);
    diversityStandardSelect.addEventListener('change', () => {
        updateAllDiversityFactors();
        calculate();
        updateMethodologyDisplay();
    });
    
    saveBtn.addEventListener('click', save);
    loadBtn.addEventListener('click', () => loadInput.click());
    loadInput.addEventListener('change', load);
    
    exportPdfBtn.addEventListener('click', exportToPdf);

    showDetailsBtn.addEventListener('click', () => {
        const isVisible = calculationDetailsEl.classList.contains('visible');
        calculationDetailsEl.classList.toggle('visible');
        showDetailsBtn.textContent = isVisible ? 'Show Calculation Details' : 'Hide Calculation Details';
        showDetailsBtn.setAttribute('aria-expanded', String(!isVisible));
    });
}

// --- INITIALIZATION ---
function init() {
    const today = new Date().toISOString().split('T')[0];
    (document.getElementById('assessment-date') as HTMLInputElement).value = today;
    (document.getElementById('revision-date') as HTMLInputElement).value = today;
    
    createDiversityTable();
    setupEventListeners();
    updateEmptyState();
    updateMethodologyDisplay();
}

init();