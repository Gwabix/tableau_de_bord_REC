grist.ready({ requiredAccess: 'full' });

const months = [
    'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
    'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'
];

let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();
let agendaDates = new Set();
let tableData = [];

function toDateStringSafe(value) {
    if (value instanceof Date) {
        return dateToString(value);
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
        const asMilliseconds = value > 100000000000 ? value : value * 1000;
        const parsed = new Date(asMilliseconds);
        return Number.isNaN(parsed.getTime()) ? null : dateToString(parsed);
    }

    if (typeof value === 'string') {
        const trimmed = value.trim();
        if (!trimmed) return null;

        const direct = new Date(trimmed);
        if (!Number.isNaN(direct.getTime())) {
            return dateToString(direct);
        }

        const isoDateMatch = /^(\d{4}-\d{2}-\d{2})/.exec(trimmed);
        return isoDateMatch ? isoDateMatch[1] : null;
    }

    return null;
}

const monthSelect = document.getElementById('monthSelect');
const yearSelect = document.getElementById('yearSelect');
const calendarGrid = document.getElementById('calendarGrid');
const prevMonthBtn = document.getElementById('prevMonthBtn');
const nextMonthBtn = document.getElementById('nextMonthBtn');
const prevMonthLabel = document.getElementById('prevMonthLabel');
const nextMonthLabel = document.getElementById('nextMonthLabel');

function updateMonthNavLabels() {
    const prevMonthIndex = (currentMonth + 11) % 12;
    const nextMonthIndex = (currentMonth + 1) % 12;
    const prevMonthName = months[prevMonthIndex];
    const nextMonthName = months[nextMonthIndex];

    if (prevMonthLabel) {
        prevMonthLabel.textContent = prevMonthName;
    }
    if (nextMonthLabel) {
        nextMonthLabel.textContent = nextMonthName;
    }
    if (prevMonthBtn) {
        prevMonthBtn.title = `Mois précédent (${prevMonthName})`;
        prevMonthBtn.setAttribute('aria-label', `Mois précédent (${prevMonthName})`);
    }
    if (nextMonthBtn) {
        nextMonthBtn.title = `Mois suivant (${nextMonthName})`;
        nextMonthBtn.setAttribute('aria-label', `Mois suivant (${nextMonthName})`);
    }
}

function syncSelects() {
    monthSelect.value = String(currentMonth);
    yearSelect.value = String(currentYear);
    updateMonthNavLabels();
}

function initSelects() {
    months.forEach((month, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = month;
        monthSelect.appendChild(option);
    });

    for (let year = currentYear - 5; year <= currentYear + 10; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearSelect.appendChild(option);
    }

    monthSelect.value = currentMonth;
    yearSelect.value = currentYear;
}

function dateToString(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function renderCalendar() {
    calendarGrid.innerHTML = '';
    updateMonthNavLabels();

    const firstDay = new Date(currentYear, currentMonth, 1);
    const lastDay = new Date(currentYear, currentMonth + 1, 0);
    const prevLastDay = new Date(currentYear, currentMonth, 0);

    let dayOfWeek = firstDay.getDay();
    dayOfWeek = dayOfWeek === 0 ? 7 : dayOfWeek;

    const prevMonthDays = dayOfWeek - 1;

    for (let i = prevMonthDays; i > 0; i--) {
        const day = prevLastDay.getDate() - i + 1;
        const cell = createDayCell(day, true);
        calendarGrid.appendChild(cell);
    }

    for (let day = 1; day <= lastDay.getDate(); day++) {
        const cell = createDayCell(day, false);
        calendarGrid.appendChild(cell);
    }

    const totalCells = calendarGrid.children.length;
    const remainingCells = (Math.ceil(totalCells / 7) * 7) - totalCells;

    for (let day = 1; day <= remainingCells; day++) {
        const cell = createDayCell(day, true);
        calendarGrid.appendChild(cell);
    }
}

function createDayCell(day, otherMonth) {
    const cell = document.createElement('div');
    cell.className = 'day-cell';
    cell.textContent = day;

    if (otherMonth) {
        cell.classList.add('other-month', 'empty');
        return cell;
    }

    const date = new Date(currentYear, currentMonth, day);
    const dateStr = dateToString(date);

    const today = new Date();
    if (date.toDateString() === today.toDateString()) {
        cell.classList.add('today');
    }

    if (agendaDates.has(dateStr)) {
        cell.classList.add('selected');
    }

    cell.addEventListener('click', () => handleDateClick(dateStr, cell));

    return cell;
}

async function handleDateClick(dateStr, cell) {
    if (cell.classList.contains('empty')) return;

    const isSelected = agendaDates.has(dateStr);

    if (isSelected) {
        const record = tableData.find(r => {
            return toDateStringSafe(r.Date) === dateStr;
        });

        if (record?.id !== undefined && record?.id !== null) {
            await grist.docApi.applyUserActions([
                ['RemoveRecord', 'Agenda', record.id]
            ]);
        }
    } else {
        const dateObj = new Date(dateStr + 'T12:00:00');
        const timestamp = dateObj.getTime() / 1000;

        await grist.docApi.applyUserActions([
            ['AddRecord', 'Agenda', null, { Date: timestamp }]
        ]);
    }
}

async function loadAgendaDates() {
    try {
        const rawData = await grist.docApi.fetchTable('Agenda');
        agendaDates.clear();

        const ids = Array.isArray(rawData?.id) ? rawData.id : [];
        const dateColumn = Array.isArray(rawData?.Date) ? rawData.Date : [];
        let titleColumn = [];
        if (Array.isArray(rawData?.['Intitulé'])) {
            titleColumn = rawData['Intitulé'];
        } else if (Array.isArray(rawData?.Intitule)) {
            titleColumn = rawData.Intitule;
        }

        dateColumn.forEach(date => {
            const dateStr = toDateStringSafe(date);
            if (dateStr) {
                agendaDates.add(dateStr);
            }
        });

        tableData = ids.map((id, index) => ({
            id: id,
            Date: dateColumn[index] ?? null,
            Intitulé: titleColumn[index] ?? ''
        }));

        renderCalendar();
    } catch (error) {
        console.error('Erreur lors du chargement des dates:', error);
    }
}

monthSelect.addEventListener('change', () => {
    currentMonth = Number.parseInt(monthSelect.value, 10);
    renderCalendar();
});

yearSelect.addEventListener('change', () => {
    currentYear = Number.parseInt(yearSelect.value, 10);
    renderCalendar();
});

prevMonthBtn?.addEventListener('click', () => {
    currentMonth -= 1;
    if (currentMonth < 0) {
        currentMonth = 11;
        currentYear -= 1;
    }
    syncSelects();
    renderCalendar();
});

nextMonthBtn?.addEventListener('click', () => {
    currentMonth += 1;
    if (currentMonth > 11) {
        currentMonth = 0;
        currentYear += 1;
    }
    syncSelects();
    renderCalendar();
});

grist.onRecords(() => {
    loadAgendaDates();
});

initSelects();
void loadAgendaDates();