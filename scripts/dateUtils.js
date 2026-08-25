/* ─────────────────────────────────────────
   dateUtils.js
   Builds the day-picker grid and auto-
   calculates D2…D7 from the D1 date.
───────────────────────────────────────── */

// ── Called whenever the day-count or D1 date changes ──
function buildDayGrid() {
  const count  = parseInt(document.getElementById('day-count').value);
  // const count = 6;
  const d1Val  = document.getElementById('d1-date').value;
  const grid   = document.getElementById('day-grid');
  grid.innerHTML = '';

  for (let i = 1; i <= count; i++) {
    grid.appendChild(createDayItem(i, d1Val));
  }
}

// ── Creates one day tile (label + date input) ──
function createDayItem(dayNumber, d1Val) {
  const wrapper = document.createElement('div');
  wrapper.className = 'day-item';

  const label = document.createElement('span');
  label.className = 'day-label';
  label.textContent = `D${dayNumber}`;

  const input = document.createElement('input');
  input.type     = 'date';
  input.id       = `day-${dayNumber}`;
  input.readOnly = dayNumber;   // Can only change prep day date, rest are calculated

  if (d1Val) {
    const date = new Date(d1Val + 'T00:00:00');  // Force local timezone
    date.setDate(date.getDate() + (dayNumber));
    input.value = date.toISOString().split('T')[0];
  }

  wrapper.appendChild(label);
  wrapper.appendChild(input);
  return wrapper;
}

// ── Returns a map of { D1: "April 29, 2025", D2: "April 30, 2025", … } ──
function buildDateMap() {
  const count   = parseInt(document.getElementById('day-count').value);
  const dateMap = {};

  for (let i = 1; i <= count; i++) {
    const inputEl = document.getElementById(`day-${i}`);
    if (!inputEl || !inputEl.value) continue;

    const date = new Date(inputEl.value + 'T00:00:00');
    dateMap[`D${i}`] = date.toLocaleDateString('en-US', {
      month: 'long',
      day:   'numeric',
      // year:  'numeric',
    });
  }

  return dateMap;
}

// ── Wire up change listeners once the DOM is ready ──
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('day-count').addEventListener('change', buildDayGrid);
  document.getElementById('d1-date').addEventListener('change', buildDayGrid);

  // Build the initial grid on page load
  buildDayGrid();
});
