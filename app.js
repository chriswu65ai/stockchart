const fileInput = document.getElementById('excel-file');
const sheetSelect = document.getElementById('sheet-select');
const resetZoomButton = document.getElementById('reset-zoom');
const statusText = document.getElementById('status');
const canvas = document.getElementById('share-chart');

let workbook = null;
let chart = null;

const DATE_KEYS = ['date', 'month', 'time'];
const PRICE_KEYS = ['price', 'share price', 'close', 'value'];
const EVENT_KEYS = ['event', 'title', 'milestone'];
const NOTE_KEYS = ['note', 'description', 'details', 'annotation'];

const normalize = (text) => String(text ?? '').trim().toLowerCase();

const getFirstMatchingKey = (headers, candidates) =>
  headers.find((header) => candidates.includes(normalize(header)));

const parseDate = (value) => {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
};

const updateStatus = (message, isError = false) => {
  statusText.textContent = message;
  statusText.style.color = isError ? '#991b1b' : 'inherit';
};

const clearChart = () => {
  if (chart) {
    chart.destroy();
    chart = null;
  }
  resetZoomButton.disabled = true;
};

const buildChart = (rows, columns) => {
  clearChart();

  const { dateKey, priceKey, eventKey, noteKey } = columns;

  const points = rows
    .map((row) => {
      const date = parseDate(row[dateKey]);
      const price = Number(row[priceKey]);

      if (!date || Number.isNaN(price)) {
        return null;
      }

      return {
        x: date,
        y: price,
        event: row[eventKey] ? String(row[eventKey]).trim() : '',
        note: row[noteKey] ? String(row[noteKey]).trim() : ''
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.x - b.x);

  if (!points.length) {
    updateStatus('No valid rows were found. Check date and price values in your file.', true);
    return;
  }

  const eventPoints = points.filter((point) => point.event || point.note);

  chart = new Chart(canvas, {
    type: 'line',
    data: {
      datasets: [
        {
          label: 'Share Price',
          data: points,
          borderColor: '#1d4ed8',
          backgroundColor: 'rgba(29, 78, 216, 0.15)',
          fill: true,
          tension: 0.2,
          pointRadius: 2,
          pointHoverRadius: 5
        },
        {
          type: 'bubble',
          label: 'Key Events',
          data: eventPoints.map((point) => ({
            x: point.x,
            y: point.y,
            r: 7,
            event: point.event,
            note: point.note
          })),
          backgroundColor: '#f59e0b',
          borderColor: '#b45309',
          borderWidth: 1,
          hoverBackgroundColor: '#f97316'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'nearest',
        intersect: false
      },
      scales: {
        x: {
          type: 'time',
          time: {
            unit: 'month'
          },
          title: {
            display: true,
            text: 'Date'
          }
        },
        y: {
          title: {
            display: true,
            text: 'Share Price'
          }
        }
      },
      plugins: {
        legend: {
          position: 'top'
        },
        tooltip: {
          callbacks: {
            label(context) {
              const base = `${context.dataset.label}: ${context.parsed.y}`;
              const event = context.raw?.event;
              const note = context.raw?.note;

              if (event && note) {
                return [base, `${event}: ${note}`];
              }
              if (event) {
                return [base, event];
              }
              if (note) {
                return [base, note];
              }

              return base;
            }
          }
        },
        zoom: {
          zoom: {
            wheel: {
              enabled: true
            },
            pinch: {
              enabled: true
            },
            mode: 'x'
          },
          pan: {
            enabled: true,
            mode: 'x'
          }
        }
      }
    }
  });

  resetZoomButton.disabled = false;
  updateStatus(`Rendered ${points.length} points with ${eventPoints.length} annotated events.`);
};

const parseSheet = (sheetName) => {
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  if (!rows.length) {
    updateStatus('Selected sheet has no data rows.', true);
    return;
  }

  const headers = Object.keys(rows[0]);
  const dateKey = getFirstMatchingKey(headers, DATE_KEYS);
  const priceKey = getFirstMatchingKey(headers, PRICE_KEYS);
  const eventKey = getFirstMatchingKey(headers, EVENT_KEYS);
  const noteKey = getFirstMatchingKey(headers, NOTE_KEYS);

  if (!dateKey || !priceKey) {
    updateStatus(
      'Unable to find required columns. Expected headers like Date and Price (or synonyms).',
      true
    );
    return;
  }

  buildChart(rows, {
    dateKey,
    priceKey,
    eventKey,
    noteKey
  });
};

fileInput.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  clearChart();

  if (!file) {
    updateStatus('No file selected.');
    sheetSelect.disabled = true;
    sheetSelect.innerHTML = '<option value="">Choose a sheet</option>';
    return;
  }

  try {
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data);

    sheetSelect.innerHTML = '<option value="">Choose a sheet</option>';
    workbook.SheetNames.forEach((name) => {
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      sheetSelect.append(option);
    });

    sheetSelect.disabled = false;
    updateStatus(`Loaded ${file.name}. Select a sheet to render the chart.`);
  } catch (error) {
    workbook = null;
    sheetSelect.disabled = true;
    updateStatus(`Could not read the Excel file: ${error.message}`, true);
  }
});

sheetSelect.addEventListener('change', (event) => {
  const sheetName = event.target.value;
  if (!sheetName || !workbook) {
    return;
  }

  parseSheet(sheetName);
});

resetZoomButton.addEventListener('click', () => {
  if (chart) {
    chart.resetZoom();
  }
});
