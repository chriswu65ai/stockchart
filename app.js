const fileInput = document.getElementById('excel-file');
const sheetSelect = document.getElementById('sheet-select');
const resetZoomButton = document.getElementById('reset-zoom');
const chartScroll = document.getElementById('chart-scroll');
const annotationToggle = document.getElementById('toggle-annotations');
const statusText = document.getElementById('status');
const canvas = document.getElementById('share-chart');

let workbook = null;
let chart = null;

let fullMinX = null;
let fullMaxX = null;
let viewSpan = null;

const DATE_KEYS = ['date', 'month', 'time'];
const PRICE_KEYS = ['price', 'share price', 'close', 'value'];
const EVENT_KEYS = ['event', 'title', 'milestone'];
const NOTE_KEYS = ['note', 'notes'];

const URL_PATTERN = /(https?:\/\/[^\s]+)/i;

const extractHttpUrl = (text) => {
  const match = String(text ?? '').match(URL_PATTERN);
  if (!match) {
    return '';
  }

  try {
    const candidate = new URL(match[1]);
    if (candidate.protocol === 'http:' || candidate.protocol === 'https:') {
      return candidate.href;
    }
  } catch (_error) {
    return '';
  }

  return '';
};

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

const formatDateOnly = (value) =>
  new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }).format(new Date(value));

const updateStatus = (message, isError = false) => {
  statusText.textContent = message;
  statusText.style.color = isError ? '#991b1b' : 'inherit';
};

const resetScrollbar = () => {
  fullMinX = null;
  fullMaxX = null;
  viewSpan = null;
  chartScroll.value = '0';
  chartScroll.min = '0';
  chartScroll.max = '0';
  chartScroll.disabled = true;
};

const syncScrollbarFromChart = () => {
  if (!chart || fullMinX === null || fullMaxX === null) {
    resetScrollbar();
    return;
  }

  const xScale = chart.scales.x;
  const visibleMin = xScale.min;
  const visibleMax = xScale.max;

  viewSpan = visibleMax - visibleMin;

  const fullSpan = fullMaxX - fullMinX;
  const maxOffset = Math.max(0, fullSpan - viewSpan);
  const offset = Math.min(Math.max(0, visibleMin - fullMinX), maxOffset);

  chartScroll.min = '0';
  chartScroll.max = String(Math.round(maxOffset));
  chartScroll.value = String(Math.round(offset));
  chartScroll.step = String(Math.max(1, Math.round(fullSpan / 1000)));
  chartScroll.disabled = maxOffset === 0;
};

const clearChart = () => {
  if (chart) {
    chart.destroy();
    chart = null;
  }
  resetZoomButton.disabled = true;
  resetScrollbar();
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

      const event = eventKey && row[eventKey] ? String(row[eventKey]).trim() : '';
      const note = noteKey && row[noteKey] ? String(row[noteKey]).trim() : '';
      const annotation = [event, note].filter(Boolean).join(' | ');

      return {
        x: date,
        y: price,
        annotation,
        link: extractHttpUrl(annotation)
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.x - b.x);

  if (!points.length) {
    updateStatus('No valid rows were found. Check date and price values in your file.', true);
    return;
  }

  const annotationPoints = points.filter((point) => point.annotation);
  const datasets = [
    {
      label: priceKey,
      data: points,
      borderColor: '#1d4ed8',
      backgroundColor: 'rgba(29, 78, 216, 0.15)',
      fill: true,
      tension: 0.2,
      pointRadius: 2,
      pointHoverRadius: 5
    }
  ];

  if (eventKey || noteKey) {
    const annotationLabel = [eventKey, noteKey].filter(Boolean).join(' / ');

    datasets.push({
      type: 'bubble',
      label: annotationLabel,
      annotationControlled: true,
      hidden: !annotationToggle.checked,
      data: annotationPoints.map((point) => ({
        x: point.x,
        y: point.y,
        r: 7,
        annotation: point.annotation,
        link: point.link
      })),
      backgroundColor: '#f59e0b',
      borderColor: '#b45309',
      borderWidth: 1,
      hoverBackgroundColor: '#f97316'
    });
  }

  chart = new Chart(canvas, {
    type: 'line',
    data: {
      datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'nearest',
        intersect: false
      },
      onClick(_event, activeElements, chartInstance) {
        if (!activeElements.length) {
          return;
        }

        const { datasetIndex, index } = activeElements[0];
        const dataset = chartInstance.data.datasets[datasetIndex];

        if (dataset.type !== 'bubble') {
          return;
        }

        const target = dataset.data[index];
        if (!target?.link) {
          return;
        }

        window.open(target.link, '_blank', 'noopener,noreferrer');
      },
      onHover(_event, activeElements, chartInstance) {
        if (!activeElements.length) {
          canvas.style.cursor = 'default';
          return;
        }

        const { datasetIndex, index } = activeElements[0];
        const dataset = chartInstance.data.datasets[datasetIndex];
        const target = dataset?.data?.[index];

        canvas.style.cursor = dataset?.type === 'bubble' && target?.link ? 'pointer' : 'default';
      },
      scales: {
        x: {
          type: 'time',
          time: {
            unit: 'month'
          },
          title: {
            display: true,
            text: dateKey
          }
        },
        y: {
          title: {
            display: true,
            text: priceKey
          }
        }
      },
      plugins: {
        legend: {
          position: 'top'
        },
        tooltip: {
          callbacks: {
            title(items) {
              if (!items.length) return '';
              return formatDateOnly(items[0].parsed.x);
            },
            label(context) {
              if (context.dataset.type === 'bubble') {
                return context.raw?.annotation || 'Annotation';
              }

              return `${context.dataset.label}: ${context.parsed.y}`;
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
            mode: 'x',
            onZoomComplete: syncScrollbarFromChart
          },
          pan: {
            enabled: true,
            mode: 'x',
            onPanComplete: syncScrollbarFromChart
          }
        }
      }
    }
  });

  fullMinX = points[0].x.getTime();
  fullMaxX = points[points.length - 1].x.getTime();
  syncScrollbarFromChart();

  resetZoomButton.disabled = false;
  updateStatus(`Rendered ${points.length} points with ${annotationPoints.length} annotated events.`);
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

chartScroll.addEventListener('input', (event) => {
  if (!chart || fullMinX === null || fullMaxX === null || viewSpan === null) {
    return;
  }

  const fullSpan = fullMaxX - fullMinX;
  const maxOffset = Math.max(0, fullSpan - viewSpan);
  const offset = Math.min(Math.max(0, Number(event.target.value)), maxOffset);

  chart.options.scales.x.min = fullMinX + offset;
  chart.options.scales.x.max = fullMinX + offset + viewSpan;
  chart.update('none');
});

annotationToggle.addEventListener('change', (event) => {
  if (!chart) {
    return;
  }

  chart.data.datasets.forEach((dataset) => {
    if (dataset.annotationControlled) {
      dataset.hidden = !event.target.checked;
    }
  });

  chart.update('none');
});

resetZoomButton.addEventListener('click', () => {
  if (chart) {
    chart.resetZoom();
    syncScrollbarFromChart();
  }
});
