const fileInput = document.getElementById('excel-file');
const sheetSelect = document.getElementById('sheet-select');
const seriesASelect = document.getElementById('series-a-select');
const seriesBSelect = document.getElementById('series-b-select');
const seriesABarToggle = document.getElementById('series-a-bar');
const seriesBBarToggle = document.getElementById('series-b-bar');
const resetZoomButton = document.getElementById('reset-zoom');
const showEventToggle = document.getElementById('show-event-annotations');
const showCommentToggle = document.getElementById('show-comment-annotations');
const statusText = document.getElementById('status');
const canvas = document.getElementById('share-chart');
const timelineWindow = document.getElementById('timeline-window');
const timelineSelection = document.getElementById('timeline-selection');
const timelineHandleLeft = document.getElementById('timeline-handle-left');
const timelineHandleRight = document.getElementById('timeline-handle-right');

let workbook = null;
let chart = null;
let chartSource = null;
let currentMeta = null;
let currentSheetContext = null;

let fullMinX = null;
let fullMaxX = null;
let viewSpan = null;

let windowStartPct = 0;
let windowSizePct = 100;
let isTimelineReady = false;

const DATE_KEYS = ['date', 'month', 'time'];
const EVENT_KEYS = ['event', 'title', 'milestone'];
const COMMENT_KEYS = ['comment', 'comments', 'note', 'notes'];
const URL_PATTERN = /(https?:\/\/[^\s]+)/i;
const MIN_WINDOW_PCT = 2;
const SERIES_A_COLOR = '#023047';
const SERIES_B_COLOR = '#22C4DD';

const extractHttpUrl = (text) => {
  const match = String(text ?? '').match(URL_PATTERN);
  if (!match) return '';

  try {
    const candidate = new URL(match[1]);
    if (candidate.protocol === 'http:' || candidate.protocol === 'https:') return candidate.href;
  } catch (_error) {
    return '';
  }

  return '';
};

const normalize = (text) => String(text ?? '').trim().toLowerCase();
const getFirstMatchingKey = (headers, candidates) =>
  headers.find((header) => candidates.includes(normalize(header)));

const parseDate = (value) => {
  if (value instanceof Date) return value;

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const parsedDate = new Date(value);
  return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
};

const parseNumeric = (value) => {
  if (typeof value === 'number') return value;
  if (typeof value === 'string' && value.trim() === '') return NaN;
  return Number(value);
};

const getWorksheetCellHyperlink = (worksheet, rowIndex, colIndex) => {
  if (colIndex < 0) return '';
  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = worksheet[cellAddress];
  const target = cell?.l?.Target || cell?.l?.target;
  return target ? String(target) : '';
};

const formatDateOnly = (value) =>
  new Intl.DateTimeFormat('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).format(
    new Date(value)
  );

const updateStatus = (message, isError = false) => {
  statusText.textContent = message;
  statusText.style.color = isError ? '#991b1b' : 'inherit';
};

const formatSeriesValue = (seriesKey, value) => {
  const fmt = currentMeta?.seriesFormats?.[seriesKey];
  if (!fmt) return value;

  try {
    return XLSX.SSF.format(fmt, value);
  } catch (_error) {
    return value;
  }
};

const wrapByPixelWidth = (text, chartInstance, maxWidthPx) => {
  const content = String(text ?? '');
  if (!content) return [''];

  const words = content.split(/\s+/);
  const lines = [];
  let line = '';
  const ctx = chartInstance.ctx;
  ctx.save();
  ctx.font = '12px sans-serif';

  words.forEach((word) => {
    const candidate = line ? `${line} ${word}` : word;
    if (ctx.measureText(candidate).width <= maxWidthPx) {
      line = candidate;
      return;
    }

    if (line) lines.push(line);

    if (ctx.measureText(word).width <= maxWidthPx) {
      line = word;
      return;
    }

    let segment = '';
    for (const char of word) {
      const next = `${segment}${char}`;
      if (ctx.measureText(next).width <= maxWidthPx) {
        segment = next;
      } else {
        if (segment) lines.push(segment);
        segment = char;
      }
    }
    line = segment;
  });

  if (line) lines.push(line);
  ctx.restore();
  return lines;
};

const makeSeriesDataset = ({ label, seriesKey, axisId, points, color, style }) => {
  const common = {
    label,
    seriesKey,
    yAxisID: axisId,
    data: points,
    borderColor: color
  };

  if (style === 'bar') {
    return {
      ...common,
      type: 'bar',
      backgroundColor: `${color}cc`,
      borderWidth: 1,
      barPercentage: 1.0,
      categoryPercentage: 0.9,
      maxBarThickness: 64
    };
  }

  return {
    ...common,
    type: 'line',
    backgroundColor: `${color}33`,
    fill: false,
    tension: 0.2,
    pointRadius: 2,
    pointHoverRadius: 5
  };
};

const buildVisibleDatasets = () => {
  if (!chartSource) return [];

  const datasets = [chartSource.seriesADataset];
  if (chartSource.seriesBDataset) datasets.push(chartSource.seriesBDataset);
  if (chartSource.eventDataset && showEventToggle.checked) datasets.push(chartSource.eventDataset);
  if (chartSource.commentDataset && showCommentToggle.checked) datasets.push(chartSource.commentDataset);
  return datasets;
};

const refreshAnnotationDatasets = () => {
  if (!chart || !chartSource) return;
  chart.data.datasets = buildVisibleDatasets();
  chart.update('none');
};

const resetTimelineWindow = () => {
  windowStartPct = 0;
  windowSizePct = 100;
  isTimelineReady = false;
  timelineWindow.classList.add('is-disabled');
  timelineSelection.style.left = '0%';
  timelineSelection.style.width = '100%';
};

const renderTimelineWindow = () => {
  timelineSelection.style.left = `${windowStartPct}%`;
  timelineSelection.style.width = `${windowSizePct}%`;
};

const syncWindowFromChart = () => {
  if (!chart || fullMinX === null || fullMaxX === null) {
    resetTimelineWindow();
    return;
  }

  const xScale = chart.scales.x;
  const visibleMin = xScale.min;
  const visibleMax = xScale.max;
  const fullSpan = fullMaxX - fullMinX;

  if (fullSpan <= 0) {
    resetTimelineWindow();
    return;
  }

  viewSpan = visibleMax - visibleMin;
  windowSizePct = Math.max(MIN_WINDOW_PCT, Math.min(100, (viewSpan / fullSpan) * 100));
  windowStartPct = Math.max(0, Math.min(100 - windowSizePct, ((visibleMin - fullMinX) / fullSpan) * 100));

  isTimelineReady = true;
  timelineWindow.classList.remove('is-disabled');
  renderTimelineWindow();
};

const applyWindowToChart = () => {
  if (!chart || fullMinX === null || fullMaxX === null) return;
  const fullSpan = fullMaxX - fullMinX;
  if (fullSpan <= 0) return;

  const nextMin = fullMinX + (windowStartPct / 100) * fullSpan;
  const nextMax = nextMin + (windowSizePct / 100) * fullSpan;

  chart.options.scales.x.min = nextMin;
  chart.options.scales.x.max = nextMax;
  chart.update('none');
  renderTimelineWindow();
};

const clearChart = () => {
  if (chart) {
    chart.destroy();
    chart = null;
  }

  chartSource = null;
  currentMeta = null;
  fullMinX = null;
  fullMaxX = null;
  viewSpan = null;
  resetZoomButton.disabled = true;
  resetTimelineWindow();
};

const resetSeriesSelectors = () => {
  seriesASelect.disabled = true;
  seriesBSelect.disabled = true;
  seriesABarToggle.disabled = true;
  seriesBBarToggle.disabled = true;
  seriesASelect.innerHTML = '<option value="">Choose series A</option>';
  seriesBSelect.innerHTML = '<option value="">None</option>';
  seriesABarToggle.checked = false;
  seriesBBarToggle.checked = false;
};

const syncSeriesSelectorOptions = () => {
  const selectedA = seriesASelect.value;
  const selectedB = seriesBSelect.value;

  Array.from(seriesASelect.options).forEach((option) => {
    option.disabled = Boolean(option.value && option.value === selectedB);
  });

  Array.from(seriesBSelect.options).forEach((option) => {
    option.disabled = Boolean(option.value && option.value === selectedA);
  });
};

const detectSeriesFormat = (worksheet, headers, seriesKey, rowCount) => {
  const colIndex = headers.indexOf(seriesKey);
  if (colIndex < 0) return '';

  for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = worksheet[cellAddress];
    if (cell && typeof cell.v === 'number' && cell.z) return cell.z;
  }

  return '';
};

const buildChart = (rows, columns) => {
  clearChart();

  const {
    dateKey,
    eventKey,
    commentKey,
    seriesAKey,
    seriesBKey,
    seriesAStyle,
    seriesBStyle,
    seriesFormats,
    worksheet,
    headers
  } = columns;
  currentMeta = { seriesFormats, seriesAKey, seriesBKey };

  const eventColIndex = eventKey ? headers.indexOf(eventKey) : -1;
  const commentColIndex = commentKey ? headers.indexOf(commentKey) : -1;

  const points = rows
    .map((row, dataIndex) => {
      const date = parseDate(row[dateKey]);
      const seriesAValue = parseNumeric(row[seriesAKey]);
      const seriesBValue = seriesBKey ? parseNumeric(row[seriesBKey]) : NaN;

      if (!date || Number.isNaN(seriesAValue)) return null;

      const event = eventKey && row[eventKey] ? String(row[eventKey]).trim() : '';
      const comment = commentKey && row[commentKey] ? String(row[commentKey]).trim() : '';

      // `sheet_to_json` strips cell hyperlink metadata; recover hyperlink targets directly from worksheet cells.
      const worksheetRowIndex = dataIndex + 1;
      const eventCellLink = getWorksheetCellHyperlink(worksheet, worksheetRowIndex, eventColIndex);
      const commentCellLink = getWorksheetCellHyperlink(
        worksheet,
        worksheetRowIndex,
        commentColIndex
      );

      return {
        x: date,
        seriesA: seriesAValue,
        seriesB: Number.isNaN(seriesBValue) ? null : seriesBValue,
        event,
        comment,
        eventLink: eventCellLink || extractHttpUrl(event),
        commentLink: commentCellLink || extractHttpUrl(comment)
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.x - b.x);

  if (!points.length) {
    updateStatus('No valid rows were found for the selected date and series columns.', true);
    return;
  }

  const eventPoints = points.filter((point) => point.event);
  const commentPoints = points.filter((point) => point.comment);

  chartSource = {
    seriesADataset: makeSeriesDataset({
      label: seriesAKey,
      seriesKey: seriesAKey,
      axisId: 'y',
      points: points.map((point) => ({ x: point.x, y: point.seriesA })),
      color: SERIES_A_COLOR,
      style: seriesAStyle
    }),
    seriesBDataset:
      seriesBKey && points.some((point) => point.seriesB !== null)
        ? makeSeriesDataset({
            label: seriesBKey,
            seriesKey: seriesBKey,
            axisId: 'y1',
            points: points.filter((point) => point.seriesB !== null).map((point) => ({ x: point.x, y: point.seriesB })),
            color: SERIES_B_COLOR,
            style: seriesBStyle
          })
        : null,
    eventDataset:
      eventKey && eventPoints.length
        ? {
            type: 'bubble',
            label: eventKey,
            data: eventPoints.map((point) => ({
              x: point.x,
              y: point.seriesA,
              r: 7,
              annotation: point.event,
              link: point.eventLink
            })),
            backgroundColor: '#f59e0b',
            borderColor: '#b45309',
            borderWidth: 1,
            hoverBackgroundColor: '#f97316'
          }
        : null,
    commentDataset:
      commentKey && commentPoints.length
        ? {
            type: 'bubble',
            label: 'Comment',
            data: commentPoints.map((point) => ({
              x: point.x,
              y: point.seriesA,
              r: 7,
              annotation: point.comment,
              link: point.commentLink
            })),
            backgroundColor: '#22c55e',
            borderColor: '#15803d',
            borderWidth: 1,
            hoverBackgroundColor: '#16a34a'
          }
        : null
  };

  chart = new Chart(canvas, {
    type: 'line',
    data: { datasets: buildVisibleDatasets() },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'nearest', intersect: false },
      onClick(event, _activeElements, chartInstance) {
        const hitElements = chartInstance.getElementsAtEventForMode(
          event,
          'nearest',
          { intersect: true },
          false
        );

        const bubbleHit = hitElements.find((item) => {
          const ds = chartInstance.data.datasets[item.datasetIndex];
          return ds?.type === 'bubble';
        });

        if (!bubbleHit) return;

        const dataset = chartInstance.data.datasets[bubbleHit.datasetIndex];
        const target = dataset?.data?.[bubbleHit.index];
        if (!target?.link) return;

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
          time: { unit: 'month' },
          title: { display: true, text: dateKey }
        },
        y: {
          position: 'left',
          title: { display: true, text: seriesAKey },
          ticks: {
            callback(value) {
              return formatSeriesValue(seriesAKey, value);
            }
          }
        },
        y1: {
          display: Boolean(seriesBKey && chartSource.seriesBDataset),
          position: 'right',
          grid: { drawOnChartArea: false },
          title: { display: true, text: seriesBKey || '' },
          ticks: {
            callback(value) {
              if (!seriesBKey) return value;
              return formatSeriesValue(seriesBKey, value);
            }
          }
        }
      },
      plugins: {
        legend: {
          position: 'top',
          labels: {
            filter(item, chartData) {
              const ds = chartData.datasets[item.datasetIndex];
              return ds?.type !== 'bubble';
            }
          }
        },
        tooltip: {
          filter(tooltipItem, _index, items) {
            const bubblePresent = items.some((item) => item.dataset.type === 'bubble');
            return bubblePresent ? tooltipItem.dataset.type === 'bubble' : true;
          },
          callbacks: {
            title(items) {
              if (!items.length) return '';
              return formatDateOnly(items[0].parsed.x);
            },
            label(context) {
              if (context.dataset.type === 'bubble') {
                const maxWidth = Math.max(120, context.chart.width * 0.25);
                return wrapByPixelWidth(context.raw?.annotation || '', context.chart, maxWidth);
              }

              const key = context.dataset.seriesKey;
              return `${context.dataset.label}: ${formatSeriesValue(key, context.parsed.y)}`;
            }
          }
        },
        zoom: {
          zoom: {
            wheel: { enabled: true },
            pinch: { enabled: true },
            mode: 'x',
            onZoomComplete: syncWindowFromChart
          },
          pan: { enabled: true, mode: 'x', onPanComplete: syncWindowFromChart }
        }
      }
    }
  });

  fullMinX = points[0].x.getTime();
  fullMaxX = points[points.length - 1].x.getTime();
  syncWindowFromChart();

  resetZoomButton.disabled = false;
  updateStatus(
    `Rendered ${points.length} points (${seriesAKey}${seriesBKey ? ` + ${seriesBKey}` : ''}) with ${eventPoints.length} events and ${commentPoints.length} comments.`
  );
};

const populateSeriesSelectors = () => {
  if (!currentSheetContext) {
    resetSeriesSelectors();
    return;
  }

  const options = currentSheetContext.seriesCandidates;
  if (!options.length) {
    resetSeriesSelectors();
    updateStatus('No numeric series columns found besides Date/Event/Comment.', true);
    return;
  }

  seriesASelect.innerHTML = '';
  options.forEach((key) => {
    const option = document.createElement('option');
    option.value = key;
    option.textContent = key;
    seriesASelect.append(option);
  });

  seriesBSelect.innerHTML = '<option value="">None</option>';
  options.forEach((key) => {
    const option = document.createElement('option');
    option.value = key;
    option.textContent = key;
    seriesBSelect.append(option);
  });

  seriesASelect.disabled = false;
  seriesBSelect.disabled = false;
  seriesABarToggle.disabled = false;
  seriesBBarToggle.disabled = false;

  seriesASelect.value = options[0];
  const defaultB = options.find((key) => key !== options[0]);
  seriesBSelect.value = defaultB || '';
  seriesABarToggle.checked = false;
  seriesBBarToggle.checked = false;
  syncSeriesSelectorOptions();
};

const renderSelectedSeries = () => {
  if (!currentSheetContext) return;

  const seriesAKey = seriesASelect.value;
  const seriesBKey = seriesBSelect.value;

  if (!seriesAKey) {
    updateStatus('Please select Series A to render the chart.', true);
    return;
  }

  if (seriesBKey && seriesBKey === seriesAKey) {
    updateStatus('Series A and Series B cannot be the same column.', true);
    return;
  }

  const headers = currentSheetContext.headers;
  const worksheet = currentSheetContext.worksheet;
  const rowCount = currentSheetContext.rows.length;

  const seriesFormats = {
    [seriesAKey]: detectSeriesFormat(worksheet, headers, seriesAKey, rowCount)
  };
  if (seriesBKey) {
    seriesFormats[seriesBKey] = detectSeriesFormat(worksheet, headers, seriesBKey, rowCount);
  }

  buildChart(currentSheetContext.rows, {
    dateKey: currentSheetContext.dateKey,
    eventKey: currentSheetContext.eventKey,
    commentKey: currentSheetContext.commentKey,
    worksheet,
    headers,
    seriesAKey,
    seriesBKey: seriesBKey || null,
    seriesAStyle: seriesABarToggle.checked ? 'bar' : 'line',
    seriesBStyle: seriesBBarToggle.checked ? 'bar' : 'line',
    seriesFormats
  });
};

const parseSheet = (sheetName) => {
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  if (!rows.length) {
    resetSeriesSelectors();
    updateStatus('Selected sheet has no data rows.', true);
    return;
  }

  const headers = Object.keys(rows[0]);
  const dateKey = getFirstMatchingKey(headers, DATE_KEYS);
  const eventKey = getFirstMatchingKey(headers, EVENT_KEYS);
  const commentKey = getFirstMatchingKey(headers, COMMENT_KEYS);

  if (!dateKey) {
    resetSeriesSelectors();
    updateStatus('Unable to find required date column.', true);
    return;
  }

  const seriesCandidates = headers.filter((key) => {
    if (key === dateKey || key === eventKey || key === commentKey) return false;

    return rows.some((row) => {
      const value = parseNumeric(row[key]);
      return !Number.isNaN(value);
    });
  });

  currentSheetContext = {
    worksheet,
    rows,
    headers,
    dateKey,
    eventKey,
    commentKey,
    seriesCandidates
  };

  populateSeriesSelectors();

  if (!seriesCandidates.length) {
    clearChart();
    return;
  }

  renderSelectedSeries();
};

fileInput.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  clearChart();
  currentSheetContext = null;
  resetSeriesSelectors();

  if (!file) {
    updateStatus('No file selected.');
    sheetSelect.disabled = true;
    sheetSelect.innerHTML = '<option value="">Choose a sheet</option>';
    return;
  }

  try {
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { cellNF: true });

    sheetSelect.innerHTML = '<option value="">Choose a sheet</option>';
    workbook.SheetNames.forEach((name) => {
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      sheetSelect.append(option);
    });

    if (!workbook.SheetNames.length) {
      sheetSelect.disabled = true;
      updateStatus(`Loaded ${file.name}, but no sheets were found.`, true);
      return;
    }

    sheetSelect.disabled = false;
    sheetSelect.value = '';
    updateStatus(`Loaded ${file.name}. Select a sheet, then choose Series A/Series B.`);
  } catch (error) {
    workbook = null;
    sheetSelect.disabled = true;
    updateStatus(`Could not read the Excel file: ${error.message}`, true);
  }
});

sheetSelect.addEventListener('change', (event) => {
  const sheetName = event.target.value;
  if (!sheetName || !workbook) return;
  parseSheet(sheetName);
});

seriesASelect.addEventListener('change', () => {
  if (!currentSheetContext) return;
  if (seriesBSelect.value && seriesBSelect.value === seriesASelect.value) {
    seriesBSelect.value = '';
  }
  syncSeriesSelectorOptions();
  renderSelectedSeries();
});

seriesBSelect.addEventListener('change', () => {
  if (!currentSheetContext) return;
  syncSeriesSelectorOptions();
  renderSelectedSeries();
});

seriesABarToggle.addEventListener('change', () => {
  if (!currentSheetContext) return;
  renderSelectedSeries();
});

seriesBBarToggle.addEventListener('change', () => {
  if (!currentSheetContext) return;
  renderSelectedSeries();
});

showEventToggle.addEventListener('change', refreshAnnotationDatasets);
showCommentToggle.addEventListener('change', refreshAnnotationDatasets);

const triggerResetZoom = (event) => {
  if (event) event.preventDefault();
  if (!chart || fullMinX === null || fullMaxX === null) return;

  // Reset both chart zoom range and the custom timeframe selector explicitly.
  windowStartPct = 0;
  windowSizePct = 100;
  renderTimelineWindow();

  chart.stop();
  chart.resetZoom();
  chart.options.scales.x.min = fullMinX;
  chart.options.scales.x.max = fullMaxX;

  // Force a full redraw (not a no-animation incremental update) to avoid stale canvas artifacts.
  chart.clear();
  chart.update();

  syncWindowFromChart();
};

resetZoomButton.addEventListener('click', triggerResetZoom);
resetZoomButton.addEventListener('touchend', triggerResetZoom, { passive: false });

const setupTimelineInteractions = () => {
  let dragMode = null;
  let startX = 0;
  let startWindow = { start: 0, size: 100 };

  const beginDrag = (mode, event) => {
    if (!isTimelineReady) return;
    dragMode = mode;
    startX = event.clientX;
    startWindow = { start: windowStartPct, size: windowSizePct };

    document.body.classList.add('is-timeline-dragging');

    const onMove = (moveEvent) => {
      if (!dragMode) return;
      moveEvent.preventDefault();
      const rect = timelineWindow.getBoundingClientRect();
      if (rect.width <= 0) return;

      const deltaPct = ((moveEvent.clientX - startX) / rect.width) * 100;

      if (dragMode === 'move') {
        windowStartPct = Math.max(0, Math.min(100 - startWindow.size, startWindow.start + deltaPct));
      } else if (dragMode === 'left') {
        const nextStart = Math.max(
          0,
          Math.min(startWindow.start + startWindow.size - MIN_WINDOW_PCT, startWindow.start + deltaPct)
        );
        const nextSize = startWindow.size + (startWindow.start - nextStart);
        windowStartPct = nextStart;
        windowSizePct = Math.max(MIN_WINDOW_PCT, Math.min(100, nextSize));
      } else if (dragMode === 'right') {
        const nextSize = Math.max(
          MIN_WINDOW_PCT,
          Math.min(100 - startWindow.start, startWindow.size + deltaPct)
        );
        windowSizePct = nextSize;
      }

      applyWindowToChart();
    };

    const onUp = () => {
      dragMode = null;
      document.body.classList.remove('is-timeline-dragging');
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };

    window.addEventListener('pointermove', onMove, { passive: false });
    window.addEventListener('pointerup', onUp);
  };

  timelineSelection.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    if (event.target === timelineHandleLeft || event.target === timelineHandleRight) return;
    timelineSelection.setPointerCapture?.(event.pointerId);
    beginDrag('move', event);
  });

  timelineHandleLeft.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    event.stopPropagation();
    timelineHandleLeft.setPointerCapture?.(event.pointerId);
    beginDrag('left', event);
  });

  timelineHandleRight.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    event.stopPropagation();
    timelineHandleRight.setPointerCapture?.(event.pointerId);
    beginDrag('right', event);
  });
};

resetSeriesSelectors();
setupTimelineInteractions();
