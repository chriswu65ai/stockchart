const fileInput = document.getElementById('excel-file');
const sheetSelect = document.getElementById('sheet-select');
const seriesASelect = document.getElementById('series-a-select');
const seriesBSelect = document.getElementById('series-b-select');
const seriesCSelect = document.getElementById('series-c-select');
const seriesABarToggle = document.getElementById('series-a-bar');
const seriesBBarToggle = document.getElementById('series-b-bar');
const resetZoomButton = document.getElementById('reset-zoom');
const showEventToggle = document.getElementById('show-event-annotations');
const showCommentToggle = document.getElementById('show-comment-annotations');
const disableLinksToggle = document.getElementById('disable-links');
const statusText = document.getElementById('status');
const canvas = document.getElementById('share-chart');
const seriesCCanvas = document.getElementById('series-c-chart');
const seriesCChartContainer = document.getElementById('series-c-chart-container');
const timelineWindow = document.getElementById('timeline-window');
const timelineSelection = document.getElementById('timeline-selection');
const timelineHandleLeft = document.getElementById('timeline-handle-left');
const timelineHandleRight = document.getElementById('timeline-handle-right');
const quickTimeframeButtons = Array.from(document.querySelectorAll('.quick-timeframe-btn[data-years]'));
const seriesAAxisControls = document.getElementById('series-a-axis-controls');
const seriesBAxisControls = document.getElementById('series-b-axis-controls');
const seriesAMaxInput = document.getElementById('series-a-max');
const seriesAMinInput = document.getElementById('series-a-min');
const seriesAResetButton = document.getElementById('series-a-reset');
const seriesBMaxInput = document.getElementById('series-b-max');
const seriesBMinInput = document.getElementById('series-b-min');
const seriesBResetButton = document.getElementById('series-b-reset');
const seriesAInvertToggle = document.getElementById('series-a-invert');
const seriesBInvertToggle = document.getElementById('series-b-invert');
const seriesLeadLagControls = document.getElementById('series-leadlag-controls');
const seriesALeadLagInput = document.getElementById('series-a-leadlag');
const seriesBLeadLagInput = document.getElementById('series-b-leadlag');
const seriesLeadLagResetButton = document.getElementById('series-leadlag-reset');

let workbook = null;
let chart = null;
let seriesCChart = null;
let chartSource = null;
let currentMeta = null;
let currentSheetContext = null;

let fullMinX = null;
let fullMaxX = null;
let latestSeriesAMaxX = null;
let viewSpan = null;

let axisDefaults = {
  y: { min: null, max: null },
  y1: { min: null, max: null }
};
let axisOverrides = {
  y: { min: null, max: null },
  y1: { min: null, max: null }
};
let axisInversions = { y: false, y1: false };
let seriesLeadLagOffsets = { seriesA: 0, seriesB: 0 };

let windowStartPct = 0;
let windowSizePct = 100;
let isTimelineReady = false;

// Backward-compat globals: older cached bundles may still reference these names.
// Keeping them defined prevents runtime ReferenceError during mixed-cache rollouts.
let preservedMinX = null;
let preservedMaxX = null;

const DATE_KEYS = ['date', 'month', 'time'];
const EVENT_KEYS = ['event', 'title', 'milestone'];
const COMMENT_KEYS = ['comment', 'comments', 'note', 'notes'];
const URL_PATTERN = /(https?:\/\/[^\s]+)/i;
const MIN_WINDOW_PCT = 2;
const SERIES_A_COLOR = '#023047';
const SERIES_B_COLOR = '#22C4DD';
const SERIES_C_COLOR = '#7c3aed';

const extractHyperlinkFromFormula = (formula) => {
  if (typeof formula !== 'string') return '';

  const match = formula.match(/^\s*=\s*HYPERLINK\s*\(\s*"((?:""|[^"])*)"/i);
  if (!match) return '';

  const url = match[1].replace(/""/g, '"').trim();

  try {
    const candidate = new URL(url);
    if (candidate.protocol === 'http:' || candidate.protocol === 'https:') return candidate.href;
  } catch (_error) {
    return '';
  }

  return '';
};

const extractHttpUrl = (text) => {
  const formulaUrl = extractHyperlinkFromFormula(text);
  if (formulaUrl) return formulaUrl;

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

const areLinksDisabled = () => Boolean(disableLinksToggle?.checked);

const syncBubbleLinks = (chartInstance) => {
  if (!chartInstance) return;

  chartInstance.data.datasets.forEach((dataset) => {
    if (dataset?.type !== 'bubble' || !Array.isArray(dataset.data)) return;

    dataset.data.forEach((point) => {
      if (!point || typeof point !== 'object') return;
      const originalLink = point.rawLink ?? point.link ?? '';
      point.rawLink = originalLink;
      point.link = areLinksDisabled() ? '' : originalLink;
    });
  });
};

const normalize = (text) => String(text ?? '').trim().toLowerCase();
const getFirstMatchingKey = (headers, candidates) =>
  headers.find((header) => candidates.includes(normalize(header)));

const parseDate = (value) => {
  if (value instanceof Date) return value;

  const buildSafeDate = (year, month, day) => {
    const y = Number(year);
    const m = Number(month);
    const d = Number(day);
    if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;

    const dt = new Date(y, m - 1, d);
    if (Number.isNaN(dt.getTime())) return null;
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
    return dt;
  };

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return buildSafeDate(parsed.y, parsed.m, parsed.d);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;

    const split = trimmed.match(/^(\d{1,4})[-\/](\d{1,2})[-\/](\d{1,4})(?:\s.*)?$/);
    if (split) {
      const a = Number(split[1]);
      const b = Number(split[2]);
      const c = Number(split[3]);

      // YYYY-MM-DD / YYYY/MM/DD
      if (split[1].length === 4) {
        const iso = buildSafeDate(a, b, c);
        if (iso) return iso;
      }

      // DD/MM/YYYY or MM/DD/YYYY (and 2-digit year variants)
      const year = split[3].length === 2 ? Number(`20${split[3]}`) : c;
      const preferDMY = a > 12 || (a <= 12 && b <= 12);
      const first = preferDMY ? buildSafeDate(year, b, a) : buildSafeDate(year, a, b);
      if (first) return first;

      const second = preferDMY ? buildSafeDate(year, a, b) : buildSafeDate(year, b, a);
      if (second) return second;
    }
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
  if (target) return extractHttpUrl(target);

  const formulaLink = extractHyperlinkFromFormula(cell?.f);
  if (formulaLink) return formulaLink;

  return '';
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

const formatLeadLagLabelSuffix = (offset) => {
  if (!Number.isFinite(offset) || offset === 0) return '';
  return ` (${offset > 0 ? `+${offset}` : offset})`;
};

const shiftSeriesPointsByDates = (points, offset, dateIndexLookup, datesByIndex) => {
  if (!Array.isArray(points)) return [];
  if (!Number.isFinite(offset) || offset === 0) return points;

  return points
    .map((point) => {
      const index = dateIndexLookup.get(point.x.getTime());
      if (!Number.isInteger(index)) return null;

      const shiftedDate = datesByIndex[index + offset];
      if (!shiftedDate) return null;

      return { ...point, x: shiftedDate };
    })
    .filter(Boolean);
};

const rebuildShiftedSeriesDatasets = () => {
  if (!chartSource) return;

  const seriesAOffset = seriesLeadLagOffsets.seriesA;
  const seriesBOffset = seriesLeadLagOffsets.seriesB;

  chartSource.seriesADataset = makeSeriesDataset({
    label: `${chartSource.seriesA.key}${formatLeadLagLabelSuffix(seriesAOffset)}`,
    seriesKey: chartSource.seriesA.key,
    axisId: 'y',
    points: shiftSeriesPointsByDates(
      chartSource.seriesA.points,
      seriesAOffset,
      chartSource.dateIndexLookup,
      chartSource.datesByIndex
    ),
    color: SERIES_A_COLOR,
    style: chartSource.seriesA.style,
    order: 2
  });

  chartSource.seriesBDataset =
    chartSource.seriesB
      ? makeSeriesDataset({
          label: `${chartSource.seriesB.key}${formatLeadLagLabelSuffix(seriesBOffset)}`,
          seriesKey: chartSource.seriesB.key,
          axisId: 'y1',
          points: shiftSeriesPointsByDates(
            chartSource.seriesB.points,
            seriesBOffset,
            chartSource.dateIndexLookup,
            chartSource.datesByIndex
          ),
          color: SERIES_B_COLOR,
          style: chartSource.seriesB.style,
          order: 3
        })
      : null;
};

const makeSeriesDataset = ({ label, seriesKey, axisId, points, color, style, order }) => {
  const common = {
    label,
    seriesKey,
    yAxisID: axisId,
    data: points,
    borderColor: color,
    order
  };

  if (style === 'bar') {
    return {
      ...common,
      type: 'bar',
      backgroundColor: `${color}cc`,
      borderWidth: 1,
      barPercentage: 0.9,
      categoryPercentage: 0.8,
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

  const datasets = [];

  // Draw order (bottom -> top): Series B, Series A, annotations.
  // This keeps markers visible above both line/bar series.
  if (chartSource.seriesBDataset) datasets.push(chartSource.seriesBDataset);
  datasets.push(chartSource.seriesADataset);
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
  syncSeriesCChartRangeFromMain();
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
  syncSeriesCChartRangeFromMain();
};

const setQuickTimeframeButtonsDisabled = (disabled) => {
  quickTimeframeButtons.forEach((button) => {
    button.disabled = disabled;
  });
};

const applyLatestYearsWindow = (years) => {
  if (!chart || fullMinX === null || fullMaxX === null || !Number.isFinite(years) || years <= 0) return;

  const anchorMaxX = Number.isFinite(latestSeriesAMaxX) ? latestSeriesAMaxX : fullMaxX;
  const maxDate = new Date(anchorMaxX);
  const minDate = new Date(maxDate);
  minDate.setFullYear(maxDate.getFullYear() - years);

  const targetMin = Math.max(fullMinX, minDate.getTime());
  chart.options.scales.x.min = targetMin;
  chart.options.scales.x.max = anchorMaxX;
  chart.update('none');
  syncWindowFromChart();
  syncSeriesCChartRangeFromMain();
};

const getCurrentTimelineWindow = () => {
  if (!isTimelineReady) return null;
  if (!Number.isFinite(windowStartPct) || !Number.isFinite(windowSizePct)) return null;

  return { start: windowStartPct, size: windowSizePct };
};

const readAxisInputValue = (input) => {
  const raw = input.value.trim();
  if (!raw) return null;

  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
};

const formatAxisInputValue = (value) => (Number.isFinite(value) ? String(value) : '');

const updateAxisControlsUI = () => {
  const hasChart = Boolean(chart);
  const hasSeriesB = Boolean(chartSource?.seriesB);

  seriesAAxisControls.classList.toggle('is-disabled', !hasChart);
  seriesBAxisControls.classList.toggle('is-disabled', !hasChart || !hasSeriesB);
  seriesLeadLagControls.classList.toggle('is-disabled', !hasChart);

  seriesAMaxInput.disabled = !hasChart;
  seriesAMinInput.disabled = !hasChart;
  seriesAResetButton.disabled = !hasChart;
  seriesAInvertToggle.disabled = !hasChart;

  seriesBMaxInput.disabled = !hasChart || !hasSeriesB;
  seriesBMinInput.disabled = !hasChart || !hasSeriesB;
  seriesBResetButton.disabled = !hasChart || !hasSeriesB;
  seriesBInvertToggle.disabled = !hasChart || !hasSeriesB;

  seriesALeadLagInput.disabled = !hasChart;
  seriesBLeadLagInput.disabled = !hasChart || !hasSeriesB;
  seriesLeadLagResetButton.disabled = !hasChart;

  seriesAMaxInput.value = formatAxisInputValue(axisOverrides.y.max);
  seriesAMinInput.value = formatAxisInputValue(axisOverrides.y.min);

  seriesBMaxInput.value = formatAxisInputValue(axisOverrides.y1.max);
  seriesBMinInput.value = formatAxisInputValue(axisOverrides.y1.min);

  seriesAInvertToggle.checked = Boolean(axisInversions.y);
  seriesBInvertToggle.checked = Boolean(axisInversions.y1);

  seriesALeadLagInput.value = String(seriesLeadLagOffsets.seriesA || 0);
  seriesBLeadLagInput.value = String(seriesLeadLagOffsets.seriesB || 0);
};

const applyAxisOverrides = () => {
  if (!chart) return;

  const scales = chart.options.scales;
  const yMin = axisOverrides.y.min;
  const yMax = axisOverrides.y.max;
  const y1Min = axisOverrides.y1.min;
  const y1Max = axisOverrides.y1.max;

  scales.y.min = Number.isFinite(yMin) ? yMin : axisDefaults.y.min;
  scales.y.max = Number.isFinite(yMax) ? yMax : axisDefaults.y.max;
  scales.y1.min = Number.isFinite(y1Min) ? y1Min : axisDefaults.y1.min;
  scales.y1.max = Number.isFinite(y1Max) ? y1Max : axisDefaults.y1.max;
  scales.y.reverse = Boolean(axisInversions.y);
  scales.y1.reverse = Boolean(axisInversions.y1);

  chart.update();
};

const setAxisOverride = (axisId, bound, value) => {
  axisOverrides[axisId][bound] = value;
  applyAxisOverrides();
};

const resetAxisOverride = (axisId) => {
  axisOverrides[axisId].min = null;
  axisOverrides[axisId].max = null;
  axisInversions[axisId] = false;
  applyAxisOverrides();
  updateAxisControlsUI();
};

const readLeadLagOffset = (input) => {
  const raw = input.value.trim();
  if (!raw) return 0;

  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return 0;
  return Math.trunc(parsed);
};

const applyLeadLagOffsets = () => {
  if (!chart || !chartSource) return;
  rebuildShiftedSeriesDatasets();
  chart.data.datasets = buildVisibleDatasets();
  chart.update();
};

const clearChart = () => {
  if (chart) {
    chart.destroy();
    chart = null;
  }
  if (seriesCChart) {
    seriesCChart.destroy();
    seriesCChart = null;
  }

  seriesCChartContainer.classList.add('is-hidden');
  chartSource = null;
  currentMeta = null;
  fullMinX = null;
  fullMaxX = null;
  latestSeriesAMaxX = null;
  viewSpan = null;
  axisDefaults = { y: { min: null, max: null }, y1: { min: null, max: null } };
  axisOverrides = { y: { min: null, max: null }, y1: { min: null, max: null } };
  axisInversions = { y: false, y1: false };
  seriesLeadLagOffsets = { seriesA: 0, seriesB: 0 };
  resetZoomButton.disabled = true;
  setQuickTimeframeButtonsDisabled(true);
  resetTimelineWindow();
  updateAxisControlsUI();
};

const resetSeriesSelectors = () => {
  seriesASelect.disabled = true;
  seriesBSelect.disabled = true;
  seriesCSelect.disabled = true;
  seriesABarToggle.disabled = true;
  seriesBBarToggle.disabled = true;
  seriesASelect.innerHTML = '<option value="">Choose series A</option>';
  seriesBSelect.innerHTML = '<option value="">None</option>';
  seriesCSelect.innerHTML = '<option value="">None</option>';
  seriesABarToggle.checked = false;
  seriesBBarToggle.checked = true;
};

const syncSeriesSelectorOptions = () => {
  const selectedA = seriesASelect.value;
  const selectedB = seriesBSelect.value;
  const selectedC = seriesCSelect.value;

  Array.from(seriesASelect.options).forEach((option) => {
    option.disabled = Boolean(option.value && (option.value === selectedB || option.value === selectedC));
  });

  Array.from(seriesBSelect.options).forEach((option) => {
    option.disabled = Boolean(option.value && (option.value === selectedA || option.value === selectedC));
  });

  Array.from(seriesCSelect.options).forEach((option) => {
    option.disabled = Boolean(option.value && (option.value === selectedA || option.value === selectedB));
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

const syncSeriesCChartRangeFromMain = () => {
  if (!chart || !seriesCChart) return;
  seriesCChart.options.scales.x.min = chart.scales.x.min;
  seriesCChart.options.scales.x.max = chart.scales.x.max;
  seriesCChart.update('none');
};


const buildChart = (rows, columns) => {
  clearChart();

  const {
    dateKey,
    eventKey,
    commentKey,
    seriesAKey,
    seriesBKey,
    seriesCKey,
    seriesAStyle,
    seriesBStyle,
    seriesFormats,
    worksheet,
    headers
  } = columns;
  currentMeta = { seriesFormats, seriesAKey, seriesBKey, seriesCKey };

  const eventColIndex = eventKey ? headers.indexOf(eventKey) : -1;
  const commentColIndex = commentKey ? headers.indexOf(commentKey) : -1;

  const points = rows
    .map((row, dataIndex) => {
      const date = parseDate(row[dateKey]);
      const seriesAValue = parseNumeric(row[seriesAKey]);
      const seriesBValue = seriesBKey ? parseNumeric(row[seriesBKey]) : NaN;
      const seriesCValue = seriesCKey ? parseNumeric(row[seriesCKey]) : NaN;

      if (!date) return null;

      const normalizedSeriesA = Number.isNaN(seriesAValue) ? null : seriesAValue;
      const normalizedSeriesB = Number.isNaN(seriesBValue) ? null : seriesBValue;
      const normalizedSeriesC = Number.isNaN(seriesCValue) ? null : seriesCValue;
      if (normalizedSeriesA === null && normalizedSeriesB === null && normalizedSeriesC === null) return null;

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
        seriesA: normalizedSeriesA,
        seriesB: normalizedSeriesB,
        seriesC: normalizedSeriesC,
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

  const seriesAPoints = points.filter((point) => point.seriesA !== null);
  if (!seriesAPoints.length) {
    updateStatus('Series A has no valid data points for the selected columns.', true);
    return;
  }

  const eventPoints = points.filter((point) => point.event && point.seriesA !== null);
  const commentPoints = points.filter((point) => point.comment && point.seriesA !== null);

  const latestDateInFile = rows.reduce((latest, row) => {
    const rowDate = parseDate(row[dateKey]);
    if (!rowDate) return latest;
    const rowTime = rowDate.getTime();
    return Math.max(latest, rowTime);
  }, Number.NEGATIVE_INFINITY);

  const nextFullMinX = points[0].x.getTime();
  const nextSeriesAMaxX = seriesAPoints[seriesAPoints.length - 1].x.getTime();
  const nextFullMaxX = Number.isFinite(latestDateInFile)
    ? latestDateInFile
    : nextSeriesAMaxX;

  const datesByIndex = [...new Set(points.map((point) => point.x.getTime()))].map((time) => new Date(time));
  const dateIndexLookup = new Map(datesByIndex.map((date, index) => [date.getTime(), index]));

  chartSource = {
    datesByIndex,
    dateIndexLookup,
    seriesA: {
      key: seriesAKey,
      style: seriesAStyle,
      points: seriesAPoints.map((point) => ({ x: point.x, y: point.seriesA }))
    },
    seriesB:
      seriesBKey && points.some((point) => point.seriesB !== null)
        ? {
            key: seriesBKey,
            style: seriesBStyle,
            points: points.filter((point) => point.seriesB !== null).map((point) => ({ x: point.x, y: point.seriesB }))
          }
        : null,
    seriesC:
      seriesCKey && points.some((point) => point.seriesC !== null)
        ? {
            key: seriesCKey,
            points: points.filter((point) => point.seriesC !== null).map((point) => ({ x: point.x, y: point.seriesC }))
          }
        : null,
    seriesADataset: null,
    seriesBDataset: null,
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
              rawLink: point.eventLink,
              link: point.eventLink
            })),
            backgroundColor: '#f59e0b',
            borderColor: '#b45309',
            borderWidth: 1,
            hoverBackgroundColor: '#f97316',
            order: 1
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
              rawLink: point.commentLink,
              link: point.commentLink
            })),
            backgroundColor: '#22c55e',
            borderColor: '#15803d',
            borderWidth: 1,
            hoverBackgroundColor: '#16a34a',
            order: 1
          }
        : null
  };

  rebuildShiftedSeriesDatasets();

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

        if (!bubbleHit || areLinksDisabled()) return;

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
        canvas.style.cursor =
          dataset?.type === 'bubble' && target?.link && !disableLinksToggle.checked ? 'pointer' : 'default';
      },
      scales: {
        x: {
          type: 'time',
          time: { unit: 'month' },
          title: { display: true, text: dateKey },
          min: nextFullMinX,
          max: nextFullMaxX
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
          display: Boolean(seriesBKey && chartSource.seriesB),
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

  fullMinX = nextFullMinX;
  fullMaxX = nextFullMaxX;
  latestSeriesAMaxX = nextSeriesAMaxX;
  preservedMinX = nextFullMinX;
  preservedMaxX = nextFullMaxX;
  syncBubbleLinks(chart);

  if (chartSource.seriesC) {
    seriesCChartContainer.classList.remove('is-hidden');
    seriesCChart = new Chart(seriesCCanvas, {
      type: 'line',
      data: {
        datasets: [
          {
            type: 'line',
            label: chartSource.seriesC.key,
            data: chartSource.seriesC.points,
            seriesKey: chartSource.seriesC.key,
            yAxisID: 'y',
            borderColor: SERIES_C_COLOR,
            backgroundColor: `${SERIES_C_COLOR}33`,
            fill: false,
            tension: 0.2,
            pointRadius: 2,
            pointHoverRadius: 5
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'nearest', intersect: false },
        scales: {
          x: {
            type: 'time',
            time: { unit: 'month' },
            title: { display: true, text: dateKey },
            min: nextFullMinX,
            max: nextFullMaxX
          },
          y: {
            position: 'left',
            title: { display: true, text: chartSource.seriesC.key },
            ticks: {
              callback(value) {
                return formatSeriesValue(chartSource.seriesC.key, value);
              }
            }
          }
        },
        plugins: {
          legend: { position: 'top' },
          tooltip: {
            callbacks: {
              title(items) {
                if (!items.length) return '';
                return formatDateOnly(items[0].parsed.x);
              },
              label(context) {
                return `${context.dataset.label}: ${formatSeriesValue(chartSource.seriesC.key, context.parsed.y)}`;
              }
            }
          }
        }
      }
    });
  } else {
    seriesCChartContainer.classList.add('is-hidden');
  }

  const preservedTimeline = columns.preserveTimeline;
  if (preservedTimeline && Number.isFinite(preservedTimeline.start) && Number.isFinite(preservedTimeline.size)) {
    windowSizePct = Math.max(MIN_WINDOW_PCT, Math.min(100, preservedTimeline.size));
    windowStartPct = Math.max(0, Math.min(100 - windowSizePct, preservedTimeline.start));
    isTimelineReady = true;
    timelineWindow.classList.remove('is-disabled');
    applyWindowToChart();
  } else {
    syncWindowFromChart();
  }

  if (chart) {
    chart.clear();
    chart.update();

    axisDefaults = {
      y: { min: chart.scales.y.min, max: chart.scales.y.max },
      y1: { min: chart.scales.y1?.min ?? null, max: chart.scales.y1?.max ?? null }
    };
    axisOverrides = { y: { min: null, max: null }, y1: { min: null, max: null } };
    axisInversions = { y: false, y1: false };
    updateAxisControlsUI();
  }

  resetZoomButton.disabled = false;
  setQuickTimeframeButtonsDisabled(false);
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

  seriesASelect.innerHTML = '<option value="">None</option>';
  options.forEach((key) => {
    const option = document.createElement('option');
    option.value = key;
    option.textContent = key;
    seriesASelect.append(option);
  });

  seriesBSelect.innerHTML = '<option value="">None</option>';
  seriesCSelect.innerHTML = '<option value="">None</option>';
  options.forEach((key) => {
    const optionB = document.createElement('option');
    optionB.value = key;
    optionB.textContent = key;
    seriesBSelect.append(optionB);

    const optionC = document.createElement('option');
    optionC.value = key;
    optionC.textContent = key;
    seriesCSelect.append(optionC);
  });

  seriesASelect.disabled = false;
  seriesBSelect.disabled = false;
  seriesCSelect.disabled = false;
  seriesABarToggle.disabled = false;
  seriesBBarToggle.disabled = false;

  seriesASelect.value = '';
  seriesBSelect.value = '';
  seriesCSelect.value = '';
  seriesABarToggle.checked = false;
  seriesBBarToggle.checked = true;
  syncSeriesSelectorOptions();
};

const renderSelectedSeries = ({ preserveTimeline = true } = {}) => {
  if (!currentSheetContext) return;

  const seriesAKey = seriesASelect.value;
  const seriesBKey = seriesBSelect.value;
  const seriesCKey = seriesCSelect.value;

  if (!seriesAKey) {
    updateStatus('Please select Series A to render the chart.', true);
    return;
  }

  if (seriesBKey && seriesBKey === seriesAKey) {
    updateStatus('Series A and Series B cannot be the same column.', true);
    return;
  }

  if (seriesCKey && (seriesCKey === seriesAKey || seriesCKey === seriesBKey)) {
    updateStatus('Series C must be different from Series A and Series B.', true);
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
  if (seriesCKey) {
    seriesFormats[seriesCKey] = detectSeriesFormat(worksheet, headers, seriesCKey, rowCount);
  }

  const preservedTimelineWindow = preserveTimeline ? getCurrentTimelineWindow() : null;

  buildChart(currentSheetContext.rows, {
    dateKey: currentSheetContext.dateKey,
    eventKey: currentSheetContext.eventKey,
    commentKey: currentSheetContext.commentKey,
    worksheet,
    headers,
    seriesAKey,
    seriesBKey: seriesBKey || null,
    seriesCKey: seriesCKey || null,
    seriesAStyle: seriesABarToggle.checked ? 'bar' : 'line',
    seriesBStyle: seriesBBarToggle.checked ? 'bar' : 'line',
    seriesFormats,
    preserveTimeline: preservedTimelineWindow
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

  clearChart();
  updateStatus('Sheet parsed. Select Series A/B/C to render chart(s).');
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
    const [firstSheet] = workbook.SheetNames;
    sheetSelect.value = firstSheet || '';

    if (firstSheet) {
      try {
        parseSheet(firstSheet);
        updateStatus(`Loaded ${file.name}. Showing ${firstSheet}. You can change Sheet/Series selections anytime.`);
      } catch (parseError) {
        updateStatus(`Loaded ${file.name}, but failed to render ${firstSheet}: ${parseError.message}`, true);
      }
    } else {
      updateStatus(`Loaded ${file.name}. Select a sheet, then choose Series A/Series B.`);
    }
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
  if (seriesCSelect.value && seriesCSelect.value === seriesASelect.value) {
    seriesCSelect.value = '';
  }
  syncSeriesSelectorOptions();
  renderSelectedSeries();
});

seriesBSelect.addEventListener('change', () => {
  if (!currentSheetContext) return;
  if (seriesCSelect.value && seriesCSelect.value === seriesBSelect.value) {
    seriesCSelect.value = '';
  }
  syncSeriesSelectorOptions();
  renderSelectedSeries();
});

seriesCSelect.addEventListener('change', () => {
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
disableLinksToggle.addEventListener('change', () => {
  syncBubbleLinks(chart);
  if (chart) chart.update('none');
  canvas.style.cursor = 'default';
});

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
  syncSeriesCChartRangeFromMain();
};

resetZoomButton.addEventListener('click', triggerResetZoom);
resetZoomButton.addEventListener('touchend', triggerResetZoom, { passive: false });

quickTimeframeButtons.forEach((button) => {
  button.addEventListener('click', () => {
    const years = Number(button.dataset.years);
    applyLatestYearsWindow(years);
  });
});

seriesAMaxInput.addEventListener('input', () => {
  if (!chart) return;
  setAxisOverride('y', 'max', readAxisInputValue(seriesAMaxInput));
});

seriesAMinInput.addEventListener('input', () => {
  if (!chart) return;
  setAxisOverride('y', 'min', readAxisInputValue(seriesAMinInput));
});

seriesBMaxInput.addEventListener('input', () => {
  if (!chart || !chartSource?.seriesB) return;
  setAxisOverride('y1', 'max', readAxisInputValue(seriesBMaxInput));
});

seriesBMinInput.addEventListener('input', () => {
  if (!chart || !chartSource?.seriesB) return;
  setAxisOverride('y1', 'min', readAxisInputValue(seriesBMinInput));
});

seriesAResetButton.addEventListener('click', () => {
  if (!chart) return;
  resetAxisOverride('y');
});

seriesBResetButton.addEventListener('click', () => {
  if (!chart || !chartSource?.seriesB) return;
  resetAxisOverride('y1');
});

seriesAInvertToggle.addEventListener('change', () => {
  if (!chart) return;
  axisInversions.y = seriesAInvertToggle.checked;
  applyAxisOverrides();
});

seriesBInvertToggle.addEventListener('change', () => {
  if (!chart || !chartSource?.seriesB) return;
  axisInversions.y1 = seriesBInvertToggle.checked;
  applyAxisOverrides();
});

seriesALeadLagInput.addEventListener('input', () => {
  if (!chart) return;
  seriesLeadLagOffsets.seriesA = readLeadLagOffset(seriesALeadLagInput);
  applyLeadLagOffsets();
});

seriesBLeadLagInput.addEventListener('input', () => {
  if (!chart || !chartSource?.seriesB) return;
  seriesLeadLagOffsets.seriesB = readLeadLagOffset(seriesBLeadLagInput);
  applyLeadLagOffsets();
});

seriesLeadLagResetButton.addEventListener('click', () => {
  if (!chart) return;
  seriesLeadLagOffsets = { seriesA: 0, seriesB: 0 };
  applyLeadLagOffsets();
  updateAxisControlsUI();
});

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
updateAxisControlsUI();
setupTimelineInteractions();
