# Interactive Share Price Chart

This project provides a browser-based share price chart that supports:

- Uploading Excel data (`.xlsx`/`.xls`)
- Zooming and panning through time
- A single horizontal Timeframe bar at the bottom of the chart (drag to move, resize to widen/compress timeframe)
- Bubble annotations on key dates from the `Event` and/or `Note` columns
- Choose Series A (left axis) and optional Series B (right axis) from non-Date/non-Event/non-Comment columns

## Run locally

```bash
python3 -m http.server 8000
```

Open `http://localhost:8000` and upload your workbook.

## Expected Excel columns

Required:
- `Date`
- `Price`

Optional (for annotations):
- `Event`
- `Comment` (or `Note`)

Numeric series selection:
- `Series A` and `Series B` dropdowns list only numeric columns excluding Date/Event/Comment

The app also accepts common synonyms, such as `Month`, `Share Price`, `Close`, `Title`, or `Milestone`.


Legend/series names come directly from your Excel header row.

If an Event cell contains an `http://` or `https://` link, clicking that bubble opens the URL in a new tab.

Use separate "Show events" and "Show comments" checkboxes to independently hide/show each annotation series.

Event and Comment are rendered as separate annotation series (Event in amber, Comment in green), and empty series are not shown in the legend.

Annotation tooltip text wraps to constrain width to approximately 25% of chart width.

Touchscreen use: chart and timeline interactions lock page scrolling while dragging so iPad/touch gestures control the chart/timeline instead of moving the page.

Series style: choose Line or Bar independently for Series A and Series B.
