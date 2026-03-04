# Interactive Share Price Chart

This project provides a browser-based share price chart that supports:

- Uploading Excel data (`.xlsx`/`.xls`)
- Zooming and panning through time
- A horizontal timeline scrollbar at the bottom of the chart
- Bubble annotations on key dates from the `Event` and/or `Note` columns

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
- `Note`

The app also accepts common synonyms, such as `Month`, `Share Price`, `Close`, `Title`, or `Milestone`.


Legend/series names come directly from your Excel header row.

If an Event cell contains an `http://` or `https://` link, clicking that bubble opens the URL in a new tab.

Use the "Show Event/Note annotations" checkbox to toggle only those annotation bubbles on/off.
