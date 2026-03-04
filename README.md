# Interactive Share Price Chart

This project provides a browser-based share price chart that supports:

- Uploading Excel data (`.xlsx`/`.xls`)
- Zooming and panning through time
- Bubble annotations on key dates from event/note columns

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

The app also accepts common synonyms, such as `Month`, `Share Price`, `Close`, `Title`, or `Description`.
