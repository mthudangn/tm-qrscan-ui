# Architecture

The system uses a three-layer architecture:

```text
GitHub Pages frontend
        ↓
Google Cloud Run API
        ↓
Google Sheets operational store
```

## Frontend Responsibilities

- Read master QR parameters
- Manage camera scanning
- Validate basic barcode format
- Submit requests
- Display OK, SKIP, or ERROR states

## Cloud Run Responsibilities

- Validate `sid`, `pxk`, `prodKey`, and `step`
- Open the correct spreadsheet
- Validate operational context
- Append accepted requests to `QUEUE`
- Return structured responses
- Keep credentials out of the browser

## Google Sheets Responsibilities

- Store PXK operational data
- Store queue requests
- Store logs
- Store catalogue or batch information

## Routing

- `sid` selects the spreadsheet
- `pxk` selects the operational context
- `prodKey` selects the production item
- `step` selects the manufacturing workflow stage
