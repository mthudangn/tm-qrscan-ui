# Manufacturing QR and Barcode Scanning System

A cloud-based mobile scanning platform for manufacturing-floor operations, production tracking, and queue-based Google Sheets updates.

The system combines a mobile-first scanning interface hosted on GitHub Pages with a Node.js backend deployed on Google Cloud Run.

```text
Master QR Code
      ↓
Mobile Scanning UI
      ↓
Resolve Spreadsheet and PXK Context
      ↓
Scan Product Barcode
      ↓
Validate Manufacturing Workflow Step
      ↓
Submit Request to Cloud Run API
      ↓
Write Request to Google Sheets Queue
      ↓
Process Manufacturing Update
      ↓
Return OK, SKIP, or ERROR Status
```

## Overview

Manufacturing workers need a fast and reliable way to record production progress without manually editing operational spreadsheets. This system provides a mobile scanning workflow in which workers scan a master QR code, scan individual product barcodes, submit manufacturing-step updates through a central cloud API, receive immediate visual feedback, and allow spreadsheet mutations to be processed asynchronously through a queue.

The architecture separates the scanning UI, API layer, routing logic, and spreadsheet mutation workflow. This improves response time, maintainability, cross-device compatibility, multi-file scalability, failure recovery, auditability, and operational consistency.

## Core Architecture

```text
┌──────────────────────────────┐
│ Master QR Code               │
│ api + sid + pxk parameters   │
└──────────────┬───────────────┘
               ↓
┌──────────────────────────────┐
│ GitHub Pages Mobile UI       │
│ Camera and barcode scanner   │
└──────────────┬───────────────┘
               ↓ HTTPS
┌──────────────────────────────┐
│ Google Cloud Run API         │
│ Validation and routing       │
└──────────────┬───────────────┘
               ↓
┌──────────────────────────────┐
│ Google Sheets QUEUE          │
│ Asynchronous scan requests   │
└──────────────┬───────────────┘
               ↓
┌──────────────────────────────┐
│ Queue Processor              │
│ Workflow and PXK updates     │
└──────────────┬───────────────┘
               ↓
┌──────────────────────────────┐
│ Google Sheets PXK            │
│ Manufacturing status data    │
└──────────────────────────────┘
```

## Mobile Scanning UI

The frontend is designed for phones used directly on the manufacturing floor.

It supports:

- Mobile-first responsive design
- Camera-based QR and barcode scanning
- Rapid consecutive scans
- Cross-platform browser access
- Android and iOS compatibility
- GitHub Pages deployment
- No native app installation
- Clear operational feedback
- Minimal worker interaction

The UI uses three primary status states:

```text
Green → OK
Blue  → SKIP
Red   → ERROR
```

### OK

The request is valid and accepted. Examples include a valid product, a valid workflow step, a step that has not yet been completed, and a successfully created queue entry.

### SKIP

No new update is required. Examples include an already completed workflow step or a duplicate barcode event.

### ERROR

The request failed and may require retry or investigation. Examples include an invalid barcode, unknown product, invalid spreadsheet, missing PXK, API failure, queue failure, or spreadsheet access error.

## Master QR Routing

A master QR code opens the scanning interface with the correct operational context.

```text
https://mthudangn.github.io/tm-qrscan-ui/
?api=<CLOUD_RUN_API>
&mode=scan
&v=1
&sid=<SPREADSHEET_ID>
&pxk=<PXK_ID>
```

| Parameter | Purpose |
|---|---|
| `api` | Google Cloud Run API base URL |
| `mode` | Application mode |
| `v` | Interface or protocol version |
| `sid` | Target Google Spreadsheet ID |
| `pxk` | Active production or dispatch identifier |

## Multi-Spreadsheet Routing Through `sid`

The `sid` parameter identifies the Google Spreadsheet that should receive the scan request.

```text
sid = Google Spreadsheet ID
```

This enables one central UI, one Cloud Run API, multiple customer files, multiple monthly PXK files, and a standardised company-wide scanning workflow.

```text
One UI
+
One Cloud Run API
+
Many Google Sheets
```

## PXK Context Routing

The `pxk` parameter identifies the active production or dispatch context inside the selected spreadsheet.

Expected filename convention:

```text
PXK-<Customer>-<Month>-<Year>
```

Examples:

```text
PXK-NITORI-T4-2026
PXK-PHAT TRIEN-04-2026
```

Normalised identifiers:

```text
NITORI042026
PT042026
```

## Product Barcode Scanning

After opening the system through the master QR, workers scan individual product barcodes. The barcode represents a stable product key used consistently by printed labels, the mobile scanner, the Cloud Run API, Google Sheets lookup logic, and the queue processor.

Example:

```text
PRODUCT123
```

or:

```text
PRODUCT|ORDER
```

## Google Cloud Run API

The backend is implemented as a Node.js service deployed on Google Cloud Run.

Cloud Run provides managed HTTPS endpoints, stateless request processing, automatic scaling, centralised backend maintenance, remote access, and separation between frontend and spreadsheet logic.

Example request:

```json
{
  "sid": "TARGET_SPREADSHEET_ID",
  "pxk": "PT042026",
  "prodKey": "PRODUCT123",
  "step": "QC",
  "timestamp": "2026-08-01T10:30:00Z"
}
```

Example response:

```json
{
  "ok": true,
  "status": "queued",
  "message": "Scan accepted for processing"
}
```

## Queue-Based Google Sheets Updates

Accepted requests are appended to the `QUEUE` sheet instead of performing every spreadsheet mutation synchronously during the scan request.

```text
Mobile scan
    ↓
Cloud Run validation
    ↓
QUEUE row created
    ↓
Background processing
    ↓
PXK row updated
```

This improves scan response time, concurrent scanning, reliability, retry handling, failure visibility, auditability, and separation between request acceptance and spreadsheet mutation.

| Field | Description |
|---|---|
| Request ID | Unique request identifier |
| Timestamp | Time the request was accepted |
| Spreadsheet ID | Target spreadsheet |
| PXK ID | Active operational context |
| Product key | Scanned product |
| Step | Manufacturing workflow stage |
| Status | Pending, processing, completed, or failed |
| Error | Failure details |

## Manufacturing Workflow

The system supports production stages such as:

```text
GTAM
→ CAN
→ IN
→ CHAP
→ DAN
→ QC
→ KHO
```

The backend validates that the requested step exists, the product exists, the PXK exists, the step is not already complete, the request is valid for the selected spreadsheet, and the queue write succeeds.

## Google Sheets Structure

### `PXK`

Main production and dispatch table containing product information, customer information, order information, product keys, workflow dates, and the PXK identifier.

### `QUEUE`

Stores asynchronous requests, processing state, completion state, errors, and retry information.

### `LOG`

Stores scan events, validation errors, processing failures, and administrative actions.

### `CAT`

Stores batch, catalogue, or configuration information used by the workflow.

## Example Scan Flow

1. Worker scans the master QR.
2. UI resolves `api`, `sid`, and `pxk`.
3. Worker scans a product barcode.
4. Frontend submits the scan request.
5. Cloud Run validates the request.
6. A queue row is created.
7. The UI displays `OK`, `SKIP`, or `ERROR`.

Example frontend request:

```javascript
await fetch(`${apiBase}/scan`, {
  method: "POST",
  headers: {
    "Content-Type": "application/json"
  },
  body: JSON.stringify({
    sid,
    pxk,
    prodKey,
    step
  })
});
```

## Error Handling

Recommended error codes:

```text
INVALID_REQUEST
INVALID_SID
PXK_NOT_FOUND
PRODUCT_NOT_FOUND
INVALID_STEP
ALREADY_COMPLETED
QUEUE_WRITE_FAILED
SHEETS_API_ERROR
INTERNAL_ERROR
```

Example:

```json
{
  "ok": false,
  "status": "error",
  "code": "PRODUCT_NOT_FOUND",
  "message": "The scanned product was not found in the selected PXK."
}
```

## Performance Design

Key engineering decisions include a static GitHub Pages frontend, a managed Cloud Run backend, minimal request payloads, cached scanning context, queue-based spreadsheet mutation, immediate UI feedback, and background processing for accepted requests.

The target workflow supports rapid consecutive scans without requiring workers to wait for every spreadsheet update to finish.

## Security

Recommended controls:

- HTTPS-only communication
- Google service-account authentication
- Restricted spreadsheet permissions
- Input validation
- Allowed-origin configuration
- Request-size limits
- Environment variables for backend secrets
- Minimal service-account privileges
- Audit logging
- No secrets stored in frontend code

Never commit service-account JSON, API keys, access tokens, spreadsheet credentials, or production secrets.

## Deployment

### Frontend

Hosted on GitHub Pages:

```text
https://mthudangn.github.io/tm-qrscan-ui/
```

### Backend

Example Cloud Run deployment:

```bash
gcloud run deploy tm-qrscan-api \
  --source . \
  --region asia-southeast1 \
  --allow-unauthenticated
```

## Technology Stack

### Frontend

- HTML
- CSS
- JavaScript
- Camera API
- Barcode and QR decoding
- GitHub Pages

### Backend

- Node.js
- Express
- Google Cloud Run
- Docker
- Google Sheets API

### Data Layer

- Google Sheets
- PXK
- QUEUE
- LOG
- CAT

### Supporting Systems

- Google Apps Script
- Google Cloud Platform
- GitHub
- QR and barcode label generation

## Repository Structure

```text
tm-qrscan-barcode-cloud/
├── README.md
├── index.html
├── .gitignore
├── LICENSE
│
├── cloudrun-api/
│   ├── package.json
│   ├── server.js
│   ├── Dockerfile
│   ├── .dockerignore
│   ├── .env.example
│   └── README.md
│
└── docs/
    ├── ARCHITECTURE.md
    ├── API_REFERENCE.md
    ├── DEPLOYMENT.md
    ├── SPREADSHEET_SCHEMA.md
    └── SECURITY.md
```

## Engineering Contribution

The project replaces a tightly coupled spreadsheet-based workflow with a modular cloud architecture.

Key contributions include:

- One central mobile scanning interface
- Master QR context routing
- Product barcode scanning
- Multi-spreadsheet routing through `sid`
- PXK-level context through `pxk`
- Managed Cloud Run API
- Queue-based Google Sheets mutation
- Clear operational status handling
- Cross-device browser deployment
- Separation of frontend, API, and spreadsheet-processing responsibilities

## Project Status

Current capabilities include GitHub Pages mobile UI, master QR routing, product barcode scanning, Google Cloud Run integration, Google Sheets routing, queue-based updates, and manufacturing workflow status handling.

Ongoing work includes performance optimisation, retry handling, queue monitoring, operational logging, low-light scanning support, deployment automation, and production observability.

## Author

**Minnie Thu Dang**

AI/ML Engineer · Data Scientist · Cloud Application Developer

Engineering interests:

- Cloud-native applications
- Manufacturing automation
- Mobile web systems
- Google Cloud Platform
- Operational data systems
- Applied AI
- Scalable workflow automation
