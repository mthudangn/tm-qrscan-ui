# Cloud Run API

Node.js backend for the manufacturing QR and barcode scanning system.

## Responsibilities

- Validate scan requests
- Route requests using `sid`
- Resolve PXK context
- Validate product keys
- Validate workflow steps
- Append accepted requests to the Google Sheets `QUEUE`
- Return structured OK, SKIP, or ERROR responses

## Local Development

```bash
npm install
npm start
```

## Deployment

```bash
gcloud run deploy tm-qrscan-api \
  --source . \
  --region asia-southeast1
```
