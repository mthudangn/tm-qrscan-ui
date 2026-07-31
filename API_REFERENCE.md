# API Reference

## Scan Endpoint

```http
POST /scan
Content-Type: application/json
```

## Request

```json
{
  "sid": "SPREADSHEET_ID",
  "pxk": "PT042026",
  "prodKey": "PRODUCT123",
  "step": "QC"
}
```

## Success

```json
{
  "ok": true,
  "status": "queued",
  "message": "Scan accepted for processing"
}
```

## Skip

```json
{
  "ok": true,
  "status": "skip",
  "code": "ALREADY_COMPLETED",
  "message": "The selected workflow step is already complete."
}
```

## Error

```json
{
  "ok": false,
  "status": "error",
  "code": "PRODUCT_NOT_FOUND",
  "message": "The scanned product was not found."
}
```
