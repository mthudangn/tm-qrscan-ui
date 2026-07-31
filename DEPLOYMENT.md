# Deployment

## Frontend

Deploy `index.html` and frontend assets through GitHub Pages.

## Backend

```bash
gcloud run deploy tm-qrscan-api \
  --source ./cloudrun-api \
  --region asia-southeast1
```

## Environment Variables

Use Cloud Run environment variables or Secret Manager for credentials, allowed spreadsheet IDs, allowed origins, and logging configuration.

Never store production secrets in frontend code or commit them to GitHub.
