# Connected Couriers — Mentor Checker

Standalone Flask app for comparing the Amazon Routes report against the daily driver list. It **compares names only**, ignores VINs, trims spaces, normalises case, and applies explicit known spelling variations. The upload files are processed in memory and are not retained. No database is required.

## Railway deployment

1. Push the repository contents to `main` on GitHub (files must be at the repository root).
2. On Railway, create a service using the repository and choose the `main` branch.
3. Railway/Nixpacks detects `requirements.txt` and `Procfile` and starts Gunicorn using the assigned `$PORT`.
4. Generate a public domain in Railway → Networking, then test `/health` and the form.

## Use

Upload an Amazon `.xlsx` or `.csv` report with a **Driver name** column. Upload a second `.xlsx`/`.csv` list or paste names (Firstname, Surname, VIN is accepted; VIN ignored). Click **Compare Drivers**. Download missing names as CSV if needed.

## Notes

The app is not authentication protected. Add access control before distributing the public URL beyond authorised office users. The default app stores no uploaded personal data; avoid adding logs containing names or file content.
