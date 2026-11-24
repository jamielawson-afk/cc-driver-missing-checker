# Connected Couriers — Driver Missing Checker

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

## Master List storage
Master List records use SQLite. For Railway, mount a persistent volume at `/data` and set `MASTER_DB_PATH=/data/mentor.sqlite3` before adding drivers. Without persistent storage records may be lost on redeploy. Protect the application with authentication before storing staff information. Set a secure `SECRET_KEY` environment variable.

## Seeded register
59 supplied drivers are preloaded on the first run of a new SQLite database. Each record has first-name and last-name code; the **pair** must be unique, not each individual code. Existing database migrations from the earlier one-code prototype are not automatic.

## Office password
Before deployment set Railway variable `OFFICE_PASSWORD_HASH` to a Werkzeug password hash. Generate with `python -c "from werkzeug.security import generate_password_hash; import getpass; print(generate_password_hash(getpass.getpass('Office password: ')))"`. Keep password and hash out of GitHub. Existing `SECRET_KEY` must remain stable. Login protects both areas, with a one-hour session.

## Master List code pair search
Paste the first-name and last-name codes separated by a tab, space, or newline. Both codes must belong to the same driver. Search also supports single codes and names; a visible message names matching drivers or says No matching driver found.
