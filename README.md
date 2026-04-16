# LinkedIn → Notion Contact Importer

Paste a LinkedIn profile URL, and this tool creates a contact in your Notion Contacts database with:

- Name
- Company
- Email (if available from enrichment)
- LinkedIn URL
- Role
- Status
- Last Connected (optional auto-set to today)

Before writing to Notion, the CLI shows each field and lets you confirm/edit it.

This version does **not** require Proxycurl.
It also supports saved LinkedIn HTML/MHTML exports for better accuracy when LinkedIn blocks live requests.

## Important

You do **not** need to share your Notion login with anyone.
Use a Notion Integration token instead.

## 1) Create a Notion integration

1. Go to Notion integrations: https://www.notion.so/my-integrations
2. Create a new internal integration.
3. Copy the integration token.
4. In your Contacts database page, click **Share** and invite that integration.

## 2) Get your Contacts database ID

Open your contacts database as a full page. The URL includes a long ID.
Use only the 32-character database ID for `NOTION_CONTACTS_DATABASE_ID` (not the `?v=...` view parameter).

## 3) Install and configure

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Then edit `.env` with your real values.

## 4) Run

Interactive mode:

```bash
python app.py
```

One URL mode:

```bash
python app.py "https://www.linkedin.com/in/someone/"
```

Saved export mode:

```bash
python app.py "/Users/you/Downloads/Someone LinkedIn.mhtml"
```

## Notes about your Notion schema

This script auto-detects the property types in your Contacts database.

- If `Company` is a relation property, it will find/create the company in the related Companies database and link it.
- If your integration does not have access to the related Companies database, contact creation still works (relation is skipped).
- If `LinkedIn` is `url`, it writes a URL; if it is text, it writes text.
- `Status` works with both `status` and `select` property types.
- If `Last Connected` is a `date` property and `AUTO_SET_LAST_CONNECTED=true`, it sets it to today.

## Limitations

- LinkedIn often hides profile data for unauthenticated requests.
- LinkedIn may return HTTP `999` (blocked automation). On macOS, the script now tries to auto-capture a local HTML export via Safari and continue automatically.
- Auto-capture works best when you are already logged into LinkedIn in Safari.
- Saved HTML/MHTML exports usually work better than live URLs.
- `Email` is usually unavailable from public profile metadata.
- `Company` / `Role` may be partial depending on profile visibility.
- Respect LinkedIn terms of service.
