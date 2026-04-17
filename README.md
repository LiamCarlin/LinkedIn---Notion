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

## Follow-Up Automation (Pending Invites -> Messaging)

This repo also includes `follow_up_automation.py` for your connection follow-up workflow:

1. Reads contacts from Notion where status is a pending invite status.
2. Checks each LinkedIn URL to determine whether invite is accepted.
3. Marks accepted contacts in Notion as `Invite Accepted`.
4. Only after all checks are complete, prepares messages for accepted contacts from `message_template.txt` with `{first_name}` / `{name}` replacements.
5. Lets you quickly approve/reject each prepared message first (no sending yet).
6. Sends all approved messages in a second pass.
7. Marks sent contacts in Notion as `Initial Reachout Initiated`.

Default behavior processes all pending contacts:

```bash
python follow_up_automation.py
```

Run one specific profile only:

```bash
python follow_up_automation.py --only-url "https://www.linkedin.com/in/michael-ku-jr-512b70194/"
```

Dry run (no Notion writes):

```bash
python follow_up_automation.py --dry-run
```

Manual-send fallback mode:

```bash
python follow_up_automation.py --manual-send
```

Notes:

- Safari automation must be allowed on macOS (`Safari > Settings > Advanced > Show Develop menu`, then `Develop > Allow Remote Automation`).
- You should stay logged in to LinkedIn in Safari.
- If sending is too early/late for your machine, tune `.env`: `SAFARI_PROFILE_LOAD_DELAY_SEC` and `SAFARI_COMPOSE_LOAD_DELAY_SEC`.
- By default, Safari runs in background mode (no app activation/pop-up). Set `SAFARI_ACTIVATE_WINDOW=true` if you want visible Safari focus while debugging.

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
