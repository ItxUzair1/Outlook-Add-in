# Backend (Milestone 2)

Local API bridge for filing emails and attachments from the Outlook add-in.

## Setup

1. `cd backend`
2. `npm install`
3. `copy .env.example .env`
4. `npm run dev`

## MSG strategies

- `MSG_STRATEGY=outlook-com`: tries strict `.msg` generation via Outlook COM automation (production path for Windows desktop agent scenarios).
- `MSG_STRATEGY=pseudo`: writes a structured placeholder payload to `.msg` extension.
- `STRICT_MSG_REQUIRED=true`: fail filing when strict generation is unavailable.

## Milestone 2 APIs

- `GET /api/health`
- `GET /api/locations`
- `POST /api/locations`
- `PUT /api/locations/:id`
- `DELETE /api/locations/:id`
- `GET /api/locations/suggested`
- `POST /api/file/email`
- `GET /api/search` (placeholder only for Milestone 2)
- `GET /api/preferences` (placeholder only for Milestone 2)

## Notes

- This implementation stores metadata in JSON files under `backend/data`.
- Filed emails/attachments are written under `backend/file-storage` by default.
- For production network drives, set location paths as absolute mapped paths like `H:\\Projects\\ProjectA\\Emails`.

## Dry run

- Run `npm run dry-run` from the `backend` folder (or `npm run backend:dry-run` from project root).
- The script creates/uses a dry-run location, files a sample email with attachment, and verifies files exist on disk.

## Manifest note

- Unified manifest (`manifest.json`) is kept for your current setup.
- For Outlook context-menu command surface compatibility, use `manifest.addin.xml`.
