# LUCAS Real App Platform

This branch keeps the current desktop/mobile bridge intact and starts the separate cloud-first platform.

## Target Model

The real app should have one hosted backend with accounts, workspaces, cloud photo storage, database-backed inventory/profit/expenses, and mobile clients that do not depend on a user's desktop computer being on.

Long-term flow:

1. User logs in on mobile.
2. User chooses a workspace.
3. Photos upload directly to the hosted LUCAS backend.
4. Backend stores photo bytes in object storage and metadata in the database.
5. OCR/linking workers process pending photos.
6. Inventory, sold history, expenses, and payouts update in the same backend.
7. Desktop LUCAS becomes a sync client, then eventually just another UI.

## First Platform Slice

The initial slice in `lucas_cloud/` provides:

- SQLite metadata store.
- File-backed object storage.
- Workspace creation with per-workspace API tokens.
- Mobile photo upload to a workspace.
- Pending photo inbox for desktop/worker processing.
- Photo status updates.
- Inventory item upsert primitives.
- Audit events.
- A small standard-library HTTP server for local development.

This code is intentionally isolated from `app.py` so the working Team/Personal mobile server remains unchanged.

## Local Development

Start the isolated cloud server:

```bash
LUCAS_CLOUD_ADMIN_TOKEN=dev-admin \
LUCAS_CLOUD_DATA_ROOT=work/lucas_cloud_data \
python3 -m lucas_cloud.server
```

Create a workspace:

```bash
curl -sS -X POST http://127.0.0.1:8780/api/v1/workspaces \
  -H 'Authorization: Bearer dev-admin' \
  -H 'Content-Type: application/json' \
  -d '{"slug":"danny","display_name":"Danny LUCAS"}'
```

The response includes a workspace token. Mobile uploads use that token:

```bash
curl -sS -X POST http://127.0.0.1:8780/api/v1/workspaces/danny/photos/upload \
  -H 'Authorization: Bearer <workspace-token>' \
  -H 'Content-Type: application/json' \
  -d '{"assigned_person":"Danny","images":[{"name":"front.jpg","image":"data:image/jpeg;base64,anBn"}]}'
```

List pending photos:

```bash
curl -sS http://127.0.0.1:8780/api/v1/workspaces/danny/photos/pending \
  -H 'Authorization: Bearer <workspace-token>'
```

## Migration Direction

Near-term, current LUCAS can keep using Google Drive JSON/XLSX files while this backend handles hosted photo upload/inbox.

Then we can add desktop sync:

- Pull pending hosted photos into desktop scan.
- Push desktop inventory/profit mutations into the backend.
- Gradually promote backend tables to the source of truth.

The production endpoint can later move from this standard-library server to FastAPI/Postgres/S3 without changing the workspace/photo lifecycle concepts.
