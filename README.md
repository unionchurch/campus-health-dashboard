# Campus Health Dashboard

This is a local custom-dashboard prototype built from:

`/Users/racheltates/Downloads/2026_Campuses Dashboard (2).xlsx`

## Run

From Terminal:

```bash
cd "/Users/racheltates/Documents/New project"
/Users/racheltates/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m http.server 4173 --bind localhost
```

Then open:

`http://localhost:4173`

## Refresh From A Downloaded Workbook

Update the workbook path in `scripts/generate_dashboard_data.py`, then run:

```bash
/Users/racheltates/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/generate_dashboard_data.py
```

That rebuilds `src/dashboard-data.js` from the Excel workbook.

The rebuild includes the weekly campus metrics plus the Big 5 raw-data model:

- Relationship Series
- Easter
- Welcome Home Sunday
- At The Movies
- Christmas

## Live Excel Online Path

The dashboard now has a browser-based Microsoft Graph connector in `src/live-excel.js`.

Configured Microsoft app:

- Client ID: `278dcce1-6a0a-4793-b0be-a444d5faed03`
- Tenant ID: `b98a50a2-9b8b-4900-81ad-a1c0d95cf816`
- Redirect URI: `http://localhost:4173`

The Entra app registration needs these delegated Microsoft Graph permissions:

- `User.Read`
- `Files.ReadWrite`

The dashboard reads the SharePoint workbook through Microsoft Graph and requests a non-persistent workbook session, so it does not save changes back to Excel. The static downloaded workbook data remains as the fallback if Microsoft sign-in or Graph access is not available.

In Microsoft Entra, also confirm:

- Platform is **Single-page application**
- Redirect URI includes `http://localhost:4173`
- The account signing in has access to the SharePoint workbook link
