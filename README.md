# Wialon Auditor

Pull your entire Wialon fleet config into a Google Sheet. One script, one file, copy-paste and go.

If you've ever tried to audit hundreds of units in Wialon -- checking hardware params, sensor setups, eco-driving rules, custom fields -- you know it's painful through the web UI. This tool dumps everything into a spreadsheet so you can actually sort, filter, compare, and share it.

Tested on fleets with 10,000+ units.

![Image](https://github.com/user-attachments/assets/db1d8256-ce30-416b-998e-5909b4d3a357)

## What it fetches

| Sheet | Data |
|---|---|
| `units` | Properties, IMEI, phone, counters, trip detector, fuel settings |
| `hardware` | Device parameters (key-value pairs), hardware type info |
| `sensors` | Names, types, params, calibration tables, validation |
| `commands` | Configured commands, link types, parameters |
| `profiles` | VIN, plate, brand, model, engine, dimensions |
| `custom_fields` | User and admin custom fields |
| `drive_rank` | Eco-driving criteria (accel, braking, turning, speeding, etc.) |

Plus a `QUERYUNIT` formula you can use in any cell:

```
=QUERYUNIT("My Vehicle")
=QUERYUNIT("My Vehicle", "sensors")
```

And a Filter sheet that shows all data for a selected unit in one view.

## Setup

You need: a Google account, a Wialon account with API access, and 5 minutes.

**1. Create the spreadsheet**

1. Open [Google Sheets](https://sheets.google.com), create a new one
2. Go to **Extensions > Apps Script**
3. Delete the default `Code.gs` content
4. Paste the entire [`WialonAuditor.gs`](WialonAuditor.gs) file
5. Save

**2. Deploy as Web App**

The Web App handles the OAuth callback from Wialon.

1. **Deploy > New deployment**
2. Gear icon > **Web app**
3. Execute as: `Me`, Access: `Anyone`
4. **Deploy**, authorize when prompted
5. Copy the Web App URL

**3. Set the redirect URL**

1. Find `startOAuthLogin` in the code
2. Replace `'PASTE WEB APP URL HERE'` with your Web App URL
3. Save

**4. Redeploy**

1. **Deploy > Manage deployments** > edit > **New version** > **Deploy**

**5. Use it**

1. Reload the spreadsheet
2. **Wialon Tools** menu appears
3. **Login** to authenticate, then fetch whatever you need

## Menu

| Item | What it does |
|---|---|
| Login | Wialon OAuth login |
| Login Status | Check auth state |
| Fetch Units / Hardware / Sensors / Commands / Profiles / Custom Fields / Drive Rank | Pull data into sheets |
| Fetch All | All 7 categories in one go (single auth, toast progress) |
| Setup Filter Sheet | Interactive unit lookup dashboard |
| Export as JSON | Save all fetched data as a `.json` file in Google Drive |
| Save Snapshot | Save current data as a baseline for comparison |
| Compare with Snapshot | Diff current data vs snapshot — color-coded: green=added, red=removed, yellow=modified |
| Get Web App URL | Shows your deployment URL |
| Logout | Clears token |

## How it works

Single file, no dependencies. Uses the Wialon Remote API:

- `token/login` for auth
- `core/search_items` to fetch units (batches of 1,000, falls back to 200)
- `core/batch` + `UrlFetchApp.fetchAll` for parallel requests (10 batches of 100)
- `unit/get_drive_rank_settings`, `unit/update_hw_params` for unit-specific data
- Sheet writes in 500-row chunks with automatic fallback

To point at a different Wialon host, change the `WIALON_HOST` variable at the top.

## Troubleshooting

| Problem | Fix |
|---|---|
| "Not logged in" | **Wialon Tools > Login** |
| Token expired | Tokens last 30 days, just log in again |
| "Setup Required" | Paste your Web App URL into `startOAuthLogin()` |
| No menu | Reload the spreadsheet |
| Timeout | GAS has a 6-min limit. Run fetches separately for huge fleets. |
| Empty sheets | Check the execution log: **View > Execution log** |

## For developers

If you prefer pushing code with [clasp](https://github.com/google/clasp) instead of copy-paste, it works out of the box — `.clasp.json` is already in `.gitignore`.

## Contributing

PRs welcome. Keep everything in the single `.gs` file, use `var` (not `const`/`let`), and add batch processing for new fetchers. Internal helpers get a `_` suffix.

## Like it?

If this saved you time, [give it a star](../../stargazers). It helps others find it.

## License

[MIT](LICENSE)

---

Not affiliated with Gurtam or Wialon. Use according to your Wialon service agreement.
