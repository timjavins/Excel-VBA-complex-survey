# Excel VBA Asset Mapping Template

This repository contains a **generic Excel/VBA template** for mapping one set of assets (for example, terminals) to another set of endpoints/devices. It was sanitized for public use; all examples are synthetic.

## Repository layout

- `/asset-mapping form module*.bas` - core workflow modules (initialize form, load data, dropdowns, validation, submission)
- `/asset-mapping form workbook module.bas` - `Workbook_Open` entry point
- `/worksheet FORM module.bas` - form worksheet change handler
- `/worksheet Terminals module.bas` - optional worksheet helper cleanup macro
- `/Dashboard-Tools/` - optional helper scripts/notebook for post-processing CSV output

## Workbook workflow

1. `Workbook_Open` runs `InitializeForm`.
2. User enters a **Location ID** in `FORM!A8`.
3. `LoadForm` loads matching terminal IDs and controller dropdown values.
4. Selecting a controller populates endpoint/device dropdowns.
5. Selecting an endpoint fills hidden metadata fields (port/channel/location).
6. `SubmitAnswers` appends rows to CSV output files.

## Expected workbook sheets and schema

By default, the code looks for neutral sheet names and falls back to legacy names:

- `FORM`
- `Terminals` (fallback: `EdgeUnits`)
  - Column E: Location ID
  - Column F: Terminal ID
- `Locations` (fallback: `Regions`)
  - Column A: Location ID
  - Column B: Location Group key used to locate a device sheet
- `<LocationGroup>Devices` (fallback: `<LocationGroup>Assets`)
  - Columns B:F used as location/controller/port/channel/endpoint mapping source

## Configuration

Update constants in `asset-mapping form module1.bas`:

- `FORM_PASSWORD`
- `BASE_DATA_PATH`
- neutral/legacy sheet-name constants (if your workbook uses different names)

Optional Dashboard constants are in `Dashboard-Tools/asset_mapping_data_csv_munger.bas`:

- `ONEDRIVE_ORG`
- `FILE_PATH`
- `TEST_USERNAME`

## Output behavior and logging

`SubmitAnswers` writes to:

- `BASE_DATA_PATH + <WindowsUsername>_asset_mapping_data.csv`
- `BASE_DATA_PATH + Completed Locations.csv`

Rows include:

- selected form fields
- derived controller-port field
- Windows username (`Environ("USERNAME")`)
- local timestamp (`yyyy-mm-dd hh:mm:ss`)

If you need to disable or modify username/timestamp logging, edit `SubmitAnswers` in `asset-mapping form module6.bas`.

## Network/OneDrive assumptions

This template assumes:

- access to a configured shared/network path (`BASE_DATA_PATH`)
- optional OneDrive path conventions in Dashboard helper scripts
- possible VPN/network dependency for remote file access

## Data handling and privacy disclaimer

This repository is a template. Do **not** commit real production datasets, credentials, customer information, or sensitive identifiers. Generated CSV outputs may contain operator usernames and timestamps depending on your configuration.
