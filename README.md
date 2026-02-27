# Multifamily Underwriter (Local, Pure Python)

This app recreates the underwriting workflow from your workbook in a local browser UI.

- Shaded input cells are editable.
- Formula cells are calculated in Python code (no Excel runtime dependency).
- `Mortgage` and `1.1` logic are internal and hidden from analyst tabs by default.
- `Rent Roll` supports dynamic unit rows driven by a user-entered unit count.

## Requirements

- Python 3.9+
- macOS, Linux, or Windows

No third-party Python packages are required.

## Run (One Command)

```bash
./run_local.sh
```

Open in Brave (or any browser):

- Analyst mode: `http://127.0.0.1:8000`
- Admin mode: `http://127.0.0.1:8000/?admin=1`

## Run in VS Code Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
python app.py
```

## Analyst Test Checklist

1. Start app:
   - `./run_local.sh`
2. Open browser:
   - `http://127.0.0.1:8000`
3. Confirm tabs appear:
   - `Rent Roll`, `Valuation`, `Returns`, `Sensitivity Analysis`
4. In `Rent Roll`, set:
   - `Property Name`, `Property Address`, and `Number Of Units`
5. Confirm the app dynamically generates one row per unit.
6. Fill row inputs (Tenant, Unit, Rent, Utilities, Parking, Pet Fee, Projected Rent, Lease Dates, Security Dep).
7. Click `Run Analysis`.
8. Confirm totals flow into `Valuation`/`Returns` outputs.
9. Click `Reset Defaults` and confirm values return to defaults.

## Admin Mortgage Logic

1. Open `http://127.0.0.1:8000/?admin=1`.
2. Open tab `Admin: Mortgage Logic`.
3. Search for a mortgage cell/formula and edit as needed.
4. Click `Save Mortgage Overrides`.
5. Use `Reload From Server` to discard unsaved edits.

## Data Files

- `data/workbook_model.json`: layout, values, and input metadata
- `data/formulas_expanded.json`: expanded formula map used by Python engine
- `data/admin_overrides.json`: persisted admin mortgage formula overrides
- `data/backup_1_1_sheet.json`: backup snapshot of the hidden `1.1` worksheet
