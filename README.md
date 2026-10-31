# PU Dataflow Clean-up

Helper script for the Productsup feed restructure clean-up. It reads a site's
Dataflow and Data Sources pages and produces an Excel sheet showing which
attributes are used and which data sources can be paused.

**Credits:** Based on the original attInvestigator script by Kell Huang.
Clean-up sheet extension by Priti More, built with AI assistance.

## What it outputs
- **all_attrs_inter_export:** Kell's original Import → Intermediate → Export view
- **Mapping:** Export → Intermediate → Import → Data source
- **Data Sources:** each source marked Keep or Pause
- **Allow Columns:** the columns to allow per source (include the join key from *ID Column In This Feed*)

## Setup
1. Install Python and the `openpyxl` package: `py -m pip install openpyxl`
2. Put `dataflowcleanup.py` in a folder, e.g. `C:\dataflowcleanup`

## How to use
1. In PU, open the site's Dataflow, press F12 → Inspect → find `<tbody>` (1 of 3) → right-click → Copy element. Save it as `<sitename>.txt` in the folder.
2. Open the site's Data Sources page, press Ctrl+A and Ctrl+C, and save it as `data_sources.txt` in the same folder.
3. Run:
   ```
   cd C:\dataflowcleanup
   py dataflowcleanup.py
   ```
4. Pick the number of your site file. The Excel output is saved in the folder.

## Important checks before pausing
- Never pause a **Main** data source. On sites with more than one main source, these are marked CHECK.
- Keep sources used by `___` filter attributes (Skip/Whitelist/Blacklist rules).
- Read the **MISSING** warning: a used column isn't arriving in the import.
- After changes: run the site, check the Error Log, and compare Imported/Exported counts.

## Tested on
The 20 US - VIDAXL sites and the CH test site 1183374.
