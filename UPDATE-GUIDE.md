# Retail Feasibility update

1. Extract this ZIP.
2. Open https://github.com/aftabz-lab/Retail-Feasibility.
3. Choose Add file → Upload files. Upload all extracted files directly at the repository root and replace existing files.
4. Commit to main and wait for the existing GitHub Pages deployment.
5. Open the dashboard and press Ctrl+Shift+R once.

Use Clear inputs for a new blank form. Existing intentional saved entries, uploaded files, category versions and signatures are retained. Older untouched sample defaults no longer populate the form.

Select Store Type applies Super Store (39), Mega Store (79), or PNP Convenience Store (25). Quantities and salaries remain editable.

The feasibility badge follows Monthly PNL row 43 only. Auto Correct uses Sales per day in multiples of 1,000 and keeps other inputs. Sales arrows and keyboard arrows change by 1,000.

The ZIP contains only changed files and the new controls module. Keep the existing workbook rules, vendor, fonts, signs, map files and other repository files. No token change or Apps Script setup is needed for this update.

Validation covered blank forms, store-type quantities, manual overrides, source uploads, field dependencies, 1,000-step controls, sales correction, saved-state migration, category history and both three-page PDF modes. Browser UI testing was unavailable; application events were tested with a simulated DOM.
