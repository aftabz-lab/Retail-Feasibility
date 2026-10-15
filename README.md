# Retail Feasibility

Shwapno retail feasibility dashboard prepared from the supplied Dhanmondi ABC Tower workbook. The interface follows the existing Feasibility dashboard and retains its signature and Google Maps distance rules.

## Publish on GitHub

1. Create a repository named **Retail-Feasibility**. The dashboard title is **Retail Feasibility**.
2. Extract the ZIP. Upload **all the files and folders inside it** to the repository root. `index.html` must appear directly on the first repository page, beside `app.mjs` and `styles.css`.
3. Commit the uploaded files to `main`.
4. Open **Settings → Pages → Build and deployment**. Select **Deploy from a branch**, branch **main**, folder **/(root)**, then **Save**.
5. Open the URL displayed by GitHub Pages after deployment finishes. For the `aftabz-lab` account and this repository name, the expected address is `https://aftabz-lab.github.io/Retail-Feasibility/`.

Do not upload the ZIP itself as the website. No build command, package installation, database, data folder, or source Excel file is needed in the repository.

Read [GUIDE.md](GUIDE.md) for upload formats, manual entry, category history, signatures, map checks, and exports. An in-app guide is also included as [guide.html](guide.html).

GitHub's setup instructions: [Configuring a publishing source](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

## Current update

Fresh project and manpower inputs start blank. Source uploads populate associated fields. Select Store Type applies the supplied 39/79/25 manpower matrix, with manual editing available. The row 43 control checks all monthly and annual values. Auto Correct changes Sales per day only, using multiples of 1,000. Sales arrows change the entered value by 1,000.

## Included behavior

- Sections: Data entry, Information, Monthly PNL, Yearly -5yrs PNL, category, location, Investment.
- Manual entry for every information field in the screenshot, plus manpower quantities and salaries.
- Separate Location, Investment, and Category Excel uploads. Filenames can vary. The two primary upload buttons use the same dimensions as Download 3 page pdf.
- Default category data comes from the supplied workbook. Every accepted category upload is retained separately as Upload 1, Upload 2, and so on. Selecting an earlier version recalculates the workbook.
- Seven default signatories, original selection defaults, automatic name links, and the shared Signatory 2/3 signature rule.
- District-based 1 KM outlet count, the original map-link parser and resolver configuration, and an editable count.
- Three-page signed PDF containing Information, Monthly PNL, and Yearly -5yrs PNL. The category, location, and Investment source sheets are excluded from both PDF modes. Management PDF omits signature images and the approval block.
- Download Excel with Rules retains the active formulas and source schedules, with approval information and the optional nearby-outlet sheet.
- Browser saving and a project backup/restore option.

## Rule preservation

`workbook-rules.mjs` contains the six source sheets, cell values, number formats, and all **1,189 formulas** from the supplied file. The application calculates those formulas directly. The calculation engine does not replace the source's financial assumptions with those of the older dashboard.

Manual edits override the corresponding cell. Formula inputs offer **Use sheet rule** to remove a manual override. Location or Investment uploads reset the mapped information overrides and their immediate totals so the newly uploaded file can populate them. Category selection restores the category-based GP formula.

The store-type update changes Information F22 to sum F6:F20, matching the supplied manpower matrix. Other financial formulas, including the five-year rent step-ups and total after financing-cost formula, remain unchanged.

## Verification

All 1,189 baseline formula results were compared against the original workbook's cached results and matched. Automated checks covered source uploads with changed filenames, linked calculations, manual zero values, category version switching and extra category rows, signature links, coordinate parsing, district filtering, distance calculations, both PDF modes, and formula-preserving Excel exports. The three signed PDF pages were rendered and visually inspected.

Interactive browser QA and live external map-service calls were unavailable in the build environment. The original resolver URL and map parsing/counting rules are retained. The guide explains retry and manual-entry options.

