# Retail Feasibility — setup and user guide

## 1. Put the dashboard on GitHub

1. Sign in to GitHub and select **New repository**.
2. Use **Retail-Feasibility** for the repository name. The displayed dashboard name remains **Retail Feasibility**.
3. Create the repository.
4. Extract `Retail-Feasibility.zip` on your computer.
5. In the repository, choose **Add file → Upload files**. Drag all extracted files and folders into the upload area. Keep the `vendor`, `signs`, and `fonts` folders together with the root files.
6. Commit the files to `main`. Confirm that `index.html` is directly in the repository root, rather than inside another Retail-Feasibility folder.
7. Open **Settings → Pages**.
8. Under **Build and deployment**, choose **Deploy from a branch**.
9. Choose **main** and **/(root)**, then click **Save**.
10. Wait for GitHub's deployment to finish. Open the site link shown on the Pages screen.

There is no installation or build step. Open the published website to use the dashboard. Double-clicking the HTML file on a device does not provide the web environment required by JavaScript modules and browser saving.

Official reference: [GitHub Pages publishing source](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

## 2. Start with the supplied defaults

The default model is the supplied Dhanmondi ABC Tower workbook. You can use the dashboard before uploading anything. The six workbook schedules are embedded in the repository, so you do not need to keep an Excel file or a data folder there.

The sections appear in this order: **Data entry → Information → Monthly PNL → Yearly -5yrs PNL → category → location → Investment**.

Data entry contains all the screenshot's project, rent, investment, sales, GP, basket, agreement, rent-increase, and manpower fields. It also contains district, Google Map Location, the existing-outlet count, Sales Given By, Opened by, and the signature manager.

## 3. Upload Location File

The filename can be anything. Upload either a workbook containing the location tab, or a file containing just that schedule. Keep the original column positions and row layout. If there is no tab named `location`, the application looks for the matching location headers.

The Information rules are:

| Information field | Information cell | Source rule |
| --- | --- | --- |
| Project name | C3 | location!B5 |
| SFT | C5 | location!F9 |
| Per SFT rent | C7 | location!G5 |
| Total rent | C8 | C7 × C5 |
| Rent with VAT | C9 | C8 × 115% |
| Security | C10 | location!N5 |
| Advance | C11 | location!M5 |
| Total investment | C13 | Fixed asset investment + Advance |

The source uses row 3 for column headings, row 5 for the first outlet, rows 5–8 for area entries, and F9 for total SFT. Google Map Location is copied from location!C5. Select the district to run the nearby-outlet check.

Ground floor, sales per day, basket size, comparing outlet name, other income %, business agreement years, and rent-increase inputs remain manual as in the Information sheet. The upload does not invent formulas for source cells that were blank or hardcoded.

## 4. Upload Investment File

Any filename is accepted. Keep the Investment schedule's layout: the detailed work headings are in row 4, and the fixed asset investment total is in **G3**. The supplied default links G3 to G285.

The upload refreshes the Investment section. Information!C12 links to Investment!G3, and Information!C13 adds the advance to the fixed asset investment. The linked depreciation, financing costs, PNL, NPV, IRR, and payback calculations then recalculate.

Use the same schedule structure as the supplied file. Uploads with missing required fields, unsupported formulas, or unavailable cross-sheet references show an error and leave the current model active.

## 5. Upload Category Excel and choose a version

The default category mix and GP rates come directly from the supplied workbook. Upload categories using these columns:

| Column | Content | Example |
| --- | --- | --- |
| A | Category name | COMPANY GOODS |
| B | Contribution % | 58.5% |
| C | Daily sales | Calculated by the dashboard |
| D | Category GP% | 19% |

Use real Excel percentages or text ending in `%`, rather than entering 58.5 as a numeric percentage value. For example, Excel's underlying value for 58.5% is 0.585.

The first row should contain Category, Contr %, Daily Sales, GP%. Subsequent rows contain the categories. A final Total row is supported. Additional category rows are supported; the GP formula links to the new total row automatically.

Category daily sales = contribution × Sales per day. Total GP% = SUMPRODUCT(category GP%, contribution). Contributions are used as uploaded; they are not automatically rescaled. The category section shows the contribution total for review.

Each accepted upload adds a version: **Default**, **Upload 1**, **Upload 2**, and so on. The newest upload becomes active. Use the dropdown to switch to any saved version. Previous uploads remain available.

Versions and project inputs are saved **in the same browser for the same dashboard URL**. They are not written back to GitHub and are not shared with other devices or users. Use **Project backup & guidance → Backup project** to keep a portable copy. Use **Restore project backup** on another device or after changing the dashboard address.

## 6. Manual entry

Click any Data entry field to change it. Percent fields use percentage points: enter **17.3** for 17.3%, and **6** for 6%.

For a calculated field, entering a value creates a manual override. Click **Use sheet rule** to return that field to its original formula. Clearing a numeric entry restores its sheet value or formula. A typed zero is retained as zero.

You can edit manpower quantities and salary rates. Each role's total remains quantity × salary. Additional expense inputs and WACC are available below the manpower panel.

The dashboard preserves the workbook's original conventions:

- Information month sales = daily sales × 30. Monthly PNL sales = daily sales × 365 ÷ 12.
- Total investment in Information = fixed assets + advance. The five-year model's initial investment also includes security, as in the source.
- Information F22 sums manpower rows 6–18. Security and Cleaner remain separate rows outside that source subtotal. The subtotal is not changed to a new range.
- Rent-increase inputs C21/C22 are recorded in Information. The original five-year rent formulas hardcode the 10% steps in years 3 and 5, rather than referencing those inputs. Those formulas are preserved.
- Yearly H54 uses the original `H51--H53` formula. It is preserved exactly rather than corrected during the dashboard build.
- The source's NPV, IRR, ROI, and payback formulas are retained.

If the inputs make IRR unavailable or a formula divides by zero, the original Excel error is shown in the model and retained in exports. It is not converted to zero. Structural errors such as a missing sheet or unsupported formula must be fixed before export.

**Restore default inputs** restores the supplied model and default signature selections. Uploaded category versions and uploaded signature images remain available for selection.

## 7. Existing Outlet No. within 1 KM

1. Choose the **District**.
2. Paste the **Google Map Location**: a Google Maps pin link, a supported short link, coordinates, or an address.
3. The check runs after the field changes. Use **Recheck map** to retry.

The dashboard uses the reference dashboard's 900 mapped outlet points and the selected-district filter. It calculates straight-line distance using the same Haversine rule and includes outlets at a distance of **1 KM or less**. Results are sorted by distance and show outlet code, outlet name, RHO, Zonal, and distance.

The parser supports pin coordinates, raw coordinates, Street View positions, and the original link-resolution rules. A named place's map-view centre is not substituted for its actual pin. Short links and address lookup require the external resolver/map services and internet access.

A failed lookup clears the automatic result instead of keeping an old count. You can enter a count manually. The status message distinguishes manual entry from an automatic map result.

The **Include in Excel & PDFs** checkbox starts checked. When unchecked, the count block is omitted from the PDFs and the nearby-outlet worksheet is omitted from the Excel download.

In the Excel nearby-outlet sheet, **B4** contains resolved coordinates when available. The uploaded Google Map Location and matched outlet details are included there.

The original resolver configuration is included in `map-resolver-config.json`. If its Apps Script service later becomes unavailable, paste a full link with coordinates or latitude/longitude, or update that configuration to an equivalent working resolver. Direct coordinates and the bundled outlet-distance calculation do not require the short-link resolver.

## 8. Signatures and approvals

All seven reference signatories are available in Data entry:

| Position | Role | Default name | Signature enabled by default |
| --- | --- | --- | --- |
| 1 | Prepared by | Md Aftab Ul Islam | Yes |
| 2 | Sales Given By | Mr. Mamun | Yes |
| 3 | Opened by | Mr. Abu Sayem Mohammad Al Azad | No |
| 4 | Viewed by | Salah Uddin Misbah | No |
| 5 | Viewed by | Saiful Alam Rasel | Yes |
| 6 | Viewed by | Tamal Paul | No |
| 7 | Approved By | Abu Naser | No |

Signatory 2 follows **Sales Given By** for the name, while its designation remains independently editable. Signatory 3 follows **Opened by** and **Opened by Designation**. Recognized names fill the opened-by designation. Other names can be typed manually.

If Signatory 3 is enabled, it uses the **exact image selected for Signatory 2**. It does not have an independent image selector. Disabling Signatory 3 removes its signature ink.

The checkbox controls the image, not the signatory's name, role, designation, or approval line. The signed PDF keeps all approval positions in the original 4 + 3 arrangement. Saiful's selected signature appears as the review signature on pages 1 and 2 when enabled. PDF signature boxes retain the reference's enlarged size.

Use **Manually change** to override an automatic name/designation link. Use **Use automatic values** to restore it. Upload PNG or JPEG signatures through the signature manager and select them for the appropriate signatory. Uploaded signature images are included in project backups.

## 9. Download reports

### Download 3 page pdf

The signed PDF contains exactly:

1. Information and manpower.
2. Monthly PNL.
3. Yearly -5yrs PNL and approvals.

The location, category, and Investment source schedules are excluded. Table text and borders are black on white. PDF fonts are included in the repository and embedded in the document.

### Management PDF

The same three financial pages are exported without signature images or the approval block.

### Download Excel with Rules

The workbook retains the six schedules and the active category's formulas. It includes an Approval sheet and, when selected, Existing outlets within 1 KM. The download includes cached results and requests recalculation when opened in Excel. Manual overrides are exported as the values entered. Formula-driven fields retain their formulas.

The Excel download represents the active model. Use the project backup to keep **all** category upload versions and browser settings.

## 10. If something does not load

- Blank dashboard: confirm that all files were uploaded at the repository root and the website was opened through GitHub Pages.
- Missing signature: confirm that the `signs` folder and its manifest were uploaded. Select an image and enable the relevant checkbox.
- PDF font error: upload the complete `fonts` folder.
- Location not resolved: select District, check the link, and use Recheck map. Try raw coordinates if an external service is unavailable.
- Old browser data: restore the required project backup, or use Restore default inputs when you want the supplied baseline.
- Upload rejected: keep the original Location/Investment layout or the Category columns described above. A rejected upload does not replace the active project.

## Validation completed

All 1,189 baseline formulas matched the source workbook's saved results. Changed uploads, category versions, signature links, map parsing and distance rules, formula exports, and both PDF modes were checked. The signed PDF was rendered and inspected for all three pages. Interactive browser testing and live external map-service checks were not available in the build environment.
