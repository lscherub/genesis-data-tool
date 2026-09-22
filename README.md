# Genesis POS Excel Cleaner

A one-page browser app for cleaning Genesis Amber/Cybex POS Excel exports.

## What it does

- Accepts `.xlsx`, `.xls`, and `.csv`
- Detects the header row automatically, including exports with blank rows above it
- Finds columns by column name rather than fixed column position
- Outputs these columns in this exact order:
  1. Sku
  2. Product Number
  3. Description
  4. Vendor
  5. Brand
  6. List Cost
  7. Price
  8. Size Desc
- Removes all other columns
- Shows a preview before downloading
- Creates a new `.xlsx` file
- Processes files locally in the browser

## GitHub Pages

1. Create a GitHub repository.
2. Upload `index.html`, `style.css`, `app.js`, and `README.md`.
3. Go to **Settings → Pages**.
4. Under **Build and deployment**, select **Deploy from a branch**.
5. Select the `main` branch and `/ (root)`.
6. Save.
7. GitHub will provide the Pages URL.

## Privacy

The Excel file is processed in the browser. This app does not contain a backend or upload the POS file to a server.

## Note

The app loads SheetJS from its public CDN. For completely offline/self-contained deployment, the SheetJS library can be downloaded and hosted locally in the repository.
