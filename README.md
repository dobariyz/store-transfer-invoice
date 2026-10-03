# Store Transfer Invoices

A React application for importing a transfer sheet, checking each product against the product database, and producing a printable store-to-store transfer invoice.

## Run locally

1. Install Node.js 20 or newer.
2. Run `npm install`.
3. Run `npm run dev`.

## Build for hosting

Run `npm run build`. Upload the contents of `dist/` to any static host, or connect the repository to Vercel or Netlify and use these settings:

- Build command: `npm run build`
- Publish directory: `dist`

## Deploy to Railway

1. Create a new Railway project from the connected GitHub repository.
2. Railway will use `pnpm run build` and then `pnpm run start` from `railway.toml`.
3. In Railway, generate a public domain after the first deployment.

### Azure invoice scanning

The Invoice to Excel tab uses Azure Document Intelligence's prebuilt Invoice or Receipt model. In Railway, add these service variables after creating a Document Intelligence resource:

- `AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT`: the resource endpoint from **Keys and Endpoint**.
- `AZURE_DOCUMENT_INTELLIGENCE_KEY`: either resource key. Keep this value in Railway variables; do not add it to frontend `VITE_` variables or commit it to Git.

The scan endpoint accepts PDF, JPG, and PNG files up to 4 MB. Uploads are analyzed in memory and are not saved by this app. Azure's F0 tier currently includes up to 500 pages per month.

## Data storage

The deployed app stores its product list and transfer history in the browser's local storage. It works without a backend, but data is private to that browser and will not sync between the Dixie and Vaughan devices. Before production use across multiple users or devices, add authentication and a shared database.

## Input sheets

For transfers, use columns `Product Name` or `SKU`, `Quantity`, `From Store`, and `To Store`. For bulk products, use `SKU`, `Product Name`, and `Unit Price`; `Category`, `Price Type`, and `Pack Qty` are optional.

## Important

The tax treatment and wording of an internal transfer document depend on how the stores are legally structured. Confirm the invoice format and whether HST should be charged with your accountant before relying on it for tax records.
