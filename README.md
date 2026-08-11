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

## Data storage

The deployed app stores its product list and transfer history in the browser's local storage. It works without a backend, but data is private to that browser and will not sync between the Dixie and Vaughan devices. Before production use across multiple users or devices, add authentication and a shared database.

## Input sheets

For transfers, use columns `Product Name` or `SKU`, `Quantity`, `From Store`, and `To Store`. For bulk products, use `SKU`, `Product Name`, and `Unit Price`; `Category`, `Price Type`, and `Pack Qty` are optional.

## Important

The tax treatment and wording of an internal transfer document depend on how the stores are legally structured. Confirm the invoice format and whether HST should be charged with your accountant before relying on it for tax records.
