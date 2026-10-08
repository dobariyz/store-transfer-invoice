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

The app can store products and transfer history in Supabase so they are shared across browsers and devices. Without Supabase settings, it falls back to browser-local storage.

### Supabase shared data setup

1. In Supabase Authentication settings, disable public sign-ups. Add one account per staff member under **Authentication → Users**.
2. In **SQL Editor**, run [`supabase/shared-data-policies.sql`](supabase/shared-data-policies.sql). It enables access only for signed-in users and keeps Row Level Security enabled.
3. In Railway service variables, add `VITE_SUPABASE_URL` (the project URL) and `VITE_SUPABASE_PUBLISHABLE_KEY` (from the Supabase Connect dialog or API Keys page). These are read during the build, so redeploy after setting them.
4. Sign in to the app with a staff account. Products, transfers, transfer items, and reconciliation links will then use the shared database.

Do not put a Supabase secret or `service_role` key in a `VITE_` variable or frontend code. Existing records in browser-local storage are not automatically merged into Supabase; export and merge those records before treating the shared database as the complete history.

## Input sheets

For transfers, use columns `Product Name` or `SKU`, `Quantity`, `From Store`, and `To Store`. For bulk products, use `SKU`, `Product Name`, and `Unit Price`; `Category`, `Price Type`, and `Pack Qty` are optional.

## Important

The tax treatment and wording of an internal transfer document depend on how the stores are legally structured. Confirm the invoice format and whether HST should be charged with your accountant before relying on it for tax records.
