# Invoice Generator

A Next.js application to generate professional Excel invoices from employee data.

## Features

- Upload Employee Data (Excel file with NAME, EPF, DEPARTMENT columns).
- Configure invoice details (Dates, Company Info, Products).
- Upload Top Logo and Bottom Left Image.
- Generate and download formatted Excel invoices.

## Setup

1. Install dependencies:
   ```bash
   npm install --ignore-scripts
   ```
2. Run the development server:
   ```bash
   npm run dev
   ```
   (Note: If `npm run dev` fails due to path issues, use the VS Code task "Run Dev Server" or run `node "node_modules/next/dist/bin/next" dev`)

## Usage

1. Open http://localhost:3000.
2. Upload your Employee Data Excel file.
3. Fill in the invoice details.
4. (Optional) Upload images for the logo and footer.
5. Click "Generate Invoices".
