# Wigwam Analytics Generator

Vercel-ready version of the Wigwam report generator.

## What it does

- Upload Amazon Business Report CSV
- Upload Amazon listing / inventory TXT or CSV
- Optional GoBros direct sales CSV
- Generates a Wigwam Excel analytics workbook from the bundled Wigwam template

## Deploy to Vercel

Create a GitHub repository from this folder:

```text
wigwam-analytics-vercel
```

Then import that repository in Vercel.

Vercel will use:

- `index.html` as the static page
- `api/generate.py` as the Python serverless function
- `requirements.txt` to install `openpyxl`

## Important upload note

Vercel serverless functions have request body size limits. The current monthly files you tested are small enough, but much larger exports may need a Blob-storage upload flow later.
