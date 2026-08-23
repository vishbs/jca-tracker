# EU JCA Tracker

An unofficial, easier-to-read interface for the European Commission's public list of **Joint Clinical Assessments (JCAs)** under HTA Regulation (EU) 2021/2282 — the EU's joint scientific evaluation process for new medicines, where one member state (the assessor) leads and another (the co-assessor) supports.

The EU only publishes this data as a periodically-refreshed Excel file, with no browsable website. This project turns that same data into:

- **A searchable, filterable table** of every ongoing, completed, and discontinued JCA — filter by status, substance type, or country, search by medicine name or indication.
- **Per-medicine detail pages** with assessor/co-assessor, key dates, regulatory flags (orphan product, accelerated assessment), and links to the official JCA report / EMA product page / Union Register entry.
- **Charts & stats** — status split, substance type mix, country participation by role, orphan/accelerated assessment counts, and time-in-assessment for ongoing JCAs. Every bar and flag is clickable and opens a popover listing the matching JCAs.
- **A "data last updated" date**, read directly from the source Excel, so it's always clear how fresh the data is.

Live site: **https://vishbs.github.io/jca-tracker/**

## Data source

Built from the EU's public JCA tracker Excel export:
https://health.ec.europa.eu/health-technology-assessment/implementation-regulation-health-technology-assessment/joint-clinical-assessments_en

## Updating the data

This project is not affiliated with the EU or EMA and does not auto-fetch new data. To refresh it:

1. Download the latest Excel from the link above and replace `data/hta_ongoing-jca_en.xlsx`.
2. Regenerate the site's data:
   ```sh
   python3 -m venv .venv && .venv/bin/pip install pandas openpyxl
   .venv/bin/python3 scripts/build_data.py
   ```
3. Commit and push — GitHub Actions rebuilds and redeploys automatically.

## Development

```sh
npm install
npm run dev       # http://localhost:4321
npm run build     # outputs to ./dist
```

Built with [Astro](https://astro.build) as a static site, deployed to GitHub Pages.
