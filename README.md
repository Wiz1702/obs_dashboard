# Jobs at the Crossroads

An interactive Observable dashboard that explains how AI adoption is showing up in job postings, salaries, automation risk, and reskilling demand.

The experience starts on a homepage narrative about how AI is changing the work inside jobs. The `See analytics` button opens a separate analytics page with a guided chart story. Each story step reveals one chart at a time instead of showing every plot at once.

## Features

- Separate homepage with AI/job-impact context and a `See analytics` button
- Separate analytics page for exploring the graphs
- Guided step-by-step analytics story
- Multi-select global filters for industry, seniority, company size, displacement risk, and AI adoption stage
- Animated year range control
- KPI summary cards
- Salary, AI adoption, risk, reskilling, and correlation views
- Seniority ordering from `Intern` through `Executive`
- Percentage-based displacement risk comparison with a fixed 0-100 axis

## Project Structure

- `index.html` - Homepage with the Jobs at the Crossroads graphic and narrative
- `analytics.html` - Analytics page that renders the Observable graph story
- `app.js` - Runtime setup, theme toggle, and homepage navigation behavior
- `notebook.js` - Observable notebook module used by the app
- `observable-notebook.js` - Standalone copy of the Observable notebook code, without HTML
- `dashboard.css` - Dashboard, homepage, and chart UI styling
- `runtime.js` - Observable runtime bundle
- `index.js` - Re-exports the active notebook module
- `files/` - Excel dataset files used by the notebook

## Run Locally

Use any static file server from the project directory.

```bash
python3 -m http.server 8000
```

Then open:

```text
http://localhost:8000/
```

The analytics page is available at:

```text
http://localhost:8000/analytics.html
```

If dependencies are installed, the npm script can also be used:

```bash
npm run dev
```

## Data Flow

The notebook loads `ai_jobs_merged_dataset.xlsx`, normalizes key fields, then filters the dataset based on the selected controls. The chart dispatcher receives the current story step, filtered rows, D3, and Observable Plot, then renders the active chart.

`observable-notebook.js` contains the same Observable module code as `notebook.js` for cases where you want the analytics logic separately from the HTML shell.
