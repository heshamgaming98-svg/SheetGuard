# SheetGuard Free

A no-backend, client-side Excel auditing prototype that performs real workbook-derived analysis in the browser.

## Run
Open `index.html` in a modern browser. For best results, serve the folder with any simple static server.

## Publish for free
GitHub Pages can host static HTML/CSS/JS from a repository on GitHub Free. See:
https://docs.github.com/en/pages/getting-started-with-github-pages/what-is-github-pages

## Important
- Excel files are processed in the browser.
- No SheetGuard server is required.
- The app uses SheetJS from its CDN.
- The audit engine is deterministic and derives its counts/findings/score from the uploaded workbook.
- This is intentionally a free static version. A future production version can add a backend, authentication, database, PDF reports, and stronger formula dependency analysis.
