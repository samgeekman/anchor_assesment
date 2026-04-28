# Dropping Anchor Self-Assessment Site

Static site for the SCDC "Dropping Anchor" self-assessment tool with:

- Two top-level sections:
  - Assessing our strengths and challenges
  - Getting access to the support you need
- 1-5 self-scoring and comments for each statement
- Automatic section score summary
- Export/download options for Word (`.doc`) and PDF (`.pdf`)

## Local preview

Run from this folder:

```bash
python3 -m http.server 8080
```

Then open `http://localhost:8080`.

## Deploy to GitHub Pages

1. Create a GitHub repository and push these files to the default branch (`main`).
2. In GitHub, open `Settings` -> `Pages`.
3. Under `Build and deployment`, set `Source` to `Deploy from a branch`.
4. Select branch `main` and folder `/ (root)`.
5. Save and wait for deployment.

Your published URL will be:

`https://<your-github-username>.github.io/<repo-name>/`

## Notes

- This is a no-build static site (`index.html`, `style.css`, `app.js`, `data.js`).
- PDF export is client-side using `jsPDF` from CDN.
- Word export downloads an HTML-based Word document (`.doc`) containing full responses and scores.
