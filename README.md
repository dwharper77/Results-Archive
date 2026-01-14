# Results Archive

Minimal, client-only Progressive Web App that:
- loads a user-uploaded `.pkl` file
- unpickles a pandas `DataFrame` using Pyodide (in-browser Python)
- converts it to JSON records
- renders simple pivot-style views

## Run locally

Service workers require `http://` (not `file://`).

From this folder:

```powershell
python -m http.server 8000
```

Then open:

- http://localhost:8000/

## Publish to GitHub + run externally (GitHub Pages)

This is a static site, so GitHub Pages works great.

1) Create a new repo on GitHub (for example `pkl-pivot-pwa`).

2) In a terminal, run these from this folder:

```powershell
git init
git add -A
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<YOUR_USER>/<YOUR_REPO>.git
git push -u origin main
```

3) Enable GitHub Pages:
- GitHub repo → **Settings** → **Pages**
- **Build and deployment**: Deploy from a branch
- Branch: `main` / Folder: `/ (root)`

Your site will be available at:
`https://<YOUR_USER>.github.io/<YOUR_REPO>/`

### Notes on updates

- This app uses a service worker for an app-shell cache, so your browser may keep old files.
- If you don't see changes after you push, do a hard refresh (`Ctrl+F5`) or clear site data for the Pages URL.

## Notes

- Dimension columns are auto-detected with simple heuristics in `app.js` (`guessDimensionColumns`).
- Pivot rendering is in `pivot.js` and currently uses "last-write-wins" when multiple rows map to the same cell.

