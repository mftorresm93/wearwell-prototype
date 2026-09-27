# Wearwell prototype

This project is a static front-end prototype for a mobile-first outfit builder.

## Share-ready prototype path

The app is already structured for static hosting. The simplest path is to deploy the repo root to a static host such as GitHub Pages, Netlify, or Vercel.

### Public URL behavior

- Root entry page: `index.html`
- App page: `data/outfit-app.html`
- Data files used by the app: `data/*.json`

### Quick deploy options

#### GitHub Pages
1. Push this repo to GitHub.
2. In the repo settings, enable GitHub Pages.
3. Publish from the root branch.
4. The app will be available at a URL like:
   `https://<username>.github.io/<repo-name>/`

#### Netlify or Vercel
1. Import the repo.
2. Set the publish directory to the repo root.
3. Deploy.

### Notes

- This is still a prototype. It is intentionally front-end only.
- For a real app later, the next step is adding auth, real storage, and affiliate-link tracking.

### Local preview

From the repo root, you can preview it with:

```bash
python -m http.server 8000
```

Then open:

```text
http://localhost:8000/
```
