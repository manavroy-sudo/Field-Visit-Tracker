# Field Visit Tracker (InsuranceDekho · Motor & Health)

Node.js/Express wrapper around the original single-page dashboard, so it can be run and deployed like a normal Node app instead of a static Netlify upload.

## Structure
```
field-visit-tracker/
├── public/
│   └── index.html   <- your original dashboard (unchanged: HTML/CSS/JS, all client-side)
├── server.js        <- Express server that serves public/index.html
├── package.json
└── .gitignore
```

## Run locally
```
npm install
npm start
```
Then open http://localhost:3000

## Notes
- All app logic (login, forms, calendar, admin dashboard) still runs entirely in the browser using `localStorage`, exactly as before — the Node server only serves the static file.
- The Google Sheet write-back still depends on `APPS_SCRIPT_URL` inside `public/index.html` — replace `YOUR_APPS_SCRIPT_URL_HERE` with your deployed Apps Script Web App URL once you set that up.
- To deploy: any Node host (Render, Railway, Fly.io, Azure App Service, etc.) works — just point it at `npm start`. Netlify itself doesn't run persistent Node servers for this pattern; if you want to stay on Netlify, deploy `public/` as a static site (same as before) instead of using `server.js`.
