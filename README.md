# HANSA-FLEX Crimp App: demo package

This is a ready-to-run packaging of the Crimp App concept as a standard
Vite + React project, so it can be uploaded to a sandbox environment or
run locally as a working demo, without any backend.

Concept demo, sample data. Not valid production specifications. Always
consult the official HANSA-FLEX crimping tables.

## What is in this package

```
crimp-app-demo/
├── package.json
├── vite.config.js
├── index.html
├── README.md
├── .gitignore
└── src/
    ├── main.jsx        (entry point, mounts the app)
    ├── App.jsx          (the application itself)
    ├── storageShim.js   (demo-only persistence, see below)
    └── index.css        (Tailwind entry point)
```

## Quick start

Requires Node.js 18 or newer.

```bash
npm install
npm run dev
```

This starts a local dev server (Vite prints the URL, typically
http://localhost:5173) with the app running and hot-reloading.

To build a static, deployable version:

```bash
npm run build      # outputs static files to dist/
npm run preview    # serves that build locally, to check it before deploying
```

The `dist/` folder after `npm run build` is a plain set of static files
(HTML, CSS, JS) that can be uploaded to any static host.

## About the "Update data" feature and storageShim.js

The app includes an "Update data" panel for uploading a new .xlsx
crimping table at runtime (a downloadable .xlsx template is provided in
the panel). In its original environment, that feature
saves data through a managed key/value service. Outside that environment
there is no such service, so `src/storageShim.js` provides a small
stand-in backed by the browser's own localStorage: uploading and
confirming new data in the demo will persist it in that browser, across
page reloads, without needing a server.

This is a demo convenience, not a real multi-user backend: the data is
local to one browser and is not shared between different visitors or
devices. For an actual shared/multi-user deployment, replace
`storageShim.js` with an adapter that calls a real backend API instead.
A worked example (a small Node.js/Express service with a database) is
described in the accompanying technical design document, section 3.10.

## Suggested tools for running this as a demo

See the recommendation in the accompanying chat message for a fuller
comparison; in short:

- For the fastest "upload the code, get a shareable working preview"
  experience, use an in-browser sandbox such as StackBlitz or
  CodeSandbox: both can import a project from a ZIP file or a Git
  repository and will install dependencies and start the dev server
  automatically.
- For a more permanent, public demo URL, build the project
  (`npm run build`) and deploy the resulting `dist/` folder to a static
  host such as Netlify or Vercel (both support drag-and-drop deployment
  of a folder, no server configuration required).

## Notes

- The application text is available in Dutch, English and German via
  the flag switcher in the header; Dutch is the default on load.
- The disclaimer banner and sample-data labelling are intentional and
  should remain visible for as long as the data set is not verified
  production data.
