# Gluth frontend

React, Vite, Tailwind, and the shadcn `sidebar-07` block provide the local
dashboard. All displayed data comes from the FastAPI server in the repository
root; the UI does not use mock project data.

From the repository root, run `python dev.py` with the virtual environment's
Python. This starts the API on port 8800 and Vite on port 5173. Vite proxies
`/api` to FastAPI, so open <http://127.0.0.1:5173>.

For a single server, run `npm run build` here and start FastAPI from the
repository root. It serves `dist/` at <http://127.0.0.1:8800>.

The station view can upload source documents and a PowerPoint template, build a
draft mapping, and generate a PowerPoint from entries that are not flagged for
review. It also offers the review spreadsheet and generated deck as local
downloads. See the root `QUICKSTART.md` for the full workflow and limits.
