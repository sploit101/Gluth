# Gluth local run

The app reads customer documents from `sources/`, builds a searchable local
corpus, extracts candidate requirements, and prepares a review spreadsheet for
one station. A human supplies the station name and decides when to write a
PowerPoint copy. Customer files and generated results stay in ignored local
folders.

## Install

Use Python 3.11 or 3.12. In a terminal at the repository root:

**macOS**

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

**Windows PowerShell**

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

On Windows, scanned documents additionally need the local Tesseract executable
on `PATH`. OCR is optional for PDFs with selectable text. Legacy `.doc` files
need LibreOffice or antiword on Windows; `.docx` works without them.

## Run

1. Put customer PDFs, Word files, and other supported source documents under
   `sources/`. Subfolders may be named like `A12345678_Project/1_Anfrage/`.
2. Keep the station template outside `sources/`; pass its path explicitly.
3. From the repository root, run the following with your platform's venv
   Python path (`.venv/bin/python` on macOS,
   `.\.venv\Scripts\python.exe` in PowerShell):

```sh
python run.py --template path/to/Stationsakte.pptx --station "Station 20 Abisolieren" --file Lastenheft
```

This writes `out/zuordnung.xlsx` for review. Yellow rows need human checking.
Enter `ja` or `nein` in the `Freigabe` column and save a reviewed copy.
`ja` includes that exact mapping as human approved; `nein` excludes it.
To use these decisions from the CLI, write the current mapping with
`python akte_schreiben.py --vorlage path/to/Stationsakte.pptx --station
"Station 20 Abisolieren" --freigabe path/to/reviewed.xlsx --nur-sichere`.
The reviewed sheet must match every row of the current mapping; changing a
source, station, or rule means downloading a new sheet. Without a reviewed
sheet, `run.py --write` creates a copy containing only unflagged mappings.
`--include-review` writes flagged entries with a visible `PRUEFEN` marker.
The copied station file gets a
unique name under `out/`; the template is not modified. A matching
`*_schreibprotokoll.csv` records its changes and sources.

For scanned documents, add `--ocr`. For optional local semantic search, add
`--embed`; the first embedded run downloads a model. Without it, BM25 search
works entirely offline after dependencies have been installed.

Start the local browser app with:

```sh
python -m uvicorn serve:app --host 127.0.0.1 --port 8800
```

Open `http://127.0.0.1:8800` and select **Station file**. Enter a station name,
optionally select source documents and a `.pptx` template, then click **Build
draft mapping**. If no files are selected, the app uses `sources/` and the
local example template at `out/vorlage_stationsakte.pptx`. Uploaded source files
are stored under ignored `sources/_uploads/`, templates under `out/uploads/`,
and both are processed locally. Download the review sheet, enter `ja` or `nein`
in the `Freigabe` column, then upload it using **Use reviewed sheet**. By
default, **Generate PowerPoint** writes only unflagged or approved entries.
You can explicitly include entries that still need review; they carry a
`PRUEFEN` marker. Download and inspect the finished slide deck. Text edits in
the spreadsheet are not imported. Each new draft needs its own reviewed sheet.
Uploaded source files and templates remain on disk to support source previews;
remove old runs from `sources/_uploads/` and `out/uploads/` when no longer
needed. Failed uploads are cleaned up automatically.

For live frontend development, install Node.js 20+ and run `npm install` once
inside `frontend/`. Then run `python dev.py` from the repository root using
the virtual environment Python. It starts FastAPI on port 8800 and Vite on
port 5173. Open `http://127.0.0.1:5173` for the live frontend. The frontend
proxies `/api` to the backend. `npm run build` in `frontend/` builds the same
dashboard for FastAPI to serve at `http://127.0.0.1:8800`.

## Station sheet: what an engineer actually reads

`out/stationsblatt.xlsx` is written on every run next to the review sheet. It
answers "what applies to *my* station?" instead of listing quotes from the whole
line. Every sentence is rated by rules in `stationen.py` (no model), and every
rating states its reason:

| Rating | Meaning |
|---|---|
| Station | names the station number or a process of this station (e.g. stripping) |
| Anlagenweit | applies to all stations (safety, capability, documentation, acceptance, ...) |
| Andere Station | names another station or process, listed only as a cross-check |
| Unklar | no hint at all, a person decides |

Sheets: `Fähigkeiten` (process / Cgk / tolerance as one table, duplicates merged),
`Diese Station`, `Anlagenweit`, `Unklar`, `Andere Stationen`, and a gap list on
`Übersicht` (fields with nothing found). The German short text is a keyword note
built from the binding level, the topic and the values found; the quoted sentence
with file and page stays the evidence. The last column `Stimmt?` (ja/nein) is for
feedback and shows which rules need tuning.

With a placeholder station such as "Gesamtanlage (Stationszuordnung folgt)" nothing
is rated; sentences are sorted by topic and `out/stationsliste_vorschlag.csv`
proposes stations from the processes found. Fill in numbers and names, then:

```sh
python run.py --template path/to/Stationsakte.pptx --station "Station 20 Abisolieren" --stationsliste stationsliste.csv
```

This rates each sentence against all stations and writes one sheet per station to
`out/stationsblaetter/`. See `stationsliste_beispiel.csv`. Rows rated "Andere
Station" or "Unklar" are marked for review, so the default PowerPoint skips them.
In the web app, upload the list under **Station list**.

Version comparison for one station:
`python ketten.py --station "Station 20 Abisolieren" --stationsliste stationsliste.csv`
puts changed values for that station first, with old and new value.

## Limits that need human input

- The source document may describe an entire line. The program cannot infer
  which station owns each requirement, so `--station` must come from a person.
- Mappings and extracted requirements are candidates, not a completeness
  guarantee. Check the discarded and unassigned rows before relying on an Akte.
- The supplied example does not include a real process specification or meeting
  minutes. Those sources have not been validated end to end.
- The current writer creates an initial station file. It does not merge edits
  from a later PDF revision into a previously reviewed Akte.

Run `python requirements.py --self-test` and `python router.py --check` for
quick checks. `python ocr.py --list` shows which documents need OCR.
