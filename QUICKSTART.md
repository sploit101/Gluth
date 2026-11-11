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

## Station sheet and slides: what an engineer actually reads

Every run writes two files for the engineer next to the review sheet:

- `out/stationsblatt.xlsx`: the full working view with evidence.
- `out/stationsfolien.pptx`: the target-image slides to discuss: station overview,
  process requirements as a table, this station, changes, all stations, open points.

Both answer "what applies to *my* station?" instead of listing quotes from the
whole line. Every sentence is rated by rules in `stationen.py` (no model), and every
rating states its reason and a confidence (hoch / mittel / niedrig):

| Rating | Meaning |
|---|---|
| Station | names the station number, a process of this station (e.g. stripping), or stands in a chapter about it |
| Anlagenweit | applies to all stations (safety, capability, documentation, acceptance, ...) |
| Andere Station | names another station or process, listed only as a cross-check |
| Unklar | no reliable hint, a person decides |

**Chapters count.** `requirements.py` now records the numbered heading each
sentence stands under ("3 Process requirements › 3.2 Stripping unit"). A sentence
without a process word of its own ("The tolerance shall be proven per variant.")
inherits the chapter's process. A station number or process in the sentence itself
still wins over the chapter. The heading is shown in the `Kapitel` column.

Sheets: `Übersicht` (counts and gaps), `Fähigkeiten` (Cgk/Cpk/MSA as one table),
`Kennwerte` (parameter, condition, value, unit: cycle time, forces, tolerances, noise;
a number is only listed when a known parameter word precedes it), `Änderungen`,
`Diese Station`, `Anlagenweit`, `Unklar` and `Andere Stationen`. The German short
text ("Muss · Abisolieren · Abisolierlänge (Toleranz) ± 0,2 mm") is built from the
binding level, the topic and the values found. The quoted sentence with file, page
and chapter stays the evidence. Gaps say when a value is in the text but no field
took it. A vague word such as "feeding" alone never moves a sentence to another
station: it becomes Unklar, so nothing relevant disappears.

With a placeholder station such as "Gesamtanlage (Stationszuordnung folgt)" nothing
is rated; sentences are sorted by topic and `out/stationsliste_vorschlag.csv`
proposes stations from the processes found. Fill in numbers and names, then:

```sh
python run.py --template path/to/Stationsakte.pptx --station "Station 20 Abisolieren" --stationsliste stationsliste.csv
```

This rates each sentence against all stations and writes a sheet and slides per
station to `out/stationsblaetter/`. See `stationsliste_beispiel.csv`. A process that
no station in the list covers stays Unklar. Rows rated "Andere Station" or "Unklar"
are marked for review, so the default PowerPoint skips them. In the web app, upload
the list under **Station list**.

## A new version of the Lastenheft arrives

1. **Only the newest version goes into the file.** If `LH_..._V1.01` and `_V1.02`
   are both in `sources/`, only V1.02 is mapped; V1.01 is used for comparison
   (`--alle-fassungen` turns this off).
2. **What changed, for this station:** sheet `Änderungen` and the slide
   "Änderungen ggü. …" list NEU, ENTFALLEN, WERT GEÄNDERT ("Abisolierlänge
   (Toleranz): ± 0,3 mm → ± 0,2 mm"), VERBINDLICHKEIT GEÄNDERT (SOLL → MUSS) and
   TEXT GEÄNDERT with a word diff, sorted with this station first. Standalone:
   `python aenderungen.py --datei LH_04K_000051 --station "Station 20 Abisolieren"`.
3. **Only review what changed:** pass the reviewed sheet of the previous version:

```sh
python run.py --template path/to/Stationsakte.pptx --station "Station 20 Abisolieren" --vorherige-freigabe path/to/reviewed_V1.01.xlsx
```

   ja/nein is carried over for sentences that are unchanged in the same field for
   the same station (grey rows). Changed rows say "GEÄNDERT ggü. Vorfassung (dort:
   ja) – old → new"; new rows say "NEU". The web app does this automatically
   with the last uploaded review sheet.

## Feedback

After someone fills in `Stimmt?` (ja = rating correct) in the station sheet:
`python stationen.py --auswerten out/stationsblatt.xlsx` prints the hit rate per
sheet and confidence and writes the wrong rows to `*_falsch.csv`.

## Wording that changed meaning

"must not exceed 6 s" / "darf 32 s nicht überschreiten" is now a limit (MUSS,
≤ 6 s), not a prohibition (VERBOT). "must not be removed" stays VERBOT.

## Limits that need human input

- The source document may describe an entire line. The program cannot infer
  which station owns each requirement, so `--station` must come from a person.
- Mappings and extracted requirements are candidates, not a completeness
  guarantee. Check the discarded and unassigned rows before relying on an Akte.
- The supplied example does not include a real process specification or meeting
  minutes. Those sources have not been validated end to end.
- The writer creates a new station file per run. A later PDF revision is handled
  through the change list and carried-over approvals, not by editing a
  previously written PowerPoint in place.
- Chapter detection needs numbered headings on their own line ("3.2 Stripping
  unit"). Unnumbered Word headings are not detected yet.

Run `python requirements.py --self-test` and `python router.py --check` for
quick checks. `python ocr.py --list` shows which documents need OCR.
