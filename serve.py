#!/usr/bin/env python3
"""
Local API for the Gluth RFQ prototype. Everything runs on this machine.

    pipeline/.venv/bin/python -m uvicorn pipeline.serve:app --port 8800 --reload

Endpoints
    GET  /api/health              index status, tiers available
    GET  /api/search?q=...        precedent search, grouped by project
    GET  /api/norms               referenced norms (regex, deterministic)
    GET  /api/norms?drift=true    norms cited at conflicting revisions
    GET  /api/manifest            extraction coverage stats
    GET  /                        serves the prototype UI if present
"""
from __future__ import annotations

import json
import asyncio
import subprocess
import sys
import threading
import uuid
from collections import Counter
from pathlib import Path

from fastapi import Body, FastAPI, File, Form, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import artifacts as artifacts_mod
import drift as drift_mod
import norms as norms_mod
import search as search_mod
import specdiff as specdiff_mod
import store as store_mod

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"
SOURCES = HERE / "sources"
WORKFLOW_FILE = OUT / "station_workflow.json"
WORKFLOW_LOCK = threading.Lock()
UI_FILE = HERE / "ui.html"
FRONTEND_DIST = HERE / "frontend" / "dist"
VENDOR = HERE / "vendor"          # local copies of CSS/JS — no internet needed

app = FastAPI(title="Gluth RFQ — local retrieval API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://127.0.0.1:5173", "http://localhost:5173"],
    allow_methods=["*"], allow_headers=["*"],
)

VENDOR.mkdir(parents=True, exist_ok=True)
app.mount("/vendor", StaticFiles(directory=VENDOR), name="vendor")

DB = store_mod.connect()
store_mod.seed_from_corpus(DB, search_mod.load_chunks())
store_mod.refresh_routing(DB)


@app.get("/api/health")
def health():
    chunks = search_mod.load_chunks()
    has_vec = (OUT / "vectors.npy").exists()
    projects = sorted({c["project_no"] for c in chunks if c["project_no"]})
    return {
        "status": "ok" if chunks and (OUT / "bm25").exists() else "setup_required",
        "chunks": len(chunks),
        "projects": projects,
        "tier0_bm25": (OUT / "bm25").exists(),
        "tier1_dense": has_vec,
        "mode": "hybrid" if has_vec else "bm25",
        "llm_used": False,
    }


@app.get("/api/search")
def api_search(
    q: str = Query(..., min_length=2),
    mode: str = Query("auto", pattern="^(auto|bm25|dense|hybrid|rerank)$"),
    limit: int = Query(5, ge=1, le=20),
    customer: str | None = None,
    project: str | None = None,
    phase: str | None = None,
):
    if not search_mod.load_chunks() or not (OUT / "bm25").exists():
        return JSONResponse({"error": "run.py must build the corpus and index first"},
                            status_code=503)
    return search_mod.search(q, mode=mode, project_limit=limit,
                             customer=customer, project=project, phase=phase)


@app.get("/api/norms")
def api_norms(
    project: str | None = None,
    drift: bool = False,
    min_hits: int = Query(2, ge=1),
):
    found = norms_mod.extract(search_mod.load_chunks(), project)
    rows = []
    for norm_id, by_project in found.items():
        total = sum(len(v) for v in by_project.values())
        if total < min_hits:
            continue
        revs = sorted({h["revision"] for hits in by_project.values()
                       for h in hits if h["revision"]})
        if drift and len(revs) < 2:
            continue
        rows.append({
            "norm_id": norm_id,
            "refs": total,
            "revisions": revs,
            "conflict": len(revs) > 1,
            "projects": {
                p: {
                    "refs": len(h),
                    "files": sorted({x["file"] for x in h}),
                    "example_page": h[0]["page"],
                }
                for p, h in by_project.items()
            },
        })
    rows.sort(key=lambda r: -r["refs"])
    return {"count": len(rows), "drift_only": drift, "norms": rows}


@app.get("/api/manifest")
def api_manifest():
    path = OUT / "manifest.json"
    if not path.exists():
        return JSONResponse({"error": "run extract.py first"}, status_code=404)
    manifest = json.loads(path.read_text(encoding="utf-8"))
    content = [r for r in manifest
               if r["status"] not in ("skipped", "out_of_scope")]
    ok = [r for r in content if r["status"] == "ok"]

    by_ext: dict[str, dict] = {}
    for r in content:
        d = by_ext.setdefault(r["ext"] or "(none)", {"ok": 0, "total": 0})
        d["total"] += 1
        d["ok"] += r["status"] == "ok"

    return {
        "files_scanned": len(manifest),
        "content_files": len(content),
        "extracted_ok": len(ok),
        "extraction_rate": round(100 * len(ok) / max(len(content), 1)),
        "total_chars": sum(r["chars"] for r in manifest),
        "by_status": dict(Counter(r["status"] for r in manifest)),
        "by_format": {
            k: {**v, "rate": round(100 * v["ok"] / v["total"])}
            for k, v in sorted(by_ext.items(), key=lambda x: -x[1]["total"])
        },
        "needs_ocr": [
            {"file": r["filename"], "reason": r["reason"], "path": r["source_path"]}
            for r in manifest if r["status"] == "needs_ocr"
        ],
    }


# ------------------------------------------------------------ queue / triage
@app.get("/api/inquiries")
def api_inquiries():
    return {"inquiries": store_mod.list_inquiries(DB)}


@app.get("/api/inquiry/{project_no}")
def api_inquiry(project_no: str):
    rows = [i for i in store_mod.list_inquiries(DB)
            if i["project_no"] == project_no]
    if not rows:
        return JSONResponse({"error": "unknown project"}, status_code=404)
    return rows[0]


@app.post("/api/inquiry")
def api_add_inquiry(payload: dict = Body(...)):
    """
    Record a project by hand. Creates a row only — no document is read.
    Documents enter the system exclusively through extract.py.
    """
    try:
        created = store_mod.add_inquiry(
            DB,
            project_no=payload.get("project_no", ""),
            customer=payload.get("customer", ""),
            project_name=payload.get("project_name", ""),
        )
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    if not created:
        return JSONResponse({"error": "project_exists"}, status_code=409)
    return {"ok": True, "project_no": payload.get("project_no")}


@app.post("/api/decision")
def api_decision(payload: dict = Body(...)):
    store_mod.set_decision(
        DB,
        project_no=payload.get("project_no", ""),
        verdict=payload.get("verdict", ""),
        criteria=payload.get("criteria", []),
        assigned_to=payload.get("assigned_to", ""),
        note=payload.get("note", ""),
    )
    return {"ok": True, "project_no": payload.get("project_no")}


# ---------------------------------------------------------------- artifacts
@app.get("/api/artifact/{project_no}")
def api_artifact(project_no: str,
                 kind: str = Query("summary",
                                   pattern="^(summary|checkliste|uebersicht)$")):
    try:
        return artifacts_mod.build(project_no, kind)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)


@app.get("/api/source")
def api_source(path: str, page: str = ""):
    """Return the real extracted text for a citation — the click-through target."""
    hits = [c for c in search_mod.load_chunks() if c["source_path"] == path]
    if not hits:
        return JSONResponse({"error": "not in corpus"}, status_code=404)
    if page:
        page_hits = [c for c in hits if str(c["page"]) == str(page)]
        if page_hits:
            hits = page_hits
    hits = sorted(hits, key=lambda c: c["id"])[:6]
    return {
        "source_path": path,
        "filename": hits[0]["filename"],
        "page": page,
        "project_no": hits[0]["project_no"],
        "phase": f'{hits[0]["phase_no"]}_{hits[0]["phase_name"]}'
                 if hits[0].get("phase_no") else "",
        "passages": [c["text"] for c in hits],
    }


# ------------------------------------------------------------- review layer
@app.get("/api/corrections")
def api_corrections(route_tag: str = "ALL"):
    return {
        "corrections": store_mod.list_corrections(DB, route_tag),
        "stats": store_mod.review_stats(DB),
    }


@app.post("/api/correction")
def api_add_correction(payload: dict = Body(...)):
    new_id = store_mod.add_correction(DB, **payload)
    return {"ok": True, "id": new_id, "stats": store_mod.review_stats(DB)}


@app.post("/api/approve")
def api_approve(payload: dict = Body(...)):
    store_mod.approve(DB, payload.get("project_no", ""),
                      payload.get("artifact", ""),
                      payload.get("reviewer", "M. Weber"))
    return {"ok": True, "stats": store_mod.review_stats(DB)}


# ---------------------------------------------------------------- dashboard
@app.get("/api/stats")
def api_stats():
    chunks = search_mod.load_chunks()
    man = api_manifest()
    if isinstance(man, JSONResponse):
        man = {}
    rs = store_mod.review_stats(DB)
    inquiries = store_mod.list_inquiries(DB)

    cited_total, cited_ok = 0, 0
    conf = Counter()
    for inq in inquiries:
        try:
            doc = artifacts_mod.build(inq["project_no"], "summary")
        except Exception:
            continue
        cited_total += doc["item_count"]
        cited_ok += round(doc["citation_rate"] * doc["item_count"] / 100)
        conf[doc["confidence"]] += 1

    return {
        "corpus": {
            "chunks": len(chunks),
            "projects": len(inquiries),
            "extraction_rate": man.get("extraction_rate", 0),
            "content_files": man.get("content_files", 0),
            "extracted_ok": man.get("extracted_ok", 0),
            "needs_ocr": len(man.get("needs_ocr", [])),
            "total_chars": man.get("total_chars", 0),
        },
        "quality": {
            "citation_rate": round(100 * cited_ok / cited_total) if cited_total else 0,
            "llm_used": False,
            "mode": "hybrid" if (OUT / "vectors.npy").exists() else "bm25",
        },
        "review": rs,
        "confidence_mix": dict(conf),
        "by_format": man.get("by_format", {}),
        "needs_ocr_list": man.get("needs_ocr", [])[:20],
    }


# -------------------------------------------------------------------- drift
@app.get("/api/drift")
def api_drift(project: str | None = None, kind: str | None = None):
    findings = drift_mod.analyse(search_mod.load_chunks(), project)
    if kind:
        findings = [f for f in findings if f["kind"] == kind]
    return {
        "count": len(findings),
        "by_kind": dict(Counter(f["kind"] for f in findings)),
        "findings": findings,
    }


# ----------------------------------------------------------------- specdiff
@app.get("/api/specdiff/pairs")
def api_specdiff_pairs():
    docs = specdiff_mod.load_docs()
    return {"pairs": [{"a": a, "b": b} for a, b in specdiff_mod.find_pairs(docs)]}


@app.get("/api/specdiff")
def api_specdiff(a: str, b: str, limit: int = 25, numeric: bool = False):
    docs = specdiff_mod.load_docs()
    try:
        d = specdiff_mod.diff(docs, a, b)
    except SystemExit as e:
        return JSONResponse({"error": str(e)}, status_code=404)
    if numeric:
        # Only changes where a number actually differs — cuts renumbering noise.
        d["changed"] = [c for c in d["changed"] if specdiff_mod.numeric_change(c)]
        d["added"] = [s for s in d["added"] if specdiff_mod.NUM_RE.search(s)]
        d["removed"] = [s for s in d["removed"] if specdiff_mod.NUM_RE.search(s)]
    return {
        "a": d["a"], "b": d["b"],
        "a_lines": d["a_lines"], "b_lines": d["b_lines"],
        "similarity": round(d["similarity"] * 100, 1),
        "counts": {"added": len(d["added"]), "removed": len(d["removed"]),
                   "changed": len(d["changed"])},
        "added": d["added"][:limit],
        "removed": d["removed"][:limit],
        # cap each side — a single replace block can span dozens of lines
        "changed": [{"from": o[:3], "to": n[:3]}
                    for o, n in d["changed"][:limit]],
    }


@app.get("/api/station")
def api_station():
    """Read the latest locally generated mapping and station output."""
    mapping = OUT / "zuordnung.jsonl"
    rows = []
    if mapping.exists():
        with mapping.open(encoding="utf-8") as stream:
            rows = [json.loads(line) for line in stream if line.strip()]
    workflow = json.loads(WORKFLOW_FILE.read_text(encoding="utf-8")) if WORKFLOW_FILE.exists() else None
    outputs = sorted(OUT.glob("Stationsakte_*.pptx"),
                     key=lambda path: path.stat().st_mtime, reverse=True)
    latest = (OUT / workflow["output"] if workflow and workflow.get("output") else None)
    if workflow is None:
        latest = next((p for p in outputs if p.name != "Stationsakte_gefuellt.pptx"),
                      outputs[0] if outputs else None)
    return {
        "count": len(rows),
        "review_count": sum(bool(row.get("pruefen")) for row in rows),
        "station": rows[0].get("station", "") if rows else "",
        "mappings": rows,
        "has_review_sheet": (OUT / "zuordnung.xlsx").exists(),
        "latest_output": latest.name if latest and latest.is_file() else None,
        "draft_ready": WORKFLOW_FILE.exists(),
    }


def _run_stage(*args: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            [sys.executable, *args], cwd=HERE, capture_output=True,
            text=True, errors="replace", timeout=600,
        )
    except subprocess.TimeoutExpired:
        return False, "Generation timed out after 10 minutes."
    output = (result.stdout + "\n" + result.stderr).strip()
    return result.returncode == 0, output[-6000:]


async def _save_upload(upload: UploadFile, destination: Path, limit: int) -> bool:
    destination.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    with destination.open("wb") as stream:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > limit:
                destination.unlink(missing_ok=True)
                return False
            stream.write(chunk)
    return size > 0


@app.post("/api/station/draft")
async def api_station_draft(
    request: Request,
    station: str = Form(...),
    file_filter: str = Form(""),
    template: UploadFile | None = File(None),
    documents: list[UploadFile] | None = File(None),
):
    """Create a review mapping from uploaded files or the existing local corpus."""
    if request.headers.get("origin") not in (None, "http://127.0.0.1:5173", "http://localhost:5173", "http://127.0.0.1:8800", "http://localhost:8800"):
        return JSONResponse({"error": "This local operation must come from the Gluth app."}, status_code=403)
    station = station.strip()
    if not station or len(station) > 120:
        return JSONResponse({"error": "Enter a station name (up to 120 characters)."}, status_code=400)
    if not WORKFLOW_LOCK.acquire(blocking=False):
        return JSONResponse({"error": "A station job is already running."}, status_code=409)
    try:
        uploads = documents or []
        if len(uploads) > 20:
            return JSONResponse({"error": "Upload at most 20 documents at a time."}, status_code=400)
        run_id = uuid.uuid4().hex
        template_path = OUT / "uploads" / run_id / "template.pptx"
        if template:
            if Path(template.filename or "").suffix.lower() != ".pptx":
                return JSONResponse({"error": "The template must be a .pptx file."}, status_code=400)
            if not await _save_upload(template, template_path, 30 * 1024 * 1024):
                return JSONResponse({"error": "Template is empty or larger than 30 MB."}, status_code=400)
        else:
            template_path = OUT / "vorlage_stationsakte.pptx"
            if not template_path.is_file():
                return JSONResponse({"error": "Choose a PowerPoint template."}, status_code=400)
        input_path = SOURCES
        if uploads:
            input_path = SOURCES / "_uploads" / run_id
            allowed = {".pdf", ".docx", ".doc", ".txt", ".md", ".xlsx", ".xlsm", ".pptx"}
            names: set[str] = set()
            for upload in uploads:
                name = Path(upload.filename or "").name
                if not name or name in names or Path(name).suffix.lower() not in allowed:
                    return JSONResponse({"error": "Use unique PDF, Word, text, spreadsheet, or PowerPoint source filenames."}, status_code=400)
                names.add(name)
                if not await _save_upload(upload, input_path / name, 30 * 1024 * 1024):
                    return JSONResponse({"error": f"{name} is empty or larger than 30 MB."}, status_code=400)
        elif not SOURCES.is_dir() or not any(path.is_file() for path in SOURCES.rglob("*")):
            return JSONResponse({"error": "Upload source documents first."}, status_code=400)
        WORKFLOW_FILE.unlink(missing_ok=True)
        args = [str(HERE / "run.py"), "--input", str(input_path),
                "--template", str(template_path), "--station", station]
        if file_filter.strip() and not uploads:
            args += ["--file", file_filter.strip()[:120]]
        ok, log = await asyncio.to_thread(_run_stage, *args)
        if not ok:
            return JSONResponse({"error": "Draft generation failed.", "log": log}, status_code=500)
        WORKFLOW_FILE.write_text(json.dumps({"station": station, "template": str(template_path)}), encoding="utf-8")
        return {"ok": True, "log": log, "station": api_station()}
    finally:
        WORKFLOW_LOCK.release()


@app.post("/api/station/presentation")
def api_station_presentation(request: Request):
    """Write safe mapped entries to a fresh template copy."""
    if request.headers.get("origin") not in (None, "http://127.0.0.1:5173", "http://localhost:5173", "http://127.0.0.1:8800", "http://localhost:8800"):
        return JSONResponse({"error": "This local operation must come from the Gluth app."}, status_code=403)
    if not WORKFLOW_LOCK.acquire(blocking=False):
        return JSONResponse({"error": "A station job is already running."}, status_code=409)
    try:
        if not WORKFLOW_FILE.exists():
            return JSONResponse({"error": "Build a draft mapping first."}, status_code=400)
        workflow = json.loads(WORKFLOW_FILE.read_text(encoding="utf-8"))
        template = Path(workflow["template"])
        if not template.is_file():
            return JSONResponse({"error": "The draft template is missing. Build the mapping again."}, status_code=400)
        output = OUT / f"Stationsakte_{uuid.uuid4().hex}.pptx"
        ok, log = _run_stage(str(HERE / "akte_schreiben.py"),
                             "--vorlage", str(template), "--station", workflow["station"],
                             "--aus", str(output), "--nur-sichere")
        if not ok:
            return JSONResponse({"error": "PowerPoint generation failed.", "log": log}, status_code=500)
        workflow["output"] = output.name
        WORKFLOW_FILE.write_text(json.dumps(workflow), encoding="utf-8")
        return {"ok": True, "filename": output.name, "log": log, "station": api_station()}
    finally:
        WORKFLOW_LOCK.release()


@app.get("/api/station/download/{kind}")
def api_station_download(kind: str):
    if kind == "review":
        path = OUT / "zuordnung.xlsx"
    elif kind == "presentation":
        workflow = json.loads(WORKFLOW_FILE.read_text(encoding="utf-8")) if WORKFLOW_FILE.exists() else None
        if workflow:
            path = OUT / workflow.get("output", "missing")
        else:
            outputs = sorted(OUT.glob("Stationsakte_*.pptx"),
                             key=lambda p: p.stat().st_mtime, reverse=True)
            path = outputs[0] if outputs else OUT / "missing"
    else:
        return JSONResponse({"error": "unknown download"}, status_code=404)
    if not path.is_file():
        return JSONResponse({"error": "file not generated yet"}, status_code=404)
    return FileResponse(path, filename=path.name)


@app.get("/")
def root():
    built = FRONTEND_DIST / "index.html"
    if built.exists():
        return FileResponse(built, headers={"Cache-Control": "no-store"})
    if UI_FILE.exists():
        # never cache the UI — otherwise edits silently don't appear
        return FileResponse(UI_FILE, headers={
            "Cache-Control": "no-store, no-cache, must-revalidate",
            "Pragma": "no-cache",
        })
    return {"hint": "drop the prototype at pipeline/ui.html, or use /api/*"}


if FRONTEND_DIST.exists():
    app.mount("/assets", StaticFiles(directory=FRONTEND_DIST / "assets"),
              name="frontend-assets")
