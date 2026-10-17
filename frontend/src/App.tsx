import { useCallback, useEffect, useState, type FormEvent, type ReactNode } from "react"
import {
  ArrowDownToLineIcon, ArrowRightIcon, ArrowUpRightIcon, CheckCircle2Icon,
  CircleAlertIcon, Clock3Icon, DatabaseIcon, FileIcon, FileSearchIcon,
  FileTextIcon, FilesIcon, FolderKanbanIcon, GitCompareArrowsIcon, Layers3Icon,
  LoaderCircleIcon, RefreshCwIcon, SearchIcon, ShieldCheckIcon, UploadCloudIcon,
} from "lucide-react"
import { AppSidebar, type View } from "@/components/app-sidebar"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Separator } from "@/components/ui/separator"
import { SidebarInset, SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar"

type Health = { status: string; chunks: number; projects: string[]; tier0_bm25: boolean; tier1_dense: boolean; mode: string }
type Stats = { corpus: { chunks: number; projects: number; extraction_rate: number; content_files: number; extracted_ok: number; needs_ocr: number; total_chars: number }; quality: { citation_rate: number; mode: string }; review: { reviews: number; approved: number; corrections: number }; by_format: Record<string, { ok: number; total: number; rate: number }> }
type Manifest = { files_scanned: number; content_files: number; extracted_ok: number; extraction_rate: number; by_status: Record<string, number>; needs_ocr: { file: string; reason: string; path: string }[] }
type Inquiry = { project_no: string; project_name: string; customer: string; stage: string; status: string; doc_count: number; chunk_count: number; updated_at: string }
type Hit = { source_path: string; filename: string; phase: string; page: string; snippet: string }
type SearchResult = { project_no: string; project_name: string; customer: string; score: number; hits: Hit[] }
type SearchData = { query: string; mode: string; projects: SearchResult[] }
type SourceData = { filename: string; page: string; source_path: string; passages: string[] }
type Mapping = { folienart: string; feld_id: string; feldname: string; regel: string; pruefen: string; station: string; datei: string; seite: string; text: string; kennung: string }
type Station = { count: number; review_count: number; station: string; mappings: Mapping[]; has_review_sheet: boolean; latest_output: string | null; draft_ready: boolean }
type Drift = { count: number; findings: { kind: string; file?: string; detail?: string; description?: string }[] }
type Pair = { a: string; b: string }

async function api<T>(path: string): Promise<T> {
  const response = await fetch(`/api/${path}`, { cache: "no-store" })
  if (!response.ok) {
    const detail = await response.json().catch(() => ({})) as { error?: string }
    throw new Error(detail.error || `Request failed (${response.status})`)
  }
  return response.json() as Promise<T>
}

function Panel({ children, className = "" }: { children: ReactNode; className?: string }) {
  return <div className={`rounded-2xl border border-[#e5e9ee] bg-white shadow-[0_4px_20px_rgba(25,47,72,0.035)] ${className}`}>{children}</div>
}
function Eyebrow({ children }: { children: ReactNode }) {
  return <div className="text-[11px] font-semibold tracking-[0.16em] text-[#6d879f] uppercase">{children}</div>
}
function Empty({ title, detail }: { title: string; detail: string }) {
  return <div className="flex min-h-40 flex-col items-center justify-center px-5 text-center"><FileSearchIcon className="mb-3 size-6 text-[#9eb0be]" /><p className="font-medium text-[#2e4c67]">{title}</p><p className="mt-1 max-w-sm text-sm text-muted-foreground">{detail}</p></div>
}
function Metric({ icon, label, value, foot }: { icon: ReactNode; label: string; value: string | number; foot: string }) {
  return <Panel className="p-5"><div className="flex items-start justify-between"><div><p className="text-sm text-[#72879a]">{label}</p><p className="mt-3 text-3xl font-semibold tracking-tight text-[#193b59]">{value}</p></div><div className="rounded-xl bg-[#edf3f6] p-2.5 text-[#44718d]">{icon}</div></div><p className="mt-4 text-xs text-[#8294a3]">{foot}</p></Panel>
}
function SectionHead({ eyebrow, title, detail, action }: { eyebrow: string; title: string; detail: string; action?: ReactNode }) {
  return <div className="mb-6 flex flex-wrap items-end justify-between gap-4"><div><Eyebrow>{eyebrow}</Eyebrow><h1 className="mt-2 text-3xl font-semibold tracking-tight text-[#173a59]">{title}</h1><p className="mt-2 text-sm text-[#70869a]">{detail}</p></div>{action}</div>
}

export default function App() {
  const [view, setView] = useState<View>("overview")
  const [health, setHealth] = useState<Health | null>(null)
  const [stats, setStats] = useState<Stats | null>(null)
  const [manifest, setManifest] = useState<Manifest | null>(null)
  const [inquiries, setInquiries] = useState<Inquiry[]>([])
  const [station, setStation] = useState<Station | null>(null)
  const [drift, setDrift] = useState<Drift | null>(null)
  const [pairs, setPairs] = useState<Pair[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState("")
  const [updated, setUpdated] = useState<Date | null>(null)
  const [query, setQuery] = useState("")
  const [searchData, setSearchData] = useState<SearchData | null>(null)
  const [searching, setSearching] = useState(false)
  const [searchError, setSearchError] = useState("")
  const [source, setSource] = useState<SourceData | null>(null)
  const [sourceError, setSourceError] = useState("")
  const [selectedProject, setSelectedProject] = useState<string | null>(null)
  const [stationName, setStationName] = useState("")
  const [fileFilter, setFileFilter] = useState("")
  const [templateFile, setTemplateFile] = useState<File | null>(null)
  const [sourceFiles, setSourceFiles] = useState<File[]>([])
  const [stationBusy, setStationBusy] = useState(false)
  const [stationMessage, setStationMessage] = useState("")
  const [stationError, setStationError] = useState("")
  const [stationLog, setStationLog] = useState("")

  const refresh = useCallback(async () => {
    const results = await Promise.allSettled([
      api<Health>("health"), api<Stats>("stats"), api<Manifest>("manifest"),
      api<{ inquiries: Inquiry[] }>("inquiries"), api<Station>("station"),
      api<Drift>("drift"), api<{ pairs: Pair[] }>("specdiff/pairs"),
    ])
    const [h, s, m, i, st, d, p] = results
    if (h.status === "fulfilled") setHealth(h.value)
    if (s.status === "fulfilled") setStats(s.value)
    if (m.status === "fulfilled") setManifest(m.value)
    if (i.status === "fulfilled") setInquiries(i.value.inquiries)
    if (st.status === "fulfilled") setStation(st.value)
    if (d.status === "fulfilled") setDrift(d.value)
    if (p.status === "fulfilled") setPairs(p.value.pairs)
    const failures = results.filter((r) => r.status === "rejected")
    setError(failures.length === results.length ? "The local API is unavailable. Start the Python backend on port 8800." : "")
    setUpdated(new Date())
    setLoading(false)
  }, [])

  useEffect(() => {
    const initial = window.setTimeout(() => void refresh(), 0)
    const timer = window.setInterval(() => void refresh(), 15000)
    return () => { window.clearTimeout(initial); window.clearInterval(timer) }
  }, [refresh])

  async function submitSearch(event?: FormEvent) {
    event?.preventDefault()
    if (query.trim().length < 2) return
    setSearching(true)
    setSearchError("")
    setSource(null)
    try { setSearchData(await api<SearchData>(`search?q=${encodeURIComponent(query.trim())}&mode=bm25&limit=8`)) }
    catch (err) { setSearchError(err instanceof Error ? err.message : "Search failed") }
    finally { setSearching(false) }
  }

  async function showSource(hit: Hit) {
    setSource(null)
    setSourceError("")
    try { setSource(await api<SourceData>(`source?path=${encodeURIComponent(hit.source_path)}&page=${encodeURIComponent(hit.page)}`)) }
    catch (err) { setSourceError(err instanceof Error ? err.message : "Could not load source") }
  }

  async function buildDraft(event: FormEvent) {
    event.preventDefault()
    setStationBusy(true); setStationError(""); setStationMessage(""); setStationLog("")
    const form = new FormData()
    form.append("station", stationName.trim())
    form.append("file_filter", fileFilter.trim())
    if (templateFile) form.append("template", templateFile)
    sourceFiles.forEach(file => form.append("documents", file))
    try {
      const response = await fetch("/api/station/draft", { method: "POST", body: form })
      const data = await response.json() as { error?: string; log?: string; station?: Station }
      if (!response.ok) throw new Error(`${data.error || "Draft failed"}${data.log ? `\n${data.log}` : ""}`)
      if (data.station) setStation(data.station)
      setStationLog(data.log || "")
      setStationMessage(`Draft ready: ${data.station?.count ?? 0} mapped entries. Review the cited rows before creating the PowerPoint.`)
      await refresh()
    } catch (err) { setStationError(err instanceof Error ? err.message : "Draft failed") }
    finally { setStationBusy(false) }
  }

  async function generatePresentation() {
    setStationBusy(true); setStationError(""); setStationMessage(""); setStationLog("")
    try {
      const response = await fetch("/api/station/presentation", { method: "POST" })
      const data = await response.json() as { error?: string; log?: string; station?: Station }
      if (!response.ok) throw new Error(`${data.error || "Generation failed"}${data.log ? `\n${data.log}` : ""}`)
      if (data.station) setStation(data.station)
      setStationLog(data.log || "")
      setStationMessage("PowerPoint created. Download it above and inspect every slide before use.")
    } catch (err) { setStationError(err instanceof Error ? err.message : "Generation failed") }
    finally { setStationBusy(false) }
  }

  const selected = inquiries.find((item) => item.project_no === selectedProject)
  const viewTitle: Record<View, string> = { overview: "Overview", search: "Search corpus", projects: "Projects", station: "Station file", sources: "Source files", changes: "Changes" }
  return (
    <SidebarProvider>
      <AppSidebar active={view} onSelect={setView} projectCount={inquiries.length} status={health?.status || "offline"} />
      <SidebarInset className="min-w-0 bg-[#f7f9fb]">
        <header className="sticky top-0 z-20 flex h-16 items-center justify-between border-b border-[#e7ebef] bg-white/90 px-4 backdrop-blur md:px-8">
          <div className="flex items-center gap-3"><SidebarTrigger className="text-[#54728b]" /><Separator orientation="vertical" className="h-5" /><span className="text-sm font-medium text-[#33536d]">{viewTitle[view]}</span></div>
          <div className="flex items-center gap-3"><span className="hidden text-xs text-[#91a0ad] sm:inline">{updated ? `Updated ${updated.toLocaleTimeString()}` : "Connecting…"}</span><Button size="icon" variant="ghost" onClick={() => void refresh()} aria-label="Refresh data"><RefreshCwIcon className="size-4" /></Button><span className={`rounded-full px-3 py-1 text-xs font-medium ${health?.status === "ok" ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>{health?.status === "ok" ? "Live local" : "Setup needed"}</span></div>
        </header>
        <main className="mx-auto w-full max-w-[1500px] flex-1 px-5 py-8 md:px-9 md:py-10">
          {error && <div className="mb-6 flex items-center gap-2 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900"><CircleAlertIcon className="size-4" />{error}</div>}
          {loading && <div className="flex items-center gap-2 text-sm text-muted-foreground"><LoaderCircleIcon className="size-4 animate-spin" />Loading workspace…</div>}

          {view === "overview" && <>
            <SectionHead eyebrow="Workspace pulse" title="Engineering overview" detail="A live view of the local document corpus and station workflow." action={<Button variant="outline" onClick={() => setView("search")}>Explore documents <ArrowRightIcon className="ml-2 size-4" /></Button>} />
            <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
              <Metric icon={<FolderKanbanIcon className="size-5" />} label="Projects" value={stats?.corpus.projects ?? "—"} foot="Tracked in the local corpus" />
              <Metric icon={<Layers3Icon className="size-5" />} label="Indexed passages" value={stats?.corpus.chunks ?? "—"} foot={health?.tier1_dense ? "BM25 + semantic index" : "BM25 lexical index"} />
              <Metric icon={<CheckCircle2Icon className="size-5" />} label="Extraction coverage" value={stats ? `${stats.corpus.extraction_rate}%` : "—"} foot={`${stats?.corpus.extracted_ok ?? 0} of ${stats?.corpus.content_files ?? 0} source files`} />
              <Metric icon={<ShieldCheckIcon className="size-5" />} label="Cited results" value={stats ? `${stats.quality.citation_rate}%` : "—"} foot="Source traceability in generated items" />
            </div>
            <div className="mt-6 grid gap-6 xl:grid-cols-[1.6fr_1fr]">
              <Panel className="overflow-hidden"><div className="flex items-center justify-between border-b border-[#eef1f3] px-6 py-5"><div><h2 className="font-semibold text-[#244661]">Recent projects</h2><p className="mt-1 text-xs text-muted-foreground">Open a project to see its source coverage.</p></div><Button variant="ghost" size="sm" onClick={() => setView("projects")}>View all <ArrowRightIcon className="ml-1 size-4" /></Button></div>{inquiries.length ? <div className="divide-y divide-[#eef1f3]">{inquiries.slice(0, 5).map((item) => <button key={item.project_no} onClick={() => {setSelectedProject(item.project_no);setView("projects")}} className="flex w-full items-center justify-between gap-4 px-6 py-4 text-left transition hover:bg-[#f8fafb]"><div className="flex min-w-0 items-center gap-3"><div className="rounded-lg bg-[#eef4f7] p-2 text-[#4a7089]"><FolderKanbanIcon className="size-4" /></div><div className="min-w-0"><p className="truncate text-sm font-medium text-[#244661]">{item.project_name || item.project_no}</p><p className="text-xs text-muted-foreground">{item.project_no} · {item.stage || "Unassigned stage"}</p></div></div><div className="shrink-0 text-right"><p className="text-sm font-medium text-[#365a75]">{item.doc_count} files</p><p className="text-xs text-muted-foreground">{item.chunk_count} passages</p></div></button>)}</div> : <Empty title="No projects indexed" detail="Place source documents under sources/ and run the pipeline to populate this view." />}</Panel>
              <div className="grid gap-6"><Panel className="p-6"><div className="flex items-center justify-between"><h2 className="font-semibold text-[#244661]">Station file</h2><FileTextIcon className="size-5 text-[#82a0b2]" /></div><p className="mt-4 text-3xl font-semibold text-[#244661]">{station?.count ?? 0} <span className="text-sm font-normal text-muted-foreground">mapped entries</span></p><p className="mt-2 text-sm text-muted-foreground">{station?.station || "No station selected yet"}</p><div className="mt-5 flex items-center justify-between border-t pt-4 text-xs"><span className="text-amber-700">{station?.review_count ?? 0} require review</span><button onClick={() => setView("station")} className="font-medium text-[#315f81] hover:underline">Open mapping →</button></div></Panel><Panel className="p-6"><h2 className="font-semibold text-[#244661]">Pipeline status</h2><div className="mt-5 space-y-3 text-sm"><div className="flex justify-between"><span className="text-muted-foreground">Search index</span><span className="font-medium">{health?.tier0_bm25 ? "Ready" : "Missing"}</span></div><div className="flex justify-between"><span className="text-muted-foreground">OCR pending</span><span className="font-medium">{stats?.corpus.needs_ocr ?? 0}</span></div><div className="flex justify-between"><span className="text-muted-foreground">Drift findings</span><span className="font-medium">{drift?.count ?? 0}</span></div></div></Panel></div>
            </div>
          </>}

          {view === "search" && <>
            <SectionHead eyebrow="Evidence first" title="Search the corpus" detail="Find prior engineering work and open the exact source passage." />
            <Panel className="p-5"><form onSubmit={(e) => void submitSearch(e)} className="flex flex-col gap-3 sm:flex-row"><div className="relative flex-1"><SearchIcon className="absolute top-1/2 left-3 size-4 -translate-y-1/2 text-[#8aa0b0]" /><Input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Try ‘crimping process capability’" className="h-11 pl-10" /></div><Button type="submit" disabled={searching || query.trim().length < 2} className="h-11 bg-[#244f70] px-6 hover:bg-[#183b58]">{searching ? <LoaderCircleIcon className="mr-2 size-4 animate-spin" /> : <SearchIcon className="mr-2 size-4" />}Search</Button></form><p className="mt-3 text-xs text-muted-foreground">Local BM25 search · Results link to the original file and page.</p></Panel>
            {searchError && <p className="mt-5 text-sm text-red-700">{searchError}</p>}
            <div className="mt-6 grid gap-6 xl:grid-cols-[minmax(0,1.25fr)_minmax(340px,.75fr)]"><div className="space-y-4">{searchData ? (searchData.projects.length ? searchData.projects.map((project) => <Panel key={project.project_no} className="overflow-hidden"><div className="flex items-center justify-between border-b px-5 py-4"><div><p className="font-semibold text-[#244661]">{project.project_name || project.project_no}</p><p className="text-xs text-muted-foreground">{project.project_no} · {project.hits.length} cited passage{project.hits.length === 1 ? "" : "s"}</p></div><span className="rounded-md bg-[#eef5f8] px-2 py-1 text-xs font-medium text-[#547b95]">Match {project.score.toFixed(2)}</span></div><div className="divide-y">{project.hits.map((hit, idx) => <button key={`${hit.source_path}-${hit.page}-${idx}`} onClick={() => void showSource(hit)} className="w-full px-5 py-4 text-left hover:bg-[#f8fafb]"><div className="flex items-center gap-2 text-xs font-medium text-[#446e8a]"><FileIcon className="size-3.5" />{hit.filename}<span className="ml-auto text-muted-foreground">p. {hit.page || "—"}</span></div><p className="mt-2 line-clamp-3 text-sm leading-relaxed text-[#425c70]">{hit.snippet}</p></button>)}</div></Panel>) : <Panel><Empty title="No matching projects" detail="Try a different technical term or check that the index has been built." /></Panel>) : <Panel><Empty title="Search is ready" detail="Enter a process, material, requirement, or standard to find source passages." /></Panel>}</div><Panel className="h-fit min-h-60 p-5"><div className="mb-4 flex items-center gap-2 text-sm font-semibold text-[#244661]"><FileTextIcon className="size-4" />Source preview</div>{source ? <><p className="text-sm font-medium text-[#244661]">{source.filename}</p><p className="mt-1 text-xs text-muted-foreground">Page {source.page || "—"} · {source.source_path}</p><div className="mt-4 space-y-3 border-t pt-4">{source.passages.map((passage, i) => <p key={i} className="whitespace-pre-wrap text-sm leading-7 text-[#4b6275]">{passage}</p>)}</div></> : <p className="text-sm text-muted-foreground">Select a result to inspect its original extracted text.</p>}{sourceError && <p className="mt-3 text-sm text-red-700">{sourceError}</p>}</Panel></div>
          </>}

          {view === "projects" && <>
            <SectionHead eyebrow="Document portfolio" title="Projects" detail="Projects are derived from the source folder structure and refreshed from the local API." />
            <div className="grid gap-6 xl:grid-cols-[minmax(0,1.1fr)_minmax(340px,.9fr)]"><Panel className="overflow-hidden">{inquiries.length ? <div className="divide-y">{inquiries.map((item) => <button key={item.project_no} onClick={() => setSelectedProject(item.project_no)} className={`flex w-full items-center justify-between gap-4 p-5 text-left hover:bg-[#f8fafb] ${selectedProject === item.project_no ? "bg-[#f2f7fa]" : ""}`}><div className="flex min-w-0 items-center gap-3"><div className="rounded-xl bg-[#eef4f7] p-3 text-[#4e7690]"><FolderKanbanIcon className="size-5" /></div><div className="min-w-0"><p className="truncate font-semibold text-[#244661]">{item.project_name || item.project_no}</p><p className="text-xs text-muted-foreground">{item.project_no} · {item.customer || "Customer not set"}</p></div></div><ArrowRightIcon className="size-4 shrink-0 text-[#8ca2b1]" /></button>)}</div> : <Empty title="No projects" detail="Run the pipeline with documents inside a project folder." />}</Panel><Panel className="h-fit p-6">{selected ? <><Eyebrow>Project detail</Eyebrow><h2 className="mt-2 text-xl font-semibold text-[#244661]">{selected.project_name || selected.project_no}</h2><p className="mt-1 text-sm text-muted-foreground">{selected.project_no}</p><div className="mt-6 grid grid-cols-2 gap-4"><div className="rounded-xl bg-[#f4f7f9] p-4"><p className="text-xs text-muted-foreground">Documents</p><p className="mt-2 text-2xl font-semibold text-[#244661]">{selected.doc_count}</p></div><div className="rounded-xl bg-[#f4f7f9] p-4"><p className="text-xs text-muted-foreground">Passages</p><p className="mt-2 text-2xl font-semibold text-[#244661]">{selected.chunk_count}</p></div></div><div className="mt-6 space-y-3 border-t pt-5 text-sm"><div className="flex justify-between"><span className="text-muted-foreground">Stage</span><span>{selected.stage || "—"}</span></div><div className="flex justify-between"><span className="text-muted-foreground">Status</span><span>{selected.status}</span></div></div><Button className="mt-6 w-full bg-[#244f70] hover:bg-[#183b58]" onClick={() => {setQuery(selected.project_name || selected.project_no);setView("search")}}>Search this project <ArrowRightIcon className="ml-2 size-4" /></Button></> : <Empty title="Select a project" detail="Choose a project to inspect its coverage and status." />}</Panel></div>
          </>}

          {view === "station" && <>
            <SectionHead eyebrow="Traceable drafting" title="Station file" detail="Each proposed field links to a source document and page. Review flagged entries before use." action={<div className="flex gap-2">{station?.has_review_sheet && <Button variant="outline" asChild><a href="/api/station/download/review"><ArrowDownToLineIcon className="mr-2 size-4" />Review sheet</a></Button>}{station?.latest_output && <Button className="bg-[#244f70] hover:bg-[#183b58]" asChild><a href="/api/station/download/presentation"><ArrowDownToLineIcon className="mr-2 size-4" />PowerPoint</a></Button>}</div>} />
            <Panel className="mb-6 p-6"><div className="flex items-start gap-3"><div className="rounded-xl bg-[#edf3f6] p-2.5 text-[#44718d]"><UploadCloudIcon className="size-5" /></div><div><h2 className="font-semibold text-[#244661]">Generate from this page</h2><p className="mt-1 text-sm text-muted-foreground">Choose documents and a PowerPoint template, or use the files already in this local workspace.</p></div></div><form onSubmit={buildDraft} className="mt-5 grid gap-4 lg:grid-cols-2"><label className="text-sm font-medium text-[#36566d]">Station name <Input className="mt-2" required maxLength={120} value={stationName} onChange={event => setStationName(event.target.value)} placeholder="Station 20 Abisolieren" /></label><label className="text-sm font-medium text-[#36566d]">Filename filter <span className="font-normal text-muted-foreground">(optional, existing files only)</span><Input className="mt-2" value={fileFilter} onChange={event => setFileFilter(event.target.value)} placeholder="Lastenheft" /></label><label className="text-sm font-medium text-[#36566d]">Source documents <span className="font-normal text-muted-foreground">(optional)</span><Input className="mt-2 h-auto py-2" type="file" multiple accept=".pdf,.docx,.doc,.txt,.md,.xlsx,.xlsm,.pptx" onChange={event => setSourceFiles(Array.from(event.target.files || []))} /><span className="mt-1 block text-xs font-normal text-muted-foreground">If empty, uses files under sources/. Up to 20 files, 30 MB each.</span></label><label className="text-sm font-medium text-[#36566d]">Station template <span className="font-normal text-muted-foreground">(.pptx)</span><Input className="mt-2 h-auto py-2" type="file" accept=".pptx" onChange={event => setTemplateFile(event.target.files?.[0] || null)} /><span className="mt-1 block text-xs font-normal text-muted-foreground">If empty, uses the local example template when available.</span></label><div className="flex flex-wrap items-center gap-3 lg:col-span-2"><Button type="submit" disabled={stationBusy || !stationName.trim()} className="bg-[#244f70] hover:bg-[#183b58]">{stationBusy ? <LoaderCircleIcon className="mr-2 size-4 animate-spin" /> : <Layers3Icon className="mr-2 size-4" />}Build draft mapping</Button><Button type="button" variant="outline" disabled={stationBusy || !station?.draft_ready || station.count === station.review_count} onClick={() => void generatePresentation()}>{stationBusy ? <LoaderCircleIcon className="mr-2 size-4 animate-spin" /> : <FileTextIcon className="mr-2 size-4" />}Generate PowerPoint</Button><span className="text-xs text-muted-foreground">Only entries without review flags are written.</span></div></form>{stationMessage && <p role="status" className="mt-4 rounded-lg bg-emerald-50 p-3 text-sm text-emerald-800">{stationMessage}</p>}{stationError && <pre role="alert" className="mt-4 max-h-40 overflow-auto whitespace-pre-wrap rounded-lg bg-red-50 p-3 text-xs text-red-800">{stationError}</pre>}{stationLog && <details className="mt-3 text-xs text-muted-foreground"><summary className="cursor-pointer">Run log</summary><pre className="mt-2 max-h-40 overflow-auto whitespace-pre-wrap">{stationLog}</pre></details>}</Panel>
            <div className="mb-6 grid gap-4 sm:grid-cols-3"><Metric icon={<FilesIcon className="size-5" />} label="Mapped entries" value={station?.count ?? 0} foot="Proposed field placements" /><Metric icon={<CircleAlertIcon className="size-5" />} label="Needs review" value={station?.review_count ?? 0} foot="Flagged by mapping rules" /><Metric icon={<FileTextIcon className="size-5" />} label="Current station" value={station?.station || "—"} foot="Supplied by a human" /></div>
            <Panel className="overflow-hidden"><div className="border-b px-6 py-5"><h2 className="font-semibold text-[#244661]">Field mapping</h2><p className="mt-1 text-xs text-muted-foreground">Latest generated mapping · Inspect the spreadsheet before use. Spreadsheet edits are not imported automatically.</p></div>{station?.mappings.length ? <div className="divide-y">{station.mappings.map((row, index) => <div key={`${row.kennung}-${index}`} className="grid gap-3 px-6 py-5 lg:grid-cols-[160px_minmax(0,1fr)_180px]"><div><span className="rounded-md bg-[#eaf2f6] px-2 py-1 font-mono text-xs text-[#416d88]">{row.kennung}</span><p className="mt-3 text-xs text-muted-foreground">{row.folienart}</p></div><div><p className="text-sm font-semibold text-[#284960]">{row.feldname}</p><p className="mt-2 text-sm leading-relaxed text-[#51697c]">{row.text}</p><p className="mt-3 text-xs text-[#6d879b]">{row.datei} · page {row.seite || "—"}</p></div><div className="lg:text-right"><span className={`inline-flex rounded-full px-2.5 py-1 text-xs font-medium ${row.pruefen ? "bg-amber-50 text-amber-700" : "bg-emerald-50 text-emerald-700"}`}>{row.pruefen ? "Review required" : "Rule checked"}</span><p className="mt-3 text-xs text-muted-foreground">{row.regel}</p></div></div>)}</div> : <Empty title="No mapping generated" detail="Run the local pipeline with a station template and a human-supplied station name." />}</Panel>
          </>}

          {view === "sources" && <>
            <SectionHead eyebrow="Corpus integrity" title="Source files" detail="Extraction coverage and OCR needs, read directly from the latest manifest." />
            <div className="grid gap-4 sm:grid-cols-3"><Metric icon={<DatabaseIcon className="size-5" />} label="Files scanned" value={manifest?.files_scanned ?? "—"} foot="Documents in the input folder" /><Metric icon={<CheckCircle2Icon className="size-5" />} label="Extracted" value={manifest?.extracted_ok ?? "—"} foot={`${manifest?.extraction_rate ?? 0}% extraction rate`} /><Metric icon={<Clock3Icon className="size-5" />} label="OCR needed" value={manifest?.needs_ocr.length ?? "—"} foot="Scanned or image-only sources" /></div>
            <div className="mt-6 grid gap-6 lg:grid-cols-2"><Panel className="p-6"><h2 className="font-semibold text-[#244661]">Extraction by status</h2><div className="mt-5 space-y-4">{Object.entries(manifest?.by_status || {}).map(([status, count]) => <div key={status}><div className="mb-2 flex justify-between text-sm"><span className="capitalize text-[#51697c]">{status.replaceAll("_", " ")}</span><span className="font-medium text-[#244661]">{count}</span></div><div className="h-2 rounded-full bg-[#edf1f4]"><div className="h-2 rounded-full bg-[#e48b49]" style={{width: `${Math.max(4, count / Math.max(manifest?.files_scanned || 1, 1) * 100)}%`}} /></div></div>)}{!manifest && <p className="text-sm text-muted-foreground">No manifest yet.</p>}</div></Panel><Panel className="p-6"><h2 className="font-semibold text-[#244661]">Needs OCR</h2>{manifest?.needs_ocr.length ? <div className="mt-4 space-y-3">{manifest.needs_ocr.map((file) => <div key={file.path} className="rounded-xl bg-amber-50 p-3"><p className="text-sm font-medium text-amber-900">{file.file}</p><p className="mt-1 text-xs text-amber-700">{file.reason}</p></div>)}</div> : <div className="mt-4 flex items-center gap-3 rounded-xl bg-emerald-50 p-4 text-sm text-emerald-800"><CheckCircle2Icon className="size-5" />No files are waiting for OCR.</div>}</Panel></div>
          </>}

          {view === "changes" && <>
            <SectionHead eyebrow="Version awareness" title="Changes" detail="Potential document drift and source revision pairs detected by the backend." />
            <div className="grid gap-6 lg:grid-cols-2"><Panel className="p-6"><div className="flex items-center justify-between"><h2 className="font-semibold text-[#244661]">Drift findings</h2><span className="rounded-full bg-[#eef4f7] px-3 py-1 text-xs text-[#557b92]">{drift?.count ?? 0} found</span></div>{drift?.findings.length ? <div className="mt-5 space-y-3">{drift.findings.slice(0, 12).map((item, index) => <div key={index} className="rounded-xl border border-[#edf0f3] p-4"><p className="text-xs font-semibold tracking-wide text-[#b76b38] uppercase">{item.kind}</p><p className="mt-1 text-sm text-[#3e5b70]">{item.description || item.detail || item.file || "Review the underlying documents."}</p></div>)}</div> : <Empty title="No drift detected" detail="The current corpus has no flagged reference or version drift." />}</Panel><Panel className="p-6"><div className="flex items-center justify-between"><h2 className="font-semibold text-[#244661]">Revision pairs</h2><GitCompareArrowsIcon className="size-5 text-[#87a1b4]" /></div>{pairs.length ? <div className="mt-5 space-y-3">{pairs.map((pair, index) => <div key={index} className="rounded-xl border border-[#edf0f3] p-4"><p className="break-all text-xs text-[#667f91]">{pair.a}</p><ArrowRightIcon className="my-2 size-4 text-[#e48b49]" /><p className="break-all text-xs font-medium text-[#315875]">{pair.b}</p></div>)}</div> : <Empty title="No revision pairs" detail="Add multiple versions of a source document to compare them." />}</Panel></div>
          </>}
          <footer className="mt-12 flex flex-wrap items-center justify-between gap-2 border-t border-[#e2e8ed] pt-5 text-xs text-[#9aaab5]"><span>Gluth local workspace · Source-backed engineering review</span><a className="flex items-center gap-1 hover:text-[#426b89]" href="https://github.com/sploit101/Gluth" target="_blank" rel="noreferrer">Repository <ArrowUpRightIcon className="size-3" /></a></footer>
        </main>
      </SidebarInset>
    </SidebarProvider>
  )
}
