import {
  ActivityIcon,
  ArrowUpRightIcon,
  FileSearchIcon,
  FilesIcon,
  FolderKanbanIcon,
  GitCompareArrowsIcon,
  PanelsTopLeftIcon,
} from "lucide-react"
import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarHeader,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarRail,
  useSidebar,
} from "@/components/ui/sidebar"

export type View = "overview" | "search" | "projects" | "station" | "sources" | "changes"

const navigation = [
  { id: "overview", label: "Overview", icon: PanelsTopLeftIcon },
  { id: "search", label: "Search corpus", icon: FileSearchIcon },
  { id: "projects", label: "Projects", icon: FolderKanbanIcon },
  { id: "station", label: "Station file", icon: FilesIcon },
  { id: "sources", label: "Source files", icon: ActivityIcon },
  { id: "changes", label: "Changes", icon: GitCompareArrowsIcon },
] as const

export function AppSidebar({
  active,
  onSelect,
  projectCount,
  status,
}: {
  active: View
  onSelect: (view: View) => void
  projectCount: number
  status: string
}) {
  const { isMobile, setOpenMobile } = useSidebar()
  return (
    <Sidebar collapsible="icon" className="border-r border-slate-200/70">
      <SidebarHeader className="px-3 pt-5 pb-3">
        <div className="flex items-center gap-3 rounded-xl px-2 py-2">
          <div className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-[#e48b49] text-white shadow-sm">
            <span className="text-xl font-bold">G</span>
          </div>
          <div className="grid min-w-0 flex-1 text-left leading-tight group-data-[collapsible=icon]:hidden">
            <span className="truncate text-base font-semibold tracking-tight">Gluth</span>
            <span className="truncate text-xs text-sidebar-foreground/55">Engineering intelligence</span>
          </div>
        </div>
      </SidebarHeader>
      <SidebarContent className="px-2">
        <SidebarGroup>
          <SidebarGroupLabel className="mb-2 px-3 text-[10px] font-semibold tracking-[0.16em] text-sidebar-foreground/45 uppercase">
            Workspace
          </SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu className="gap-1">
              {navigation.map((item) => (
                <SidebarMenuItem key={item.id}>
                  <SidebarMenuButton
                    isActive={active === item.id}
                    tooltip={item.label}
                    onClick={() => {
                      onSelect(item.id)
                      if (isMobile) setOpenMobile(false)
                    }}
                    className="h-10 rounded-lg px-3 text-sm data-[active=true]:bg-[#e9eff6] data-[active=true]:font-semibold data-[active=true]:text-[#1d4165]"
                  >
                    <item.icon className="size-4" />
                    <span>{item.label}</span>
                    {item.id === "projects" && projectCount > 0 && (
                      <span className="ml-auto text-xs text-sidebar-foreground/50">{projectCount}</span>
                    )}
                  </SidebarMenuButton>
                </SidebarMenuItem>
              ))}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>
      <SidebarFooter className="p-4 group-data-[collapsible=icon]:p-2">
        <div className="rounded-xl border border-sidebar-border bg-sidebar-accent p-3 text-sidebar-foreground group-data-[collapsible=icon]:hidden">
          <div className="flex items-center gap-2 text-xs font-medium">
            <span className={`size-2 rounded-full ${status === "ok" ? "bg-emerald-500" : "bg-amber-500"}`} />
            Local API {status === "ok" ? "online" : "needs setup"}
          </div>
          <p className="mt-1 text-[11px] leading-relaxed text-muted-foreground">Documents stay on this machine.</p>
        </div>
        <a
          href="https://github.com/sploit101/Gluth/issues"
          target="_blank"
          rel="noreferrer"
          className="mt-2 flex items-center gap-2 px-2 text-xs text-sidebar-foreground/55 hover:text-sidebar-foreground group-data-[collapsible=icon]:justify-center"
          title="Project issues"
        >
          <ArrowUpRightIcon className="size-3.5" />
          <span className="group-data-[collapsible=icon]:hidden">Project issues</span>
        </a>
      </SidebarFooter>
      <SidebarRail />
    </Sidebar>
  )
}
