import {
  Bell,
  CalendarClock,
  Check,
  Circle,
  ListTodo,
  Pin,
  Plus,
  RefreshCw,
  Save,
  Search,
  Star,
  StickyNote,
  Trash2,
  X,
} from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import {
  deleteTask,
  getCachedMeta,
  getCachedNotes,
  getCachedTasks,
  getMeta,
  getNotes,
  getStorageBackendLabel,
  getTasks,
  migrateLegacyStorage,
  saveNotes,
  saveTask,
  setMeta,
} from "./storage";
import type { AppNote, AppNoteEntry, ChecklistItem, Priority, ProjectId, Requirements, TaskStatus, WorkTask } from "./types";

type ViewFilter = "active" | "today" | "upcoming" | "closed" | "all";
type ActiveView = "tasks" | "notes";
type NoteFilter = "active" | "completed" | "all";

const taskStorageKey = "ai-workbench-work-tasks";
const outputStorageKey = "ai-workbench-saved-outputs";
const reminderStorageKey = "ai-workbench-triggered-reminders";
const projectSettingsMetaKey = "projectSettings";
const taskTemplatesMetaKey = "taskTemplates";

type ProjectSettings = Record<ProjectId, { name: string; context: string }>;
type TaskTemplate = {
  id: string;
  label: string;
  description: string;
  category: string;
  requirements: Requirements;
};

const defaultRequirements: Requirements = {
  outputType: "Task note",
  format: "Markdown",
  tone: "Clear",
  audience: "Personal workspace",
  length: "Short",
  sections: "Details, checklist, next action",
  constraints: "Use this only as local task metadata.",
  imageRequirements: "",
};

const defaultProjects: ProjectSettings = {
  avbob: { name: "AVBOB", context: "Client communication, documents, reports, and operations." },
  naha: { name: "Naha Banking", context: "Banking work, client summaries, process notes, and product copy." },
  personal: { name: "Personal", context: "Personal planning, admin, reminders, and everyday lists." },
  supplysync360: { name: "SupplySync360", context: "Supply chain, supplier coordination, and product operations." },
  bma: { name: "BMA Customer Success", context: "Customer success updates, follow-ups, and issue summaries." },
  thenga: { name: "Thenga", context: "General project work, communication, and documentation." },
};

const defaultTaskTemplates: TaskTemplate[] = [
  { id: "general-task", label: "General task", description: "A flexible task with details and checklist items.", category: "Task", requirements: defaultRequirements },
  { id: "follow-up", label: "Follow-up", description: "A reminder-friendly follow-up item.", category: "Follow-up", requirements: defaultRequirements },
  { id: "checklist", label: "Checklist", description: "A simple tickable list.", category: "Checklist", requirements: defaultRequirements },
  { id: "decision", label: "Decision", description: "A decision or note that needs an outcome.", category: "Decision", requirements: defaultRequirements },
];

const priorities: Priority[] = ["Low", "Normal", "High", "Urgent"];
const statuses: TaskStatus[] = ["Open", "In Progress", "Blocked", "To Do Later", "Closed"];

const emptyTaskDraft = {
  title: "",
  details: "",
  checklistText: "",
  dueDate: "",
  reminderAt: "",
  priority: "Normal" as Priority,
};

const emptyNoteDraft = {
  title: "",
  content: "",
};

function App() {
  const [projects, setProjects] = useState<ProjectSettings>(defaultProjects);
  const [taskTemplates, setTaskTemplates] = useState<TaskTemplate[]>(defaultTaskTemplates);
  const [tasks, setTasks] = useState<WorkTask[]>([]);
  const [notes, setNotes] = useState<AppNote[]>([]);
  const [activeTaskId, setActiveTaskId] = useState("");
  const [activeNoteId, setActiveNoteId] = useState("");
  const [projectId, setProjectId] = useState<ProjectId>("personal");
  const [activeView, setActiveView] = useState<ActiveView>("tasks");
  const [filter, setFilter] = useState<ViewFilter>("active");
  const [noteFilter, setNoteFilter] = useState<NoteFilter>("active");
  const [query, setQuery] = useState("");
  const [taskDraft, setTaskDraft] = useState(emptyTaskDraft);
  const [noteDraft, setNoteDraft] = useState(emptyNoteDraft);
  const [taskModalOpen, setTaskModalOpen] = useState(false);
  const [noteModalOpen, setNoteModalOpen] = useState(false);
  const [showCompletedChecklist, setShowCompletedChecklist] = useState<Record<string, boolean>>({});
  const [message, setMessage] = useState("Loading local workspace...");
  const [loading, setLoading] = useState(true);

  const projectTasks = useMemo(() => tasks.filter((task) => task.projectId === projectId), [projectId, tasks]);
  const projectNotes = useMemo(() => notes.filter((note) => note.projectId === projectId), [notes, projectId]);
  const metrics = useMemo(() => buildMetrics(projectTasks), [projectTasks]);
  const planner = useMemo(() => buildPlanner(projectTasks), [projectTasks]);
  const projectStats = useMemo(() => buildProjectStats(tasks, projects), [tasks, projects]);
  const activeTask = projectTasks.find((task) => task.id === activeTaskId) ?? null;
  const visibleTasks = useMemo(() => {
    return sortedTasks(projectTasks).filter((task) => {
      const haystack = `${task.title} ${task.details} ${task.category} ${projects[task.projectId]?.name ?? ""}`.toLowerCase();
      const matchesQuery = !query.trim() || haystack.includes(query.trim().toLowerCase());
      const matchesFilter =
        filter === "all" ||
        (filter === "active" && task.status !== "Closed") ||
        (filter === "closed" && task.status === "Closed") ||
        (filter === "today" && isTodayOrOverdue(task)) ||
        (filter === "upcoming" && isUpcoming(task));
      return matchesQuery && matchesFilter;
    });
  }, [filter, projectTasks, projects, query]);

  const visibleNotes = useMemo(() => {
    return sortedNotes(projectNotes).filter((note) => {
      const content = note.entries.map((entry) => entry.content).join(" ");
      const haystack = `${note.title} ${content} ${projects[note.projectId]?.name ?? ""}`.toLowerCase();
      const matchesQuery = !query.trim() || haystack.includes(query.trim().toLowerCase());
      const status = note.status ?? "Active";
      const matchesStatus =
        noteFilter === "all" ||
        (noteFilter === "active" && status !== "Closed") ||
        (noteFilter === "completed" && status === "Closed");
      return matchesQuery && matchesStatus;
    });
  }, [noteFilter, projectNotes, projects, query]);
  const activeNote = visibleNotes.find((note) => note.id === activeNoteId) ?? null;

  useEffect(() => {
    void loadCachedData();
    void refreshData();
  }, []);

  useEffect(() => {
    if (!activeTask && projectTasks.length) {
      const firstOpen = sortedTasks(projectTasks).find((task) => task.status !== "Closed") ?? sortedTasks(projectTasks)[0];
      setActiveTaskId(firstOpen.id);
    } else if (!projectTasks.length && activeTaskId) {
      setActiveTaskId("");
    }
  }, [activeTask, activeTaskId, projectTasks]);

  useEffect(() => {
    if (!activeNote && visibleNotes.length) {
      setActiveNoteId(visibleNotes[0].id);
    } else if (!visibleNotes.length && activeNoteId) {
      setActiveNoteId("");
    }
  }, [activeNote, activeNoteId, visibleNotes]);

  async function loadCachedData() {
    try {
      const [taskResult, noteResult, projectResult, templateResult] = await Promise.all([
        getCachedTasks(),
        getCachedNotes(),
        getCachedMeta<ProjectSettings>(projectSettingsMetaKey, defaultProjects),
        getCachedMeta<TaskTemplate[]>(taskTemplatesMetaKey, defaultTaskTemplates),
      ]);
      setTasks(taskResult);
      setNotes(normalizeNotes(noteResult));
      setProjects({ ...defaultProjects, ...projectResult });
      setTaskTemplates(normalizeTaskTemplates(templateResult));
      if (taskResult[0]) {
        setActiveTaskId(taskResult[0].id);
        setProjectId(taskResult[0].projectId);
      }
      setMessage(taskResult.length || noteResult.length ? "Loaded saved tasks and notes from this device." : "Ready for tasks and notes.");
    } catch (error) {
      setMessage(error instanceof Error ? `Local load failed: ${error.message}` : "Local load failed.");
    } finally {
      setLoading(false);
    }
  }

  async function refreshData() {
    try {
      await migrateLegacyStorage(taskStorageKey, outputStorageKey, reminderStorageKey);
      const [taskResult, noteResult, projectResult, templateResult] = await Promise.all([
        getTasks(),
        getNotes(),
        getMeta<ProjectSettings>(projectSettingsMetaKey, defaultProjects),
        getMeta<TaskTemplate[]>(taskTemplatesMetaKey, defaultTaskTemplates),
      ]);
      setTasks(taskResult);
      setNotes(normalizeNotes(noteResult));
      setProjects({ ...defaultProjects, ...projectResult });
      setTaskTemplates(normalizeTaskTemplates(templateResult));
      setMessage(`Synced ${taskResult.length} tasks and ${noteResult.length} notes from ${getStorageBackendLabel()}.`);
    } catch (error) {
      setMessage(error instanceof Error ? `Sync failed: ${error.message}` : "Sync failed.");
    } finally {
      setLoading(false);
    }
  }

  function createTask() {
    const title = taskDraft.title.trim();
    if (!title) {
      setMessage("Add a task title first.");
      return;
    }

    const now = new Date().toISOString();
    const checklist = taskDraft.checklistText
      .split("\n")
      .map((item) => item.trim())
      .filter(Boolean)
      .map((text): ChecklistItem => ({ id: createId("check"), text, done: false }));
    const template = taskTemplates[0] ?? defaultTaskTemplates[0];
    const nextTask: WorkTask = {
      id: createId("task"),
      projectId,
      templateId: template.id,
      title,
      details: taskDraft.details.trim(),
      category: template.category,
      priority: taskDraft.priority,
      dueDate: taskDraft.dueDate,
      reminderAt: taskDraft.reminderAt,
      recurrence: "None",
      recurrenceNote: "",
      status: "Open",
      favorite: false,
      statusHistory: [{ status: "Open", changedAt: now }],
      checklist,
      input: "",
      assets: [],
      requirements: template.requirements,
      gptPrompt: "",
      result: "",
      createdAt: now,
      updatedAt: now,
    };

    persistTask(nextTask);
    setTaskDraft(emptyTaskDraft);
    setActiveTaskId(nextTask.id);
    setActiveView("tasks");
    setTaskModalOpen(false);
    setFilter("active");
    setMessage("Task saved.");
  }

  function patchTask(id: string, patch: Partial<WorkTask>) {
    const task = tasks.find((item) => item.id === id);
    if (!task) return;
    const now = new Date().toISOString();
    const statusChanged = patch.status && patch.status !== task.status;
    const updated: WorkTask = {
      ...task,
      ...patch,
      updatedAt: now,
      statusHistory: statusChanged ? [...task.statusHistory, { status: patch.status as TaskStatus, changedAt: now }] : task.statusHistory,
    };
    persistTask(updated);
  }

  function persistTask(task: WorkTask) {
    setTasks((current) => sortedTasks([task, ...current.filter((item) => item.id !== task.id)]));
    void saveTask(task);
  }

  function removeTask(id: string) {
    setTasks((current) => current.filter((task) => task.id !== id));
    if (activeTaskId === id) setActiveTaskId("");
    void deleteTask(id);
    setMessage("Task deleted.");
  }

  function addChecklistItem(task: WorkTask) {
    const text = window.prompt("Checklist item");
    if (!text?.trim()) return;
    patchTask(task.id, {
      checklist: [...task.checklist, { id: createId("check"), text: text.trim(), done: false }],
    });
  }

  function patchChecklistItem(task: WorkTask, itemId: string, patch: Partial<ChecklistItem>) {
    patchTask(task.id, {
      checklist: task.checklist.map((item) => (item.id === itemId ? { ...item, ...patch } : item)),
    });
  }

  function removeChecklistItem(task: WorkTask, itemId: string) {
    patchTask(task.id, { checklist: task.checklist.filter((item) => item.id !== itemId) });
  }

  function createNote() {
    const title = noteDraft.title.trim();
    const content = noteDraft.content.trim();
    if (!title && !content) {
      setMessage("Add a note title or content first.");
      return;
    }
    const now = new Date().toISOString();
    const note: AppNote = {
      id: createId("note"),
      projectId,
      title: title || "Untitled note",
      entries: content ? [{ id: createId("entry"), content, createdAt: now, updatedAt: now }] : [],
      status: "Active",
      pinned: false,
      favorite: false,
      createdAt: now,
      updatedAt: now,
    };
    persistNotes([note, ...notes]);
    setNoteDraft(emptyNoteDraft);
    setActiveNoteId(note.id);
    setActiveView("notes");
    setNoteModalOpen(false);
    setMessage("Note saved.");
  }

  function patchNote(id: string, patch: Partial<AppNote>) {
    persistNotes(notes.map((note) => (note.id === id ? { ...note, ...patch, updatedAt: new Date().toISOString() } : note)));
  }

  function completeNote(note: AppNote) {
    patchNote(note.id, { status: "Closed" });
    if (activeNoteId === note.id) setActiveNoteId("");
    setMessage("Note completed.");
  }

  function addNoteEntry(note: AppNote) {
    const now = new Date().toISOString();
    const entry: AppNoteEntry = { id: createId("entry"), content: "", createdAt: now, updatedAt: now };
    patchNote(note.id, { entries: [...note.entries, entry] });
  }

  function patchNoteEntry(note: AppNote, entryId: string, content: string) {
    patchNote(note.id, {
      entries: note.entries.map((entry) => (entry.id === entryId ? { ...entry, content, updatedAt: new Date().toISOString() } : entry)),
    });
  }

  function removeNoteEntry(note: AppNote, entryId: string) {
    patchNote(note.id, { entries: note.entries.filter((entry) => entry.id !== entryId) });
  }

  function removeNote(id: string) {
    const next = notes.filter((note) => note.id !== id);
    persistNotes(next);
    if (activeNoteId === id) setActiveNoteId("");
    setMessage("Note deleted.");
  }

  function persistNotes(nextNotes: AppNote[]) {
    const normalized = normalizeNotes(nextNotes);
    setNotes(normalized);
    void saveNotes(normalized);
  }

  function updateProjectName(id: ProjectId, name: string) {
    const next = { ...projects, [id]: { ...(projects[id] ?? { context: "" }), name } };
    setProjects(next);
    void setMeta(projectSettingsMetaKey, next);
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="brand-block">
          <span className="brand-icon"><ListTodo size={20} /></span>
          <div>
            <h1>TaskOS</h1>
            <p>Tasks, notes, reminders, and project focus.</p>
          </div>
        </div>
        <div className="view-tabs" aria-label="Workspace view">
          <button className={activeView === "tasks" ? "active" : ""} onClick={() => setActiveView("tasks")} type="button">
            <ListTodo size={16} />
            Tasks
          </button>
          <button className={activeView === "notes" ? "active" : ""} onClick={() => setActiveView("notes")} type="button">
            <StickyNote size={16} />
            Notes
          </button>
        </div>
        <div className="search-shell">
          <Search size={17} />
          <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder={activeView === "tasks" ? "Search tasks" : "Search notes"} />
        </div>
        <button className="ghost-button sync-button" onClick={() => void refreshData()} type="button">
          <RefreshCw size={16} />
          Sync
        </button>
      </header>

      <section className="project-toolbar" aria-label="Project filter">
        <label>
          Project
          <select value={projectId} onChange={(event) => setProjectId(event.target.value)}>
            {projectStats.map((project) => (
              <option key={project.id} value={project.id}>
                {project.name} ({project.open})
              </option>
            ))}
          </select>
        </label>
          <label>
            Project name
            <input
              value={projects[projectId]?.name ?? ""}
              onChange={(event) => updateProjectName(projectId, event.target.value)}
            />
          </label>
      </section>

      <main className="workspace-main">
        {activeView === "tasks" && (
          <section className="status-strip" aria-label="Task summary">
            <button className={filter === "active" ? "metric-card active" : "metric-card"} onClick={() => setFilter("active")} type="button">
              <strong>{metrics.active}</strong>
              <span>Active</span>
            </button>
            <button className={filter === "today" ? "metric-card active urgent" : "metric-card urgent"} onClick={() => setFilter("today")} type="button">
              <strong>{metrics.dueNow}</strong>
              <span>Due now</span>
            </button>
            <button className={filter === "upcoming" ? "metric-card active" : "metric-card"} onClick={() => setFilter("upcoming")} type="button">
              <strong>{metrics.upcoming}</strong>
              <span>Upcoming</span>
            </button>
            <button className={filter === "closed" ? "metric-card active" : "metric-card"} onClick={() => setFilter("closed")} type="button">
              <strong>{metrics.closed}</strong>
              <span>Closed</span>
            </button>
          </section>
        )}

        <p className={loading ? "sync-banner loading" : "sync-banner"}>{message}</p>

        {activeView === "tasks" ? (
          <>
            <section className="workspace-grid task-workspace">
              <section className="task-column">
                <section className="panel list-panel">
                  <div className="section-heading">
                    <h2><ListTodo size={18} /> Tasks</h2>
                    <div className="compact-actions">
                      <button className="primary-button compact-button" onClick={() => setTaskModalOpen(true)} type="button">
                        <Plus size={15} />
                        New task
                      </button>
                      <select value={filter} onChange={(event) => setFilter(event.target.value as ViewFilter)}>
                        <option value="active">Active</option>
                        <option value="today">Due now</option>
                        <option value="upcoming">Upcoming</option>
                        <option value="closed">Closed</option>
                        <option value="all">All</option>
                      </select>
                    </div>
                  </div>
                  <div className="task-list">
                    {visibleTasks.map((task) => (
                      <button className={activeTaskId === task.id ? "task-card active" : "task-card"} key={task.id} onClick={() => { setActiveTaskId(task.id); setProjectId(task.projectId); }} type="button">
                        <span className={`status-dot ${statusSlug(task.status)}`} />
                        <span>
                          <strong>{task.title}</strong>
                          <small>{projects[task.projectId]?.name ?? task.projectId} / {task.category} / {task.priority}</small>
                        </span>
                        {task.favorite && <Star size={15} />}
                      </button>
                    ))}
                    {visibleTasks.length === 0 && <p className="empty">No tasks match this view.</p>}
                  </div>
                </section>
              </section>

              <section className="detail-column">
                {activeTask ? (
                  <TaskDetail
                    addChecklistItem={addChecklistItem}
                    patchChecklistItem={patchChecklistItem}
                    patchTask={patchTask}
                    projects={projects}
                    removeChecklistItem={removeChecklistItem}
                    removeTask={removeTask}
                    task={activeTask}
                    showCompletedChecklist={Boolean(showCompletedChecklist[activeTask.id])}
                    toggleCompletedChecklist={() => setShowCompletedChecklist((current) => ({ ...current, [activeTask.id]: !current[activeTask.id] }))}
                  />
                ) : (
                  <section className="panel empty-detail">
                    <h2>Select or create a task</h2>
                    <p>Task details, checklist, due dates, reminders, and history will appear here.</p>
                  </section>
                )}
              </section>
            </section>

            <section className="planner-grid" aria-label="Reminder planner">
              <PlannerPanel title="Overdue" tasks={planner.overdue} projects={projects} onOpen={(id) => { setActiveTaskId(id); setActiveView("tasks"); }} />
              <PlannerPanel title="Today" tasks={planner.today} projects={projects} onOpen={(id) => { setActiveTaskId(id); setActiveView("tasks"); }} />
              <PlannerPanel title="Tomorrow" tasks={planner.tomorrow} projects={projects} onOpen={(id) => { setActiveTaskId(id); setActiveView("tasks"); }} />
              <PlannerPanel title="Upcoming" tasks={planner.upcoming} projects={projects} onOpen={(id) => { setActiveTaskId(id); setActiveView("tasks"); }} />
            </section>
          </>
        ) : (
          <section className="workspace-grid notes-workspace">
            <section className="notes-list-column">
              <section className="panel notes-panel">
                <div className="section-heading">
                  <h2><StickyNote size={18} /> Notes</h2>
                  <div className="compact-actions">
                    <button className="primary-button compact-button" onClick={() => setNoteModalOpen(true)} type="button">
                      <Plus size={15} />
                      New note
                    </button>
                    <select value={noteFilter} onChange={(event) => setNoteFilter(event.target.value as NoteFilter)}>
                      <option value="active">Active</option>
                      <option value="completed">Completed</option>
                      <option value="all">All</option>
                    </select>
                    <span className="pill">{visibleNotes.length}</span>
                  </div>
                </div>
                <div className="notes-list">
                  {visibleNotes.map((note) => (
                    <button className={activeNoteId === note.id ? "note-row active" : "note-row"} key={note.id} onClick={() => setActiveNoteId(note.id)} type="button">
                      {note.pinned ? <Pin size={15} /> : <StickyNote size={15} />}
                      <span>
                        <strong>{note.title}</strong>
                        <small>{projects[note.projectId]?.name ?? note.projectId} / {note.entries.length} entries</small>
                      </span>
                    </button>
                  ))}
                  {visibleNotes.length === 0 && <p className="empty">No notes match this view.</p>}
                </div>
              </section>
            </section>

            <section className="detail-column">
              {activeNote ? (
                <NoteDetail
                  addNoteEntry={addNoteEntry}
                  note={activeNote}
                  patchNote={patchNote}
                  patchNoteEntry={patchNoteEntry}
                  projects={projects}
                  completeNote={completeNote}
                  removeNote={removeNote}
                  removeNoteEntry={removeNoteEntry}
                />
              ) : (
                <section className="panel empty-detail">
                  <h2>Select or create a note</h2>
                  <p>Note details and entries for this project will appear here.</p>
                </section>
              )}
            </section>
          </section>
        )}
      </main>

      {taskModalOpen && (
        <div className="modal-layer" role="presentation">
          <button className="modal-backdrop" onClick={() => setTaskModalOpen(false)} type="button" aria-label="Close new task modal" />
          <section className="modal-panel" role="dialog" aria-modal="true" aria-labelledby="new-task-title">
            <div className="modal-header">
              <div>
                <h2 id="new-task-title"><Plus size={18} /> New task</h2>
                <p>{projects[projectId]?.name ?? projectId}</p>
              </div>
              <button className="icon-button" onClick={() => setTaskModalOpen(false)} type="button" aria-label="Close">
                <X size={16} />
              </button>
            </div>
            <input value={taskDraft.title} onChange={(event) => setTaskDraft((current) => ({ ...current, title: event.target.value }))} placeholder="Task title" />
            <textarea value={taskDraft.details} onChange={(event) => setTaskDraft((current) => ({ ...current, details: event.target.value }))} placeholder="Details, context, links, names, anything useful" />
            <textarea className="compact-textarea" value={taskDraft.checklistText} onChange={(event) => setTaskDraft((current) => ({ ...current, checklistText: event.target.value }))} placeholder="Checklist items, one per line" />
            <div className="form-grid">
              <label>Priority<select value={taskDraft.priority} onChange={(event) => setTaskDraft((current) => ({ ...current, priority: event.target.value as Priority }))}>{priorities.map((item) => <option key={item}>{item}</option>)}</select></label>
              <label>Due date<input type="date" value={taskDraft.dueDate} onChange={(event) => setTaskDraft((current) => ({ ...current, dueDate: event.target.value }))} /></label>
              <label>Reminder<input type="datetime-local" value={toLocalInput(taskDraft.reminderAt)} onChange={(event) => setTaskDraft((current) => ({ ...current, reminderAt: fromLocalInput(event.target.value) }))} /></label>
            </div>
            <div className="modal-actions">
              <button className="ghost-button" onClick={() => setTaskModalOpen(false)} type="button">Cancel</button>
              <button className="primary-button" onClick={createTask} type="button"><Save size={16} /> Save task</button>
            </div>
          </section>
        </div>
      )}

      {noteModalOpen && (
        <div className="modal-layer" role="presentation">
          <button className="modal-backdrop" onClick={() => setNoteModalOpen(false)} type="button" aria-label="Close new note modal" />
          <section className="modal-panel" role="dialog" aria-modal="true" aria-labelledby="new-note-title">
            <div className="modal-header">
              <div>
                <h2 id="new-note-title"><StickyNote size={18} /> New note</h2>
                <p>{projects[projectId]?.name ?? projectId}</p>
              </div>
              <button className="icon-button" onClick={() => setNoteModalOpen(false)} type="button" aria-label="Close">
                <X size={16} />
              </button>
            </div>
            <input value={noteDraft.title} onChange={(event) => setNoteDraft((current) => ({ ...current, title: event.target.value }))} placeholder="Note title" />
            <textarea value={noteDraft.content} onChange={(event) => setNoteDraft((current) => ({ ...current, content: event.target.value }))} placeholder="Reference, decision, contact detail, or thought" />
            <div className="modal-actions">
              <button className="ghost-button" onClick={() => setNoteModalOpen(false)} type="button">Cancel</button>
              <button className="primary-button" onClick={createNote} type="button"><Save size={16} /> Save note</button>
            </div>
          </section>
        </div>
      )}
    </div>
  );
}

function TaskDetail({
  addChecklistItem,
  patchChecklistItem,
  patchTask,
  projects,
  removeChecklistItem,
  removeTask,
  showCompletedChecklist,
  task,
  toggleCompletedChecklist,
}: {
  addChecklistItem: (task: WorkTask) => void;
  patchChecklistItem: (task: WorkTask, itemId: string, patch: Partial<ChecklistItem>) => void;
  patchTask: (id: string, patch: Partial<WorkTask>) => void;
  projects: ProjectSettings;
  removeChecklistItem: (task: WorkTask, itemId: string) => void;
  removeTask: (id: string) => void;
  showCompletedChecklist: boolean;
  task: WorkTask;
  toggleCompletedChecklist: () => void;
}) {
  const complete = task.checklist.filter((item) => item.done).length;
  const visibleChecklist = showCompletedChecklist ? task.checklist : task.checklist.filter((item) => !item.done);

  return (
    <section className="panel task-detail">
      <div className="detail-toolbar">
        <span className={`pill status-${statusSlug(task.status)}`}>{task.status}</span>
        <button className={task.favorite ? "icon-button active" : "icon-button"} onClick={() => patchTask(task.id, { favorite: !task.favorite })} type="button" aria-label="Toggle favorite">
          <Star size={16} />
        </button>
        <button className="icon-button danger-icon" onClick={() => removeTask(task.id)} type="button" aria-label="Delete task">
          <Trash2 size={16} />
        </button>
      </div>
      <input className="title-input" value={task.title} onChange={(event) => patchTask(task.id, { title: event.target.value })} />
      <div className="form-grid">
        <label>Project<input readOnly value={projects[task.projectId]?.name ?? task.projectId} /></label>
        <label>Status<select value={task.status} onChange={(event) => patchTask(task.id, { status: event.target.value as TaskStatus })}>{statuses.map((status) => <option key={status}>{status}</option>)}</select></label>
        <label>Priority<select value={task.priority} onChange={(event) => patchTask(task.id, { priority: event.target.value as Priority })}>{priorities.map((priority) => <option key={priority}>{priority}</option>)}</select></label>
      </div>
      <label>Details<textarea value={task.details} onChange={(event) => patchTask(task.id, { details: event.target.value })} /></label>
      <div className="form-grid">
        <label>Due date<input type="date" value={task.dueDate} onChange={(event) => patchTask(task.id, { dueDate: event.target.value })} /></label>
        <label>Reminder<input type="datetime-local" value={toLocalInput(task.reminderAt)} onChange={(event) => patchTask(task.id, { reminderAt: fromLocalInput(event.target.value) })} /></label>
        <label>Recurrence<select value={task.recurrence ?? "None"} onChange={(event) => patchTask(task.id, { recurrence: event.target.value as WorkTask["recurrence"] })}>
          {["None", "Daily", "Weekly", "Monthly", "Quarterly", "Yearly"].map((item) => <option key={item}>{item}</option>)}
        </select></label>
      </div>
      <div className="section-heading">
        <h2><Check size={18} /> Checklist</h2>
        <div className="compact-actions">
          {complete > 0 && (
            <button className="ghost-button compact-button" onClick={toggleCompletedChecklist} type="button">
              {showCompletedChecklist ? "Hide completed" : `Show completed (${complete})`}
            </button>
          )}
          <span className="pill">{complete}/{task.checklist.length}</span>
        </div>
      </div>
      <div className="checklist">
        {visibleChecklist.map((item) => (
          <div className="check-item" key={item.id}>
            <button className={item.done ? "icon-button checked" : "icon-button"} onClick={() => patchChecklistItem(task, item.id, { done: !item.done })} type="button" aria-label="Toggle checklist item">
              {item.done ? <Check size={15} /> : <Circle size={15} />}
            </button>
            <input value={item.text} onChange={(event) => patchChecklistItem(task, item.id, { text: event.target.value })} />
            <button className="icon-button" onClick={() => removeChecklistItem(task, item.id)} type="button" aria-label="Delete checklist item"><Trash2 size={15} /></button>
          </div>
        ))}
        {task.checklist.length === 0 && <p className="empty">No checklist items yet.</p>}
        {task.checklist.length > 0 && visibleChecklist.length === 0 && <p className="empty">All checklist items are completed.</p>}
      </div>
      <button className="ghost-button" onClick={() => addChecklistItem(task)} type="button"><Plus size={16} /> Add checklist item</button>
      <p className="history-line">Created {formatDateTime(task.createdAt)} / Updated {formatDateTime(task.updatedAt)}</p>
    </section>
  );
}

function NoteDetail({
  addNoteEntry,
  completeNote,
  note,
  patchNote,
  patchNoteEntry,
  projects,
  removeNote,
  removeNoteEntry,
}: {
  addNoteEntry: (note: AppNote) => void;
  completeNote: (note: AppNote) => void;
  note: AppNote;
  patchNote: (id: string, patch: Partial<AppNote>) => void;
  patchNoteEntry: (note: AppNote, entryId: string, content: string) => void;
  projects: ProjectSettings;
  removeNote: (id: string) => void;
  removeNoteEntry: (note: AppNote, entryId: string) => void;
}) {
  return (
    <section className="panel note-detail">
      <div className="detail-toolbar">
        <span className="pill">{note.status === "Closed" ? "Completed" : projects[note.projectId]?.name ?? note.projectId}</span>
        {note.status === "Closed" ? (
          <button className="ghost-button compact-button" onClick={() => patchNote(note.id, { status: "Active" })} type="button">
            Reopen
          </button>
        ) : (
          <button className="ghost-button compact-button" onClick={() => completeNote(note)} type="button">
            Complete
          </button>
        )}
        <button className={note.favorite ? "icon-button active" : "icon-button"} onClick={() => patchNote(note.id, { favorite: !note.favorite })} type="button" aria-label="Toggle note favorite"><Star size={16} /></button>
        <button className={note.pinned ? "icon-button active" : "icon-button"} onClick={() => patchNote(note.id, { pinned: !note.pinned })} type="button" aria-label="Toggle note pin"><Pin size={16} /></button>
        <button className="icon-button danger-icon" onClick={() => removeNote(note.id)} type="button" aria-label="Delete note"><Trash2 size={16} /></button>
      </div>
      <input className="title-input" value={note.title} onChange={(event) => patchNote(note.id, { title: event.target.value })} />
      <div className="note-entry-stack">
        {note.entries.map((entry) => (
          <article className="note-entry" key={entry.id}>
            <textarea value={entry.content} onChange={(event) => patchNoteEntry(note, entry.id, event.target.value)} placeholder="Add note detail" />
            <div className="entry-footer">
              <small>{formatDateTime(entry.updatedAt)}</small>
              <button className="icon-button" onClick={() => removeNoteEntry(note, entry.id)} type="button" aria-label="Delete note entry"><Trash2 size={15} /></button>
            </div>
          </article>
        ))}
        {note.entries.length === 0 && <p className="empty">No entries under this note yet.</p>}
      </div>
      <button className="ghost-button" onClick={() => addNoteEntry(note)} type="button"><Plus size={16} /> Add entry</button>
    </section>
  );
}

function PlannerPanel({ title, tasks, projects, onOpen }: { title: string; tasks: WorkTask[]; projects: ProjectSettings; onOpen: (id: string) => void }) {
  return (
    <section className="panel planner-panel">
      <div className="section-heading">
        <h2><CalendarClock size={18} /> {title}</h2>
        <span className="pill">{tasks.length}</span>
      </div>
      <div className="planner-list">
        {tasks.slice(0, 6).map((task) => (
          <button key={task.id} className="planner-row" onClick={() => onOpen(task.id)} type="button">
            <Bell size={15} />
            <span>
              <strong>{task.title}</strong>
              <small>{projects[task.projectId]?.name ?? task.projectId} / {task.dueDate || task.reminderAt ? reminderLabel(task) : "unscheduled"}</small>
            </span>
          </button>
        ))}
        {tasks.length === 0 && <p className="empty">Nothing here.</p>}
      </div>
    </section>
  );
}

function normalizeNotes(items: AppNote[]) {
  return items.map((note) => ({
    ...note,
    entries: note.entries?.length
      ? note.entries
      : note.content
        ? [{ id: createId("entry"), content: note.content, createdAt: note.createdAt, updatedAt: note.updatedAt }]
        : [],
    pinned: Boolean(note.pinned),
    favorite: Boolean(note.favorite),
    status: note.status ?? "Active",
  }));
}

function normalizeTaskTemplates(items: TaskTemplate[]) {
  const blockedWords = ["ai", "prompt"];
  const cleaned = items.filter((item) => {
    const label = `${item.id} ${item.label} ${item.category}`.toLowerCase();
    return !blockedWords.some((word) => label.includes(word));
  });
  return cleaned.length ? cleaned : defaultTaskTemplates;
}

function sortedTasks(items: WorkTask[]) {
  return [...items].sort((a, b) => {
    const aDue = dateValue(a.dueDate || a.reminderAt);
    const bDue = dateValue(b.dueDate || b.reminderAt);
    if (a.status === "Closed" && b.status !== "Closed") return 1;
    if (a.status !== "Closed" && b.status === "Closed") return -1;
    if (aDue && bDue && aDue !== bDue) return aDue - bDue;
    if (aDue && !bDue) return -1;
    if (!aDue && bDue) return 1;
    return dateValue(b.updatedAt) - dateValue(a.updatedAt);
  });
}

function sortedNotes(items: AppNote[]) {
  return [...items].sort((a, b) => Number(b.pinned) - Number(a.pinned) || dateValue(b.updatedAt) - dateValue(a.updatedAt));
}

function buildMetrics(tasks: WorkTask[]) {
  return {
    active: tasks.filter((task) => task.status !== "Closed").length,
    closed: tasks.filter((task) => task.status === "Closed").length,
    dueNow: tasks.filter(isTodayOrOverdue).length,
    upcoming: tasks.filter(isUpcoming).length,
  };
}

function buildPlanner(tasks: WorkTask[]) {
  const open = tasks.filter((task) => task.status !== "Closed");
  return {
    overdue: sortedTasks(open.filter((task) => dayDiff(task.dueDate || task.reminderAt) < 0)),
    today: sortedTasks(open.filter((task) => dayDiff(task.dueDate || task.reminderAt) === 0)),
    tomorrow: sortedTasks(open.filter((task) => dayDiff(task.dueDate || task.reminderAt) === 1)),
    upcoming: sortedTasks(open.filter((task) => dayDiff(task.dueDate || task.reminderAt) > 1)),
  };
}

function buildProjectStats(tasks: WorkTask[], projects: ProjectSettings) {
  return (Object.keys(projects) as ProjectId[]).map((id) => ({
    id,
    name: projects[id].name,
    total: tasks.filter((task) => task.projectId === id).length,
    open: tasks.filter((task) => task.projectId === id && task.status !== "Closed").length,
  }));
}

function isTodayOrOverdue(task: WorkTask) {
  const diff = dayDiff(task.dueDate || task.reminderAt);
  return task.status !== "Closed" && Number.isFinite(diff) && diff <= 0;
}

function isUpcoming(task: WorkTask) {
  const diff = dayDiff(task.dueDate || task.reminderAt);
  return task.status !== "Closed" && Number.isFinite(diff) && diff > 0;
}

function dayDiff(value: string) {
  if (!value) return Number.POSITIVE_INFINITY;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return Number.POSITIVE_INFINITY;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(date);
  target.setHours(0, 0, 0, 0);
  return Math.round((target.getTime() - today.getTime()) / 86_400_000);
}

function dateValue(value: string) {
  const date = new Date(value).getTime();
  return Number.isFinite(date) ? date : 0;
}

function statusSlug(status: TaskStatus) {
  return status.toLowerCase().replace(/\s+/g, "-");
}

function createId(prefix: string) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function toLocalInput(value: string) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const offset = date.getTimezoneOffset() * 60_000;
  return new Date(date.getTime() - offset).toISOString().slice(0, 16);
}

function fromLocalInput(value: string) {
  return value ? new Date(value).toISOString() : "";
}

function formatDateTime(value: string) {
  if (!value) return "not set";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "not set" : date.toLocaleString();
}

function reminderLabel(task: WorkTask) {
  const value = task.dueDate || task.reminderAt;
  const diff = dayDiff(value);
  if (diff < 0) return "overdue";
  if (diff === 0) return "today";
  if (diff === 1) return "tomorrow";
  return `in ${diff} days`;
}

export { App };
