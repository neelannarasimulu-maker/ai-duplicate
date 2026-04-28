import {
  ArrowRight,
  BriefcaseBusiness,
  BarChart3,
  CalendarClock,
  Check,
  Clipboard,
  Download,
  FileImage,
  FileText,
  History,
  ListTodo,
  Save,
  Smartphone,
  Sparkles,
  Trash2,
  Upload,
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useState } from "react";
import { Document, Packer, Paragraph, TextRun } from "docx";
import {
  deleteTask,
  getStorageBackendLabel,
  getMeta,
  getOutputs,
  getTasks,
  migrateLegacyStorage,
  saveOutput,
  saveTask,
  setMeta,
} from "./storage";
import type { Format, InputAsset, Priority, ProjectId, Requirements, SavedOutput, TaskStatus, WorkTask } from "./types";

type TaskTemplate = {
  id: string;
  label: string;
  description: string;
  category: string;
};

const commonTasks: TaskTemplate[] = [
  { id: "draft-document", label: "Draft document", description: "Create a polished document from rough input.", category: "Documentation" },
  { id: "summarize", label: "Summarize information", description: "Extract key points, risks, and next steps.", category: "Analysis" },
  { id: "draft-email", label: "Draft email", description: "Prepare a professional email draft.", category: "Communication" },
  { id: "create-report", label: "Create report", description: "Build a structured report with findings.", category: "Reporting" },
  { id: "proposal-copy", label: "Create proposal copy", description: "Turn notes into proposal-ready language.", category: "Proposal" },
  { id: "presentation-text", label: "Presentation text", description: "Create slide-ready wording and speaker notes.", category: "Presentation" },
];

const projects: Record<ProjectId, { name: string; context: string; tasks: TaskTemplate[] }> = {
  avbob: {
    name: "AVBOB",
    context: "Client communication, document preparation, reports, presentations, and polished business content.",
    tasks: [
      { id: "client-communication", label: "Write client communication", description: "Turn notes into clear client-facing messages.", category: "Communication" },
      { id: "summarize-documents", label: "Summarize documents", description: "Extract key points, risks, and action items.", category: "Analysis" },
      ...commonTasks,
    ],
  },
  naha: {
    name: "Naha Banking",
    context: "Banking-related drafts, client summaries, process notes, reports, proposals, and product copy.",
    tasks: [
      { id: "banking-documents", label: "Draft banking documents", description: "Create clear banking-related documents.", category: "Documentation" },
      { id: "process-notes", label: "Create process notes", description: "Turn workflows into simple process documentation.", category: "Operations" },
      ...commonTasks,
    ],
  },
  personal: {
    name: "Personal",
    context: "Personal planning, admin, writing, research, reminders, and everyday task output.",
    tasks: commonTasks,
  },
  supplysync360: {
    name: "SupplySync360",
    context: "Supply chain, supplier coordination, operational notes, client updates, and product or process documents.",
    tasks: commonTasks,
  },
  bma: {
    name: "BMA Customer Success",
    context: "Customer success updates, client follow-ups, reports, onboarding material, and issue summaries.",
    tasks: commonTasks,
  },
  thenga: {
    name: "Thenga",
    context: "General project work, communications, documentation, reports, and structured output creation.",
    tasks: commonTasks,
  },
};

const defaultRequirements: Requirements = {
  outputType: "Business document",
  format: "Markdown",
  tone: "Professional and clear",
  audience: "Internal team",
  length: "Medium",
  sections: "Purpose, Key points, Final output, Next steps",
  constraints: "",
  imageRequirements: "",
};

const outputStorageKey = "ai-workbench-saved-outputs";
const taskStorageKey = "ai-workbench-work-tasks";
const reminderStorageKey = "ai-workbench-triggered-reminders";
const taskStatuses: TaskStatus[] = ["Open", "In Progress", "Blocked", "Closed"];

function createId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function App() {
  const [projectId, setProjectId] = useState<ProjectId>("avbob");
  const [taskId, setTaskId] = useState(projects.avbob.tasks[0].id);
  const [input, setInput] = useState("");
  const [assets, setAssets] = useState<InputAsset[]>([]);
  const [requirements, setRequirements] = useState<Requirements>(defaultRequirements);
  const [gptPrompt, setGptPrompt] = useState("");
  const [result, setResult] = useState("");
  const [savedOutputs, setSavedOutputs] = useState<SavedOutput[]>([]);
  const [workTasks, setWorkTasks] = useState<WorkTask[]>([]);
  const [activeWorkTaskId, setActiveWorkTaskId] = useState("");
  const [taskProjectFilter, setTaskProjectFilter] = useState<ProjectId | "all">("all");
  const [viewMode, setViewMode] = useState<"dashboard" | "mobile">("dashboard");
  const [quickTaskTitle, setQuickTaskTitle] = useState("");
  const [quickTaskDetails, setQuickTaskDetails] = useState("");
  const [reminderIds, setReminderIds] = useState<string[]>([]);
  const [message, setMessage] = useState("");

  const project = projects[projectId];
  const task = project.tasks.find((item) => item.id === taskId) ?? project.tasks[0];
  const projectHistory = savedOutputs.filter((item) => item.projectId === projectId);
  const activeWorkTask = workTasks.find((item) => item.id === activeWorkTaskId);
  const projectDashboard = buildProjectDashboard(workTasks);
  const selectedProjectStats = projectDashboard.find((item) => item.projectId === projectId);
  const visibleWorkTasks = taskProjectFilter === "all" ? workTasks : workTasks.filter((item) => item.projectId === taskProjectFilter);
  const sortedWorkTasks = [...visibleWorkTasks].sort((a, b) => {
    if (a.status !== b.status) return a.status === "Open" ? -1 : 1;
    return priorityRank(b.priority) - priorityRank(a.priority) || dateValue(a.dueDate) - dateValue(b.dueDate);
  });
  const summary = buildTaskSummary(workTasks);

  const missingDetails = useMemo(() => {
    const missing: string[] = [];
    if (input.trim().length < 20 && assets.length === 0) missing.push("Add source input, notes, or a readable file.");
    if (!requirements.outputType.trim()) missing.push("Choose the output type you need.");
    if (!requirements.audience.trim()) missing.push("Describe who the output is for.");
    if (!requirements.sections.trim()) missing.push("List the required sections or structure.");
    return missing;
  }, [assets.length, input, requirements.audience, requirements.outputType, requirements.sections]);

  useEffect(() => {
    let isMounted = true;

    async function loadDatabase() {
      try {
        await migrateLegacyStorage(taskStorageKey, outputStorageKey, reminderStorageKey);
        const [dbTasks, dbOutputs, dbReminderIds] = await Promise.all([
          getTasks(),
          getOutputs(),
          getMeta<string[]>("triggeredReminderIds", []),
        ]);

        if (!isMounted) return;
        setWorkTasks(dbTasks.map(normalizeWorkTask));
        setSavedOutputs(dbOutputs);
        setReminderIds(dbReminderIds);
        setMessage(`Database loaded. Tasks and history are saved in ${getStorageBackendLabel()}.`);
      } catch (error) {
        setMessage(error instanceof Error ? `Database load failed: ${error.message}` : "Database load failed.");
      }
    }

    void loadDatabase();

    return () => {
      isMounted = false;
    };
  }, []);

  useEffect(() => {
    void setMeta("triggeredReminderIds", reminderIds);
  }, [reminderIds]);

  useEffect(() => {
    const timer = window.setInterval(() => {
      const now = Date.now();
      const dueReminders = workTasks.filter((item) => {
        if (!item.reminderAt || item.status === "Closed" || reminderIds.includes(item.id)) return false;
        return new Date(item.reminderAt).getTime() <= now;
      });

      if (dueReminders.length > 0) {
        const reminderText = dueReminders.map((item) => `${projects[item.projectId].name}: ${item.title}`).join("\n");
        setMessage(`Reminder reached: ${dueReminders.map((item) => item.title).join(", ")}`);
        window.alert(`Task reminder\n\n${reminderText}`);
        setReminderIds((current) => [...current, ...dueReminders.map((item) => item.id)]);
      }
    }, 30000);

    return () => window.clearInterval(timer);
  }, [reminderIds, workTasks]);

  function updateProject(nextProjectId: ProjectId) {
    setProjectId(nextProjectId);
    setTaskId(projects[nextProjectId].tasks[0].id);
    setGptPrompt("");
    setResult("");
    setMessage("");
  }

  function updateRequirement<K extends keyof Requirements>(key: K, value: Requirements[K]) {
    const next = { ...requirements, [key]: value };
    setRequirements(next);
    updateActiveWorkTask("requirements", next);
  }

  function createWorkTask() {
    createTask({
      title: task.label,
      details: "",
      taskProjectId: projectId,
      templateId: task.id,
      category: task.category,
    });
  }

  function createQuickTask() {
    if (!quickTaskTitle.trim() && !quickTaskDetails.trim()) {
      setMessage("Add a quick title or note first.");
      return;
    }

    createTask({
      title: quickTaskTitle.trim() || "Mobile note",
      details: quickTaskDetails.trim(),
      taskProjectId: projectId,
      templateId: projects[projectId].tasks[0].id,
      category: "Mobile",
    });
    setQuickTaskTitle("");
    setQuickTaskDetails("");
    setMessage("Mobile task captured into the shared dashboard.");
  }

  function createTask({
    title,
    details,
    taskProjectId,
    templateId,
    category,
  }: {
    title: string;
    details: string;
    taskProjectId: ProjectId;
    templateId: string;
    category: string;
  }) {
    const newTask: WorkTask = {
      id: createId(),
      projectId: taskProjectId,
      templateId,
      title,
      details,
      category,
      priority: "Normal",
      dueDate: "",
      reminderAt: "",
      status: "Open",
      statusHistory: [{ status: "Open", changedAt: new Date().toISOString() }],
      input: "",
      assets: [],
      requirements: defaultRequirements,
      gptPrompt: "",
      result: "",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    setWorkTasks((current) => [newTask, ...current]);
    void saveTask(newTask);
    openWorkTask(newTask);
    setMessage("Work task created. Add dates or reminders if needed.");
  }

  function updateWorkTask<K extends keyof WorkTask>(id: string, key: K, value: WorkTask[K]) {
    setWorkTasks((current) =>
      current.map((item) => {
        if (item.id !== id) return item;
        const nextStatusHistory =
          key === "status" && item.status !== value
            ? [...item.statusHistory, { status: value as TaskStatus, changedAt: new Date().toISOString() }]
            : item.statusHistory;
        const updated = { ...item, [key]: value, statusHistory: nextStatusHistory, updatedAt: new Date().toISOString() };
        void saveTask(updated);
        return updated;
      }),
    );
    if (key === "projectId") setProjectId(value as ProjectId);
    if (key === "templateId") setTaskId(value as string);
    if (key === "reminderAt") {
      setReminderIds((current) => current.filter((itemId) => itemId !== id));
    }
  }

  function updateActiveWorkTask<K extends keyof WorkTask>(key: K, value: WorkTask[K]) {
    if (!activeWorkTaskId) return;
    updateWorkTask(activeWorkTaskId, key, value);
  }

  function openWorkTask(item: WorkTask) {
    const normalized = normalizeWorkTask(item);
    setActiveWorkTaskId(normalized.id);
    setProjectId(normalized.projectId);
    setTaskId(normalized.templateId);
    setInput(normalized.input);
    setAssets(normalized.assets);
    setRequirements(normalized.requirements);
    setGptPrompt(normalized.gptPrompt);
    setResult(normalized.result);
    setMessage("Task opened in AI mode.");
  }

  function removeWorkTask(id: string) {
    setWorkTasks((current) => current.filter((item) => item.id !== id));
    void deleteTask(id);
    setReminderIds((current) => current.filter((itemId) => itemId !== id));
    if (activeWorkTaskId === id) setActiveWorkTaskId("");
  }

  function saveTaskOnly() {
    if (!activeWorkTask) return;
    updateWorkTask(activeWorkTask.id, "updatedAt", new Date().toISOString());
    setMessage("Task saved. AI prompt/output is optional for this task.");
  }

  async function handleFiles(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    const nextAssets = await Promise.all(
      files.map(async (file) => {
        if (file.type.startsWith("image/")) {
          return {
            id: createId(),
            name: file.name,
            type: "image" as const,
            content: await readAsDataUrl(file),
          };
        }

        const isDocx = /\.docx$/i.test(file.name);
        const canReadText = file.type.startsWith("text/") || /\.(md|txt|csv|json|html)$/i.test(file.name);
        const content = isDocx
          ? await readDocxText(file)
          : canReadText
            ? await file.text()
            : "This file is attached but cannot be read in the browser. For PDF or scanned files, paste the important text into the input box.";

        return {
          id: createId(),
          name: file.name,
          type: canReadText || isDocx ? ("text" as const) : ("file" as const),
          content,
        };
      }),
    );
    setAssets((current) => {
      const next = [...current, ...nextAssets];
      updateActiveWorkTask("assets", next);
      return next;
    });
    event.target.value = "";
  }

  function generatePrompt() {
    if (missingDetails.length > 0) {
      setGptPrompt(buildClarifyingQuestions(project.name, task.label, missingDetails));
      setResult("");
      setMessage("A few details are missing. Answer these first for a stronger result.");
      return;
    }

    const prompt = buildFullLlmPrompt(project.name, project.context, input, assets, requirements);
    setGptPrompt(prompt);
    updateActiveWorkTask("gptPrompt", prompt);
    updateActiveWorkTask("input", input);
    updateActiveWorkTask("assets", assets);
    updateActiveWorkTask("requirements", requirements);
    setResult("");
    updateActiveWorkTask("result", "");
    setMessage("Prompt prepared. Copy it into ChatGPT Plus, then paste the answer below.");
  }

  async function copyPrompt() {
    await navigator.clipboard.writeText(gptPrompt);
    setMessage("Copied the full ChatGPT prompt.");
  }

  async function copyResult() {
    await navigator.clipboard.writeText(result);
    setMessage("Copied result to clipboard.");
  }

  async function downloadResult(format: Format) {
    const extension = format === "Markdown" ? "md" : format === "TXT" ? "txt" : "docx";
    const blob = format === "DOCX" ? await toDocxBlob(result) : new Blob([result], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${project.name}-${task.label}`.replace(/[^a-z0-9]+/gi, "-").toLowerCase() + `.${extension}`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function saveResult() {
    if (!result.trim()) return;
    const saved: SavedOutput = {
      id: createId(),
      projectId,
      workTaskId: activeWorkTaskId,
      taskId: task.id,
      title: `${activeWorkTask?.title ?? task.label} - ${new Date().toLocaleString()}`,
      createdAt: new Date().toISOString(),
      input,
      requirements,
      result,
    };
    setSavedOutputs((current) => [saved, ...current]);
    void saveOutput(saved);
    setMessage("Saved to project history.");
  }

  function loadSaved(saved: SavedOutput) {
    setTaskId(saved.taskId);
    const savedTask = workTasks.find((item) => item.id === saved.workTaskId);
    if (savedTask) openWorkTask(savedTask);
    setInput(saved.input);
    setRequirements(saved.requirements);
    setGptPrompt("");
    setResult(saved.result);
    setMessage("Loaded saved output.");
  }

  function clearWorkspace() {
    setInput("");
    setAssets([]);
    setRequirements(defaultRequirements);
    setGptPrompt("");
    setResult("");
    if (activeWorkTaskId) {
      updateWorkTask(activeWorkTaskId, "input", "");
      updateWorkTask(activeWorkTaskId, "assets", []);
      updateWorkTask(activeWorkTaskId, "requirements", defaultRequirements);
      updateWorkTask(activeWorkTaskId, "gptPrompt", "");
      updateWorkTask(activeWorkTaskId, "result", "");
    }
    setMessage("");
  }

  return (
    <main className="app-shell">
      <section className="topbar">
        <div>
          <p className="eyebrow">Functional AI Workbench</p>
          <h1>Project work assistant</h1>
        </div>
        <button className="ghost-button" onClick={clearWorkspace} type="button">
          <Trash2 size={16} />
          Clear
        </button>
        <button className="ghost-button" onClick={() => setViewMode(viewMode === "dashboard" ? "mobile" : "dashboard")} type="button">
          <Smartphone size={16} />
          {viewMode === "dashboard" ? "Mobile view" : "Dashboard view"}
        </button>
      </section>

      {viewMode === "mobile" && (
        <section className="mobile-shell">
          <section className="panel mobile-capture">
            <h2>
              <Smartphone size={18} />
              Mobile task capture
            </h2>
            <label>
              Project
              <select value={projectId} onChange={(event) => updateProject(event.target.value as ProjectId)}>
                {(Object.keys(projects) as ProjectId[]).map((id) => (
                  <option key={id} value={id}>
                    {projects[id].name}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Task or note title
              <input value={quickTaskTitle} onChange={(event) => setQuickTaskTitle(event.target.value)} placeholder="Call client, capture idea, follow up..." />
            </label>
            <label>
              Notes
              <textarea className="small-textarea" value={quickTaskDetails} onChange={(event) => setQuickTaskDetails(event.target.value)} placeholder="Type or paste notes from your phone." />
            </label>
            <button className="primary-button" onClick={createQuickTask} type="button">
              <Save size={16} />
              Save mobile task
            </button>
          </section>

          <section className="panel">
            <h2>
              <ListTodo size={18} />
              My tasks
            </h2>
            <div className="mobile-task-list">
              {sortedWorkTasks.map((item) => (
                <button key={item.id} className={taskClassName(item, activeWorkTaskId)} onClick={() => openWorkTask(item)} type="button">
                  <span>
                    <strong>{item.title}</strong>
                    <small>{projects[item.projectId].name} - {item.status}</small>
                  </span>
                  <small>{taskDateLabel(item)}</small>
                </button>
              ))}
            </div>
          </section>
        </section>
      )}

      {viewMode === "dashboard" && (
        <>

      <section className="summary-grid">
        <div className="summary-card summary-total">
          <strong>{summary.total}</strong>
          <span>Total tasks</span>
        </div>
        <div className="summary-card summary-open">
          <strong>{summary.open}</strong>
          <span>Open</span>
        </div>
        <div className="summary-card summary-progress">
          <strong>{summary.inProgress}</strong>
          <span>In progress</span>
        </div>
        <div className="summary-card summary-blocked">
          <strong>{summary.blocked}</strong>
          <span>Blocked</span>
        </div>
        <div className="summary-card summary-urgent">
          <strong>{summary.urgent}</strong>
          <span>Urgent</span>
        </div>
        <div className="summary-card summary-reminders">
          <strong>{summary.reminders}</strong>
          <span>Due reminders</span>
        </div>
        <div className="summary-card summary-complete">
          <strong>{summary.closed}</strong>
          <span>Closed</span>
        </div>
      </section>

      <section className="panel dashboard-panel">
        <div className="result-header">
          <h2>
            <BarChart3 size={18} />
            Project dashboard
          </h2>
        </div>
        <div className="project-dashboard-grid">
          {projectDashboard.map((item) => (
            <button key={item.projectId} className={projectId === item.projectId ? "project-dashboard-card active" : "project-dashboard-card"} onClick={() => updateProject(item.projectId)} type="button">
              <strong>{projects[item.projectId].name}</strong>
              <span>{item.total} tasks</span>
              <div className="mini-status-grid">
                <small>Open {item.open}</small>
                <small>Progress {item.inProgress}</small>
                <small>Blocked {item.blocked}</small>
                <small>Closed {item.closed}</small>
              </div>
            </button>
          ))}
        </div>
      </section>

      <section className="workflow">
        <div className="step done">View tasks</div>
        <ArrowRight size={16} />
        <div className="step done">Open task</div>
        <ArrowRight size={16} />
        <div className="step">Save details</div>
        <ArrowRight size={16} />
        <div className="step">Optional AI mode</div>
      </section>

      <div className="layout">
        <aside className="sidebar">
          <section className="panel">
            <h2>
              <BriefcaseBusiness size={18} />
              Project
            </h2>
            <div className="project-grid">
              {(Object.keys(projects) as ProjectId[]).map((id) => (
                <button key={id} className={projectId === id ? "active" : ""} onClick={() => updateProject(id)} type="button">
                  {projects[id].name}
                </button>
              ))}
            </div>
            <p className="context-copy">{project.context}</p>
          </section>

          <section className="panel">
            <h2>
              <BarChart3 size={18} />
              Project status
            </h2>
            {selectedProjectStats && (
              <div className="project-stat-list">
                {taskStatuses.map((status) => (
                  <div className={`project-stat-row status-${statusSlug(status)}`} key={status}>
                    <span>{status}</span>
                    <strong>{selectedProjectStats[statusKey(status)]}</strong>
                  </div>
                ))}
              </div>
            )}
          </section>

          <section className="panel">
            <h2>
              <Sparkles size={18} />
              Task templates
            </h2>
            <div className="task-list">
              {project.tasks.map((item) => (
                <button key={item.id} className={taskId === item.id ? "task-card selected" : "task-card"} onClick={() => setTaskId(item.id)} type="button">
                  <strong>{item.label}</strong>
                  <span>{item.category} - {item.description}</span>
                </button>
              ))}
            </div>
          </section>

          <section className="panel history-panel">
            <h2>
              <History size={18} />
              Project history
            </h2>
            {projectHistory.length === 0 ? (
              <p className="empty">Saved outputs for {project.name} will appear here.</p>
            ) : (
              <div className="history-list">
                {projectHistory.map((saved) => (
                  <button key={saved.id} className="history-item" onClick={() => loadSaved(saved)} type="button">
                    <strong>{saved.title}</strong>
                    <span>{new Date(saved.createdAt).toLocaleDateString()}</span>
                  </button>
                ))}
              </div>
            )}
          </section>
        </aside>

        <section className="work-area">
          <section className="panel">
            <div className="result-header">
              <h2>
                <ListTodo size={18} />
                All tasks
              </h2>
              <button className="primary-button" onClick={createWorkTask} type="button">
                <ListTodo size={16} />
                New task
              </button>
            </div>
            <div className="filter-row">
              <label>
                View by project
                <select value={taskProjectFilter} onChange={(event) => setTaskProjectFilter(event.target.value as ProjectId | "all")}>
                  <option value="all">All projects</option>
                  {(Object.keys(projects) as ProjectId[]).map((id) => (
                    <option key={id} value={id}>
                      {projects[id].name}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            {sortedWorkTasks.length === 0 ? (
              <p className="empty">Create your first task. Tasks appear here across all projects, then expand into AI mode when selected.</p>
            ) : (
              <div className="all-task-list">
                {sortedWorkTasks.map((item) => (
                  <button key={item.id} className={taskClassName(item, activeWorkTaskId)} onClick={() => openWorkTask(item)} type="button">
                    <span>
                      <strong>{item.title}</strong>
                      <small>{projects[item.projectId].name} - {item.category} - {item.priority} - {item.status}</small>
                    </span>
                    <small>{taskDateLabel(item)}</small>
                  </button>
                ))}
              </div>
            )}
          </section>

          {activeWorkTask ? (
            <section className="panel">
              <div className="result-header">
                <h2>
                  <ListTodo size={18} />
                  Task details
                </h2>
                <button className="ghost-button" onClick={() => updateWorkTask(activeWorkTask.id, "status", activeWorkTask.status === "Closed" ? "Open" : "Closed")} type="button">
                  <Check size={16} />
                  Mark {activeWorkTask.status === "Closed" ? "open" : "closed"}
                </button>
              </div>
              <div className="task-status-strip">
                <span className={`status-pill status-${activeWorkTask.status.toLowerCase()}`}>{activeWorkTask.status}</span>
                <span className={`priority-pill priority-${activeWorkTask.priority.toLowerCase()}`}>{activeWorkTask.priority}</span>
                {activeWorkTask.reminderAt && (
                  <span className="reminder-pill">
                    <CalendarClock size={14} />
                    {new Date(activeWorkTask.reminderAt).toLocaleString()}
                  </span>
                )}
              </div>
              <TaskEditor task={activeWorkTask} onChange={updateWorkTask} />
              <label>
                Task notes and details
                <textarea
                  className="small-textarea"
                  value={activeWorkTask.details}
                  onChange={(event) => updateWorkTask(activeWorkTask.id, "details", event.target.value)}
                  placeholder="Record task background, decisions, follow-up notes, or status details here."
                />
              </label>
              <p className="context-copy">
                Created {new Date(activeWorkTask.createdAt).toLocaleString()} - Updated {new Date(activeWorkTask.updatedAt).toLocaleString()}
              </p>
              <div className="status-history">
                <strong>Status updates</strong>
                {activeWorkTask.statusHistory.map((entry, index) => (
                  <span key={`${entry.status}-${entry.changedAt}-${index}`}>
                    {entry.status} - {new Date(entry.changedAt).toLocaleString()}
                  </span>
                ))}
              </div>
              <div className="export-bar">
                <button className="primary-button" onClick={saveTaskOnly} type="button">
                  <Save size={16} />
                  Save task
                </button>
              </div>
              <button className="danger-button" onClick={() => removeWorkTask(activeWorkTask.id)} type="button">
                <Trash2 size={16} />
                Delete selected task
              </button>
            </section>
          ) : (
            <section className="panel">
              <p className="empty">Select or create a task to expand it into AI mode.</p>
            </section>
          )}
          <section className="panel">
            <h2>
              <Smartphone size={18} />
              Mobile capture
            </h2>
            <p className="context-copy">
              Phone notes sync through the configured database when Supabase credentials and tables are ready.
            </p>
          </section>

          {activeWorkTask && (
            <>
              <section className="panel">
                <h2>
                  <FileText size={18} />
                  Input
                </h2>
                <textarea
                  aria-label="Source input"
                  value={input}
                  onChange={(event) => {
                    setInput(event.target.value);
                    updateActiveWorkTask("input", event.target.value);
                  }}
                  placeholder="Paste all notes, instructions, document text, client details, or rough content here. This full input is sent into the generation prompt."
                />
                <label className="upload-control">
                  <Upload size={17} />
                  Upload text files or images
                  <input multiple onChange={handleFiles} type="file" />
                </label>
                {assets.length > 0 && (
                  <div className="asset-list">
                    {assets.map((asset) => (
                      <div className="asset" key={asset.id}>
                        {asset.type === "image" ? <FileImage size={16} /> : <FileText size={16} />}
                        <span>{asset.name}</span>
                        <button
                          type="button"
                          onClick={() =>
                            setAssets((current) => {
                              const next = current.filter((item) => item.id !== asset.id);
                              updateActiveWorkTask("assets", next);
                              return next;
                            })
                          }
                        >
                          Remove
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </section>

              <section className="panel">
                <h2>
                  <Check size={18} />
                  Output requirements
                </h2>
                <div className="form-grid">
                  <label>
                    Output type
                    <input value={requirements.outputType} onChange={(event) => updateRequirement("outputType", event.target.value)} />
                  </label>
                  <label>
                    File format
                    <select value={requirements.format} onChange={(event) => updateRequirement("format", event.target.value as Format)}>
                      <option>Markdown</option>
                      <option>TXT</option>
                      <option>DOCX</option>
                    </select>
                  </label>
                  <label>
                    Tone
                    <input value={requirements.tone} onChange={(event) => updateRequirement("tone", event.target.value)} />
                  </label>
                  <label>
                    Audience
                    <input value={requirements.audience} onChange={(event) => updateRequirement("audience", event.target.value)} />
                  </label>
                  <label>
                    Length
                    <select value={requirements.length} onChange={(event) => updateRequirement("length", event.target.value)}>
                      <option>Short</option>
                      <option>Medium</option>
                      <option>Detailed</option>
                    </select>
                  </label>
                  <label>
                    Required sections
                    <input value={requirements.sections} onChange={(event) => updateRequirement("sections", event.target.value)} />
                  </label>
                </div>
                <label>
                  Constraints
                  <textarea className="small-textarea" value={requirements.constraints} onChange={(event) => updateRequirement("constraints", event.target.value)} />
                </label>
                <label>
                  Image requirements
                  <textarea className="small-textarea" value={requirements.imageRequirements} onChange={(event) => updateRequirement("imageRequirements", event.target.value)} />
                </label>
              </section>

              <section className="panel result-panel">
                <div className="result-header">
                  <h2>
                    <Sparkles size={18} />
                    ChatGPT Plus prompt
                  </h2>
                  <button className="primary-button" onClick={generatePrompt} type="button">
                    <Sparkles size={17} />
                    Prepare prompt
                  </button>
                </div>
                {message && <p className="status-message">{message}</p>}
                <pre className="result-preview">{gptPrompt || "Your full ChatGPT prompt will appear here. Copy it into your ChatGPT Plus session."}</pre>
                <div className="export-bar">
                  <button disabled={!gptPrompt} onClick={copyPrompt} type="button">
                    <Clipboard size={16} />
                    Copy prompt for ChatGPT
                  </button>
                </div>
              </section>

              <section className="panel result-panel">
                <div className="result-header">
                  <h2>
                    <FileText size={18} />
                    Paste ChatGPT output
                  </h2>
                </div>
                <textarea
                  className="output-textarea"
                  value={result}
                  onChange={(event) => {
                    setResult(event.target.value);
                    updateActiveWorkTask("result", event.target.value);
                  }}
                  placeholder="Paste the answer from ChatGPT Plus here. This is the final output you can save or export."
                />
                <div className="export-bar">
                  <button disabled={!result} onClick={copyResult} type="button">
                    <Clipboard size={16} />
                    Copy output
                  </button>
                  <button disabled={!result} onClick={() => void downloadResult(requirements.format)} type="button">
                    <Download size={16} />
                    Download {requirements.format}
                  </button>
                  <button disabled={!result} onClick={saveResult} type="button">
                    <Save size={16} />
                    Save
                  </button>
                </div>
              </section>
            </>
          )}
        </section>
      </div>
        </>
      )}
    </main>
  );
}

function TaskEditor({ task, onChange }: { task: WorkTask; onChange: <K extends keyof WorkTask>(id: string, key: K, value: WorkTask[K]) => void }) {
  const availableTemplates = projects[task.projectId].tasks;

  return (
    <div className="task-editor">
      <label>
        Project
        <select
          value={task.projectId}
          onChange={(event) => {
            const nextProjectId = event.target.value as ProjectId;
            onChange(task.id, "projectId", nextProjectId);
            onChange(task.id, "templateId", projects[nextProjectId].tasks[0].id);
          }}
        >
          {(Object.keys(projects) as ProjectId[]).map((id) => (
            <option key={id} value={id}>
              {projects[id].name}
            </option>
          ))}
        </select>
      </label>
      <label>
        AI task type
        <select value={task.templateId} onChange={(event) => onChange(task.id, "templateId", event.target.value)}>
          {availableTemplates.map((template) => (
            <option key={template.id} value={template.id}>
              {template.label}
            </option>
          ))}
        </select>
      </label>
      <label>
        Task title
        <input value={task.title} onChange={(event) => onChange(task.id, "title", event.target.value)} />
      </label>
      <label>
        Category
        <input value={task.category} onChange={(event) => onChange(task.id, "category", event.target.value)} />
      </label>
      <label>
        Priority
        <select value={task.priority} onChange={(event) => onChange(task.id, "priority", event.target.value as Priority)}>
          <option>Low</option>
          <option>Normal</option>
          <option>High</option>
          <option>Urgent</option>
        </select>
      </label>
      <label>
        Due date
        <input value={task.dueDate} onChange={(event) => onChange(task.id, "dueDate", event.target.value)} type="date" />
      </label>
      <label>
        Reminder
        <input value={task.reminderAt} onChange={(event) => onChange(task.id, "reminderAt", event.target.value)} type="datetime-local" />
      </label>
      <label>
        Status
        <select value={task.status} onChange={(event) => onChange(task.id, "status", event.target.value as TaskStatus)}>
          {taskStatuses.map((status) => (
            <option key={status}>{status}</option>
          ))}
        </select>
      </label>
    </div>
  );
}

function normalizeWorkTask(task: WorkTask): WorkTask {
  const fallbackProjectId = task.projectId && projects[task.projectId] ? task.projectId : "avbob";
  const fallbackTemplateId = projects[fallbackProjectId].tasks.some((template) => template.id === task.templateId)
    ? task.templateId
    : projects[fallbackProjectId].tasks[0].id;

  return {
    ...task,
    projectId: fallbackProjectId,
    templateId: fallbackTemplateId,
    title: task.title || "Untitled task",
    details: task.details ?? "",
    category: task.category || "General",
    priority: task.priority || "Normal",
    dueDate: task.dueDate ?? "",
    reminderAt: task.reminderAt ?? "",
    status: normalizeStatus(task.status),
    statusHistory: task.statusHistory?.length
      ? task.statusHistory.map((entry) => ({ ...entry, status: normalizeStatus(entry.status) }))
      : [{ status: normalizeStatus(task.status), changedAt: task.updatedAt || task.createdAt || new Date().toISOString() }],
    input: task.input ?? "",
    assets: task.assets ?? [],
    requirements: task.requirements ?? defaultRequirements,
    gptPrompt: task.gptPrompt ?? "",
    result: task.result ?? "",
    createdAt: task.createdAt || new Date().toISOString(),
    updatedAt: task.updatedAt || task.createdAt || new Date().toISOString(),
  };
}

function priorityRank(priority: Priority) {
  return { Low: 1, Normal: 2, High: 3, Urgent: 4 }[priority] ?? 0;
}

function dateValue(date: string) {
  return date ? new Date(date).getTime() : Number.MAX_SAFE_INTEGER;
}

function buildTaskSummary(tasks: WorkTask[]) {
  const now = Date.now();
  return {
    total: tasks.length,
    open: tasks.filter((item) => item.status === "Open").length,
    inProgress: tasks.filter((item) => item.status === "In Progress").length,
    blocked: tasks.filter((item) => item.status === "Blocked").length,
    closed: tasks.filter((item) => item.status === "Closed").length,
    urgent: tasks.filter((item) => item.status !== "Closed" && item.priority === "Urgent").length,
    reminders: tasks.filter((item) => item.status !== "Closed" && item.reminderAt && new Date(item.reminderAt).getTime() <= now).length,
  };
}

function buildProjectDashboard(tasks: WorkTask[]) {
  return (Object.keys(projects) as ProjectId[]).map((projectId) => {
    const projectTasks = tasks.filter((task) => task.projectId === projectId);
    const summary = buildTaskSummary(projectTasks);
    return { projectId, ...summary };
  });
}

function statusKey(status: TaskStatus): "open" | "inProgress" | "blocked" | "closed" {
  if (status === "In Progress") return "inProgress";
  return status.toLowerCase() as "open" | "blocked" | "closed";
}

function taskClassName(task: WorkTask, activeWorkTaskId: string) {
  const classes = ["work-task", `status-${statusSlug(task.status)}`, `priority-${task.priority.toLowerCase()}`];
  if (activeWorkTaskId === task.id) classes.push("active");
  if (isOverdue(task)) classes.push("overdue");
  if (isReminderDue(task)) classes.push("reminder-due");
  return classes.join(" ");
}

function taskDateLabel(task: WorkTask) {
  if (isOverdue(task)) return `Overdue: ${task.dueDate}`;
  if (task.dueDate) return `Due: ${task.dueDate}`;
  if (isReminderDue(task)) return "Reminder due";
  return "No due date";
}

function isOverdue(task: WorkTask) {
  if (!task.dueDate || task.status === "Closed") return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(task.dueDate).getTime() < today.getTime();
}

function isReminderDue(task: WorkTask) {
  return Boolean(task.reminderAt && task.status !== "Closed" && new Date(task.reminderAt).getTime() <= Date.now());
}

function statusSlug(status: TaskStatus) {
  return status.toLowerCase().replace(/\s+/g, "-");
}

function normalizeStatus(status: TaskStatus | "Complete" | undefined): TaskStatus {
  if (status === "Complete") return "Closed";
  if (status && taskStatuses.includes(status as TaskStatus)) return status as TaskStatus;
  return "Open";
}

function readAsDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function readDocxText(file: File) {
  const mammoth = await import("mammoth");
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value.trim() || "The DOCX file was read, but no text content was found.";
}

function buildClarifyingQuestions(projectName: string, taskLabel: string, missingDetails: string[]) {
  return [
    "# Clarifying details needed",
    "",
    `Project: ${projectName}`,
    `Task: ${taskLabel}`,
    "",
    "Before generating the final output, please provide:",
    ...missingDetails.map((item, index) => `${index + 1}. ${item}`),
    "",
    "Once these details are added, click Generate again.",
  ].join("\n");
}

function buildFullLlmPrompt(
  projectName: string,
  projectContext: string,
  input: string,
  assets: InputAsset[],
  requirements: Requirements,
) {
  const readableFiles = assets.filter((asset) => asset.type === "text");
  const imageFiles = assets.filter((asset) => asset.type === "image");
  const unreadableFiles = assets.filter((asset) => asset.type === "file");

  return [
    "You are a practical AI work assistant. Process the full source material and create a new output document.",
    "I am using ChatGPT Plus manually, so produce the final answer directly in this chat.",
    "",
    "PROJECT",
    `Name: ${projectName}`,
    `Context: ${projectContext}`,
    "",
    "OUTPUT REQUIREMENTS",
    `Output type: ${requirements.outputType}`,
    `Format: ${requirements.format}`,
    `Tone: ${requirements.tone}`,
    `Audience: ${requirements.audience}`,
    `Length: ${requirements.length}`,
    `Required sections: ${requirements.sections}`,
    `Constraints: ${requirements.constraints || "None"}`,
    `Image requirements: ${requirements.imageRequirements || "None"}`,
    "",
    "SOURCE INPUT",
    input.trim() || "No typed input provided.",
    "",
    "READABLE FILE CONTENT",
    readableFiles.length ? readableFiles.map((asset) => `--- ${asset.name} ---\n${asset.content}`).join("\n\n") : "No readable text files attached.",
    "",
    "IMAGE ATTACHMENTS",
    imageFiles.length ? imageFiles.map((asset) => `- ${asset.name}`).join("\n") : "No image attachments.",
    "",
    "OTHER ATTACHMENTS",
    unreadableFiles.length ? unreadableFiles.map((asset) => `- ${asset.name}: ${asset.content}`).join("\n") : "No other attachments.",
    "",
    "INSTRUCTIONS",
    "Create the requested output as a new, business-ready document. Do not summarize the prompt. Use all relevant source material. Ask clarifying questions only if required information is genuinely missing.",
  ].join("\n");
}

async function toDocxBlob(markdown: string) {
  const children = markdown.split("\n").map((line) => {
    const trimmed = line.trim();
    const isHeading = trimmed.startsWith("#");
    const text = trimmed.replace(/^#+\s*/, "").replace(/^\*\*([^*]+):\*\*\s*/, "$1: ");

    return new Paragraph({
      spacing: { after: isHeading ? 180 : 90 },
      children: [
        new TextRun({
          text,
          bold: isHeading || /^\w[\w\s]+:/.test(text),
          size: isHeading ? 30 : 22,
        }),
      ],
    });
  });

  const document = new Document({
    sections: [{ children }],
  });

  return Packer.toBlob(document);
}

export { App };
