import {
  ArrowRight,
  BriefcaseBusiness,
  BarChart3,
  BellRing,
  CalendarClock,
  Check,
  CheckSquare,
  Clipboard,
  Download,
  FileImage,
  FileText,
  History,
  LayoutDashboard,
  ListTodo,
  Pin,
  Plus,
  Save,
  Smartphone,
  Sparkles,
  StickyNote,
  Trash2,
  Upload,
  WandSparkles,
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useState } from "react";
import { Document, Packer, Paragraph, TextRun } from "docx";
import {
  deleteTask,
  getStorageBackendLabel,
  getMeta,
  getNotes,
  getOutputs,
  getTasks,
  migrateLegacyStorage,
  saveOutput,
  saveNotes,
  saveTask,
  setMeta,
} from "./storage";
import type { AppNote, AppNoteEntry, ChecklistItem, Format, InputAsset, Priority, ProjectId, Requirements, SavedOutput, TaskStatus, WorkTask } from "./types";

type TaskTemplate = {
  id: string;
  label: string;
  description: string;
  category: string;
  requirements: Requirements;
};

const requirementPresets = {
  summary: {
    outputType: "Information summary document",
    format: "Markdown",
    tone: "Clear, practical, and neutral",
    audience: "Internal team",
    length: "Medium",
    sections: "Executive summary, Key facts, Decisions or implications, Risks or gaps, Action items",
    constraints: "Separate confirmed facts from assumptions. Keep the wording easy to reuse in a document.",
    imageRequirements: "",
  },
  document: {
    outputType: "New business document",
    format: "Markdown",
    tone: "Professional and clear",
    audience: "Internal team",
    length: "Detailed",
    sections: "Title, Purpose, Background, Main content, Recommendations, Next steps",
    constraints: "Produce a complete document, not notes about the document. Use headings and concise paragraphs.",
    imageRequirements: "",
  },
  email: {
    outputType: "Email draft",
    format: "Markdown",
    tone: "Warm, concise, and professional",
    audience: "Recipient named in the source notes",
    length: "Short",
    sections: "Subject line, Email body, Optional follow-up note",
    constraints: "Write ready-to-send email copy. Include a clear ask or next step where appropriate.",
    imageRequirements: "",
  },
  report: {
    outputType: "Structured report",
    format: "Markdown",
    tone: "Evidence-led and businesslike",
    audience: "Decision makers",
    length: "Detailed",
    sections: "Executive summary, Context, Findings, Evidence, Risks, Recommendations, Next steps",
    constraints: "Make the report useful for decisions. Flag missing evidence clearly.",
    imageRequirements: "",
  },
  meetingBrief: {
    outputType: "Meeting or forum briefing pack",
    format: "Markdown",
    tone: "Concise, structured, and decision-focused",
    audience: "Meeting attendees and decision makers",
    length: "Medium",
    sections: "Purpose, Agenda, Current status, Key discussion points, Risks, Decisions needed, Actions",
    constraints: "Use the source notes to prepare a practical meeting pack. Separate discussion points from decisions needed.",
    imageRequirements: "",
  },
  marketResearch: {
    outputType: "Market research and opportunity note",
    format: "Markdown",
    tone: "Analytical, balanced, and commercially useful",
    audience: "Business development and leadership team",
    length: "Detailed",
    sections: "Executive summary, Company profile, Market context, Opportunity fit, Risks and unknowns, Recommended next steps",
    constraints: "Distinguish source-provided facts from assumptions. Include gaps that need verification.",
    imageRequirements: "",
  },
  proposal: {
    outputType: "Proposal-ready document",
    format: "Markdown",
    tone: "Confident, useful, and client-ready",
    audience: "Client or sponsor",
    length: "Detailed",
    sections: "Overview, Client need, Proposed approach, Deliverables, Timeline, Assumptions, Next steps",
    constraints: "Avoid generic sales language. Convert rough notes into concrete proposed work.",
    imageRequirements: "",
  },
  presentation: {
    outputType: "Presentation content",
    format: "Markdown",
    tone: "Clear, structured, and presentation-ready",
    audience: "Meeting audience",
    length: "Medium",
    sections: "Slide title, Slide bullets, Speaker notes, Closing action",
    constraints: "Keep slide bullets short. Put detail in speaker notes.",
    imageRequirements: "Suggest useful visuals, charts, or screenshots where they would improve the presentation.",
  },
  process: {
    outputType: "Process document",
    format: "Markdown",
    tone: "Plain, precise, and operational",
    audience: "People following the process",
    length: "Detailed",
    sections: "Purpose, Scope, Roles, Steps, Exceptions, Checklist, Owner and review date",
    constraints: "Write actionable steps in order. Include a checklist that can be ticked off.",
    imageRequirements: "",
  },
  clientUpdate: {
    outputType: "Client or stakeholder update",
    format: "Markdown",
    tone: "Clear, calm, and accountable",
    audience: "Client, sponsor, or internal stakeholder",
    length: "Short",
    sections: "Subject line, Status summary, Progress, Risks or blockers, Decisions needed, Next steps",
    constraints: "Write copy that can be sent or pasted into an email/message with minimal editing.",
    imageRequirements: "",
  },
  checklist: {
    outputType: "Checklist or shopping list",
    format: "Markdown",
    tone: "Plain and practical",
    audience: "Personal use or project team",
    length: "Short",
    sections: "Checklist, Notes, Next steps",
    constraints: "Keep items tickable, specific, and easy to complete.",
    imageRequirements: "",
  },
} satisfies Record<string, Requirements>;

const commonTasks: TaskTemplate[] = [
  { id: "summarize", label: "Summarize information", description: "Turn notes, documents, or emails into facts, gaps, risks, and action items.", category: "Analysis", requirements: requirementPresets.summary },
  { id: "draft-document", label: "Draft business document", description: "Produce a complete reusable document with headings, flow, and next steps.", category: "Documentation", requirements: requirementPresets.document },
  { id: "draft-email", label: "Draft email or message", description: "Prepare ready-to-send communication with a clear ask and next step.", category: "Communication", requirements: requirementPresets.email },
  { id: "client-update", label: "Client/stakeholder update", description: "Create a status update covering progress, blockers, decisions, and next steps.", category: "Communication", requirements: requirementPresets.clientUpdate },
  { id: "meeting-brief", label: "Meeting/forum brief", description: "Prepare agenda notes, decision points, risks, and follow-up actions.", category: "Planning", requirements: requirementPresets.meetingBrief },
  { id: "market-research", label: "Market research note", description: "Structure company research, opportunity fit, risks, and recommended next steps.", category: "Research", requirements: requirementPresets.marketResearch },
  { id: "create-report", label: "Create report", description: "Build a structured report with findings, evidence, and recommendations.", category: "Reporting", requirements: requirementPresets.report },
  { id: "proposal-copy", label: "Create proposal copy", description: "Turn notes into proposal-ready language with scope and deliverables.", category: "Proposal", requirements: requirementPresets.proposal },
  { id: "process-document", label: "Process/support document", description: "Document operational steps, roles, exceptions, and a tickable process checklist.", category: "Operations", requirements: requirementPresets.process },
  { id: "presentation-text", label: "Presentation text", description: "Create slide-ready bullets with practical speaker notes.", category: "Presentation", requirements: requirementPresets.presentation },
  { id: "checklist", label: "Checklist or shopping list", description: "Create tickable items for shopping, subtasks, admin, or follow-up work.", category: "Checklist", requirements: requirementPresets.checklist },
];

const projects: Record<ProjectId, { name: string; context: string; tasks: TaskTemplate[] }> = {
  avbob: {
    name: "AVBOB",
    context: "Client communication, document preparation, reports, presentations, and polished business content.",
    tasks: [
      { id: "client-communication", label: "Write client communication", description: "Turn notes into a ready-to-send client email or letter.", category: "Communication", requirements: requirementPresets.email },
      { id: "summarize-documents", label: "Summarize documents", description: "Extract key points, risks, gaps, decisions, and action items.", category: "Analysis", requirements: requirementPresets.summary },
      { id: "advisory-forum", label: "Advisory forum brief", description: "Prepare a concise meeting pack with risks, decisions, and actions.", category: "Planning", requirements: requirementPresets.meetingBrief },
      { id: "support-process", label: "Support process document", description: "Draft customer support, technical support, escalation, and SLA process content.", category: "Operations", requirements: requirementPresets.process },
      ...commonTasks,
    ],
  },
  naha: {
    name: "Naha Banking",
    context: "Banking-related drafts, client summaries, process notes, reports, proposals, and product copy.",
    tasks: [
      { id: "banking-documents", label: "Draft banking documents", description: "Create clear banking documents with assumptions and compliance-sensitive wording flagged.", category: "Documentation", requirements: requirementPresets.document },
      { id: "process-notes", label: "Create process notes", description: "Turn workflows into simple process documentation and tickable steps.", category: "Operations", requirements: requirementPresets.process },
      { id: "banking-client-update", label: "Banking client update", description: "Summarize progress, dependencies, commercial items, and decisions needed.", category: "Communication", requirements: requirementPresets.clientUpdate },
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
    tasks: [
      { id: "bd-opportunity-note", label: "BD opportunity note", description: "Prepare company research, opportunity fit, risks, and next engagement steps.", category: "Research", requirements: requirementPresets.marketResearch },
      { id: "customer-success-update", label: "Customer success update", description: "Draft a practical update for client, internal, or delivery stakeholders.", category: "Communication", requirements: requirementPresets.clientUpdate },
      ...commonTasks,
    ],
  },
  thenga: {
    name: "Thenga",
    context: "General project work, communications, documentation, reports, and structured output creation.",
    tasks: commonTasks,
  },
};

const defaultRequirements: Requirements = {
  outputType: "Information summary or business document",
  format: "Markdown",
  tone: "Professional and clear",
  audience: "Internal team",
  length: "Medium",
  sections: "Executive summary, Key information, Draft output, Action items, Next steps",
  constraints: "Produce a usable final output, not only advice. Make action items clear and practical.",
  imageRequirements: "",
};

const outputStorageKey = "ai-workbench-saved-outputs";
const taskStorageKey = "ai-workbench-work-tasks";
const reminderStorageKey = "ai-workbench-triggered-reminders";
const triggeredReminderMetaKey = "triggeredReminderIds";
const taskStatuses: TaskStatus[] = ["Open", "In Progress", "Blocked", "To Do Later", "Closed"];
const mobileStatusOrder: TaskStatus[] = ["In Progress", "Open", "Blocked", "To Do Later", "Closed"];
const maxUploadSizeBytes = 8 * 1024 * 1024;
const maxTextAssetCharacters = 120_000;

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
  const [notes, setNotes] = useState<AppNote[]>([]);
  const [activeWorkTaskId, setActiveWorkTaskId] = useState("");
  const [viewMode, setViewMode] = useState<"home" | "work" | "projects" | "reminders" | "notes" | "mobile">("home");
  const [quickTaskTitle, setQuickTaskTitle] = useState("");
  const [quickTaskDetails, setQuickTaskDetails] = useState("");
  const [quickChecklistDraft, setQuickChecklistDraft] = useState("");
  const [quickChecklistItems, setQuickChecklistItems] = useState<ChecklistItem[]>([]);
  const [quickChecklistItem, setQuickChecklistItem] = useState("");
  const [noteDraft, setNoteDraft] = useState({ title: "", content: "" });
  const [mobileFullOpen, setMobileFullOpen] = useState(false);
  const [mobileSection, setMobileSection] = useState<"capture" | "tasks" | "reminders">("tasks");
  const [reminderDrafts, setReminderDrafts] = useState<Record<string, string>>({});
  const [triggeredReminderIds, setTriggeredReminderIds] = useState<string[]>([]);
  const [notificationPermission, setNotificationPermission] = useState<NotificationPermission | "unsupported">(() =>
    typeof Notification === "undefined" ? "unsupported" : Notification.permission,
  );
  const [now, setNow] = useState(() => Date.now());
  const [syncing, setSyncing] = useState(false);
  const [message, setMessage] = useState("");

  const project = projects[projectId];
  const task = project.tasks.find((item) => item.id === taskId) ?? project.tasks[0];
  const projectHistory = savedOutputs.filter((item) => item.projectId === projectId);
  const activeWorkTask = workTasks.find((item) => item.id === activeWorkTaskId);
  const projectDashboard = buildProjectDashboard(workTasks);
  const selectedProjectStats = projectDashboard.find((item) => item.projectId === projectId);
  const projectWorkTasks = workTasks.filter((item) => item.projectId === projectId);
  const mobileProjectGroups = buildMobileProjectGroups(workTasks);
  const mobileCurrentProjectGroup = mobileProjectGroups.find((item) => item.projectId === projectId) ?? mobileProjectGroups[0];
  const sortedWorkTasks = [...workTasks].sort((a, b) => {
    if (a.status !== b.status) return statusRank(a.status) - statusRank(b.status);
    return priorityRank(b.priority) - priorityRank(a.priority) || dateValue(a.dueDate) - dateValue(b.dueDate);
  });
  const sortedProjectTasks = [...projectWorkTasks].sort((a, b) => {
    if (a.status !== b.status) return statusRank(a.status) - statusRank(b.status);
    return priorityRank(b.priority) - priorityRank(a.priority) || dateValue(a.dueDate) - dateValue(b.dueDate);
  });
  const summary = buildTaskSummary(workTasks);
  const reminderPlanner = buildReminderPlanner(workTasks, now);
  const noteEntryCount = notes.reduce((total, note) => total + note.entries.length, 0);
  const activeFocusTasks = sortedWorkTasks.filter((item) => item.status !== "Closed").slice(0, 6);
  const recentNotes = [...notes].sort((a, b) => dateValue(b.updatedAt) - dateValue(a.updatedAt)).slice(0, 4);

  const missingDetails = useMemo(() => {
    const missing: string[] = [];
    if (input.trim().length < 20 && assets.length === 0) missing.push("Add source input, notes, or a readable file.");
    if (!requirements.outputType.trim()) missing.push("Choose the output type you need.");
    if (!requirements.audience.trim()) missing.push("Describe who the output is for.");
    if (!requirements.sections.trim()) missing.push("List the required sections or structure.");
    return missing;
  }, [assets.length, input, requirements.audience, requirements.outputType, requirements.sections]);

  async function refreshData() {
    setSyncing(true);
    try {
      try {
        await migrateLegacyStorage(taskStorageKey, outputStorageKey, reminderStorageKey);
      } catch {
        // Legacy localStorage migration should not block a fresh Supabase pull.
      }
      const taskResult = await getTasks();
      let outputResult: SavedOutput[] = [];
      let outputWarning = "";

      try {
        outputResult = await getOutputs();
      } catch (error) {
        outputWarning = error instanceof Error ? ` Saved outputs failed: ${error.message}` : " Saved outputs failed.";
      }
      const noteResult = await getNotes();
      const triggeredResult = await getMeta<string[]>(triggeredReminderMetaKey, []);

      setWorkTasks(taskResult.map(normalizeWorkTask));
      setNotes(noteResult.map(normalizeNote));
      setSavedOutputs(outputResult);
      setTriggeredReminderIds(triggeredResult);
      setNow(Date.now());
      setMessage(`Synced ${taskResult.length} tasks, ${noteResult.length} notes, and ${outputResult.length} saved outputs from ${getStorageBackendLabel()}.${outputWarning}`);
    } catch (error) {
      setMessage(error instanceof Error ? `Database sync failed: ${error.message}` : "Database sync failed.");
    } finally {
      setSyncing(false);
    }
  }

  useEffect(() => {
    void refreshData();
  }, []);

  useEffect(() => {
    function refreshWhenVisible() {
      if (document.visibilityState === "visible") void refreshData();
    }

    document.addEventListener("visibilitychange", refreshWhenVisible);
    window.addEventListener("focus", refreshWhenVisible);

    return () => {
      document.removeEventListener("visibilitychange", refreshWhenVisible);
      window.removeEventListener("focus", refreshWhenVisible);
    };
  }, []);

  useEffect(() => {
    const timer = window.setInterval(() => {
      setNow(Date.now());
    }, 30000);

    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    const dueTasks = workTasks.filter((item) => {
      if (!item.reminderAt || item.status === "Closed") return false;
      const dueAt = new Date(item.reminderAt).getTime();
      return Number.isFinite(dueAt) && dueAt <= now && !triggeredReminderIds.includes(reminderTriggerId(item));
    });

    if (dueTasks.length === 0) return;

    const nextTriggeredIds = [...triggeredReminderIds, ...dueTasks.map(reminderTriggerId)];
    setTriggeredReminderIds(nextTriggeredIds);
    void setMeta(triggeredReminderMetaKey, nextTriggeredIds);
    dueTasks.forEach((item) => {
      void notifyReminder(item);
    });
  }, [now, triggeredReminderIds, workTasks]);

  function updateProject(nextProjectId: ProjectId) {
    setProjectId(nextProjectId);
    const nextTaskId = projects[nextProjectId].tasks[0].id;
    setTaskId(nextTaskId);
    setRequirements(taskRequirements(nextProjectId, nextTaskId));
    setGptPrompt("");
    setResult("");
    setMessage("");
  }

  function updateRequirement<K extends keyof Requirements>(key: K, value: Requirements[K]) {
    const next = { ...requirements, [key]: value };
    setRequirements(next);
    updateActiveWorkTask("requirements", next);
  }

  function selectTemplate(nextTaskId: string) {
    setTaskId(nextTaskId);
    const nextTemplate = project.tasks.find((item) => item.id === nextTaskId) ?? project.tasks[0];
    const nextRequirements = taskRequirements(projectId, nextTaskId);
    setRequirements(nextRequirements);
    if (activeWorkTask) {
      updateWorkTask(activeWorkTask.id, "templateId", nextTaskId);
      updateWorkTask(activeWorkTask.id, "category", nextTemplate.category);
      updateWorkTask(activeWorkTask.id, "requirements", nextRequirements);
    }
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
    if (!quickTaskTitle.trim() && !quickTaskDetails.trim() && quickChecklistItems.length === 0) {
      setMessage("Add a quick title, note, or checklist item first.");
      return;
    }

    createTask({
      title: quickTaskTitle.trim() || "Mobile note",
      details: quickTaskDetails.trim(),
      taskProjectId: projectId,
      templateId: projects[projectId].tasks[0].id,
      category: "Mobile",
      checklist: quickChecklistItems,
    });
    setQuickTaskTitle("");
    setQuickTaskDetails("");
    setQuickChecklistDraft("");
    setQuickChecklistItems([]);
    setMessage("Mobile task captured into the shared dashboard.");
  }

  function addQuickChecklistItem() {
    if (!quickChecklistDraft.trim()) return;
    setQuickChecklistItems((current) => [
      ...current,
      { id: createId(), text: quickChecklistDraft.trim(), done: false },
    ]);
    setQuickChecklistDraft("");
  }

  function removeQuickChecklistItem(id: string) {
    setQuickChecklistItems((current) => current.filter((item) => item.id !== id));
  }

  function createTask({
    title,
    details,
    taskProjectId,
    templateId,
    category,
    checklist = [],
  }: {
    title: string;
    details: string;
    taskProjectId: ProjectId;
    templateId: string;
    category: string;
    checklist?: ChecklistItem[];
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
      checklist,
      input: "",
      assets: [],
      requirements: taskRequirements(taskProjectId, templateId),
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

  function addChecklistItem() {
    if (!activeWorkTask || !quickChecklistItem.trim()) return;
    const nextChecklist = [
      ...activeWorkTask.checklist,
      { id: createId(), text: quickChecklistItem.trim(), done: false },
    ];
    updateWorkTask(activeWorkTask.id, "checklist", nextChecklist);
    setQuickChecklistItem("");
  }

  function updateChecklistItem(itemId: string, patch: Partial<ChecklistItem>) {
    if (!activeWorkTask) return;
    updateWorkTask(
      activeWorkTask.id,
      "checklist",
      activeWorkTask.checklist.map((item) => (item.id === itemId ? { ...item, ...patch } : item)),
    );
  }

  function removeChecklistItem(itemId: string) {
    if (!activeWorkTask) return;
    updateWorkTask(
      activeWorkTask.id,
      "checklist",
      activeWorkTask.checklist.filter((item) => item.id !== itemId),
    );
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
    if (key === "reminderAt") forgetTriggeredReminder(id);
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

  function openMobileTask(item: WorkTask, target: "details" | "ai") {
    openWorkTask(item);
    setViewMode("work");
    window.setTimeout(() => {
      document.getElementById(target === "details" ? "task-details-section" : "ai-workspace-section")?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    }, 80);
  }

  function removeWorkTask(id: string) {
    setWorkTasks((current) => current.filter((item) => item.id !== id));
    void deleteTask(id);
    if (activeWorkTaskId === id) setActiveWorkTaskId("");
  }

  function saveTaskOnly() {
    if (!activeWorkTask) return;
    const reminderAt = reminderDrafts[activeWorkTask.id] ?? activeWorkTask.reminderAt;
    const updated: WorkTask = {
      ...activeWorkTask,
      reminderAt,
      updatedAt: new Date().toISOString(),
    };
    setWorkTasks((current) => current.map((item) => (item.id === activeWorkTask.id ? updated : item)));
    void saveTask(updated);
    forgetTriggeredReminder(activeWorkTask.id);
    setNow(Date.now());
    setMessage("Task saved, including reminder date.");
  }

  function clearReminder(id: string) {
    setReminderDrafts((current) => ({ ...current, [id]: "" }));
    updateWorkTask(id, "reminderAt", "");
    forgetTriggeredReminder(id);
    setMessage("Reminder cleared.");
  }

  function updateReminder(id: string, value: string) {
    setReminderDrafts((current) => ({ ...current, [id]: value }));
  }

  function saveReminder(id: string) {
    const taskToUpdate = workTasks.find((item) => item.id === id);
    if (!taskToUpdate) return;
    const value = reminderDrafts[id] ?? taskToUpdate.reminderAt;
    const updated: WorkTask = {
      ...taskToUpdate,
      reminderAt: value,
      updatedAt: new Date().toISOString(),
    };
    setWorkTasks((current) => current.map((item) => (item.id === id ? updated : item)));
    void saveTask(updated);
    forgetTriggeredReminder(id);
    setNow(Date.now());
    setMessage(value ? "Reminder saved." : "Reminder cleared.");
  }

  function scheduleReminderToday(id: string) {
    const value = localDatetimeValue(nextPlanningHour(now));
    setReminderDrafts((current) => ({ ...current, [id]: value }));
    updateWorkTask(id, "reminderAt", value);
    forgetTriggeredReminder(id);
    setNow(Date.now());
    setMessage("Reminder scheduled for today.");
  }

  function reminderValue(task: WorkTask) {
    return reminderDrafts[task.id] ?? task.reminderAt;
  }

  async function enableNotifications() {
    if (typeof Notification === "undefined") {
      setNotificationPermission("unsupported");
      setMessage("This browser does not support reminder notifications.");
      return;
    }

    const permission = await Notification.requestPermission();
    setNotificationPermission(permission);
    setMessage(permission === "granted" ? "Reminder notifications enabled for this device." : "Reminder notifications were not enabled.");
  }

  function forgetTriggeredReminder(taskId: string) {
    setTriggeredReminderIds((current) => {
      const next = current.filter((item) => !item.startsWith(`${taskId}:`));
      void setMeta(triggeredReminderMetaKey, next);
      return next;
    });
  }

  function persistNotes(nextNotes: AppNote[]) {
    setNotes(nextNotes);
    void saveNotes(nextNotes);
  }

  function createNote() {
    if (!noteDraft.title.trim() && !noteDraft.content.trim()) {
      setMessage("Add a note title or information before saving.");
      return;
    }

    const timestamp = new Date().toISOString();
    const newNote: AppNote = {
      id: createId(),
      projectId,
      title: noteDraft.title.trim() || "Untitled note",
      entries: noteDraft.content.trim()
        ? [{ id: createId(), content: noteDraft.content.trim(), createdAt: timestamp, updatedAt: timestamp }]
        : [],
      pinned: false,
      createdAt: timestamp,
      updatedAt: timestamp,
    };
    persistNotes([newNote, ...notes]);
    setNoteDraft({ title: "", content: "" });
    setMessage("Note saved.");
  }

  function updateNote(id: string, patch: Partial<AppNote>) {
    persistNotes(notes.map((item) => (item.id === id ? { ...item, ...patch, updatedAt: new Date().toISOString() } : item)));
  }

  function addNoteEntry(id: string) {
    const timestamp = new Date().toISOString();
    persistNotes(
      notes.map((item) =>
        item.id === id
          ? {
              ...item,
              entries: [{ id: createId(), content: "", createdAt: timestamp, updatedAt: timestamp }, ...item.entries],
              updatedAt: timestamp,
            }
          : item,
      ),
    );
  }

  function updateNoteEntry(noteId: string, entryId: string, content: string) {
    const timestamp = new Date().toISOString();
    persistNotes(
      notes.map((item) =>
        item.id === noteId
          ? {
              ...item,
              entries: item.entries.map((entry) => (entry.id === entryId ? { ...entry, content, updatedAt: timestamp } : entry)),
              updatedAt: timestamp,
            }
          : item,
      ),
    );
  }

  function removeNoteEntry(noteId: string, entryId: string) {
    const timestamp = new Date().toISOString();
    persistNotes(
      notes.map((item) =>
        item.id === noteId
          ? {
              ...item,
              entries: item.entries.filter((entry) => entry.id !== entryId),
              updatedAt: timestamp,
            }
          : item,
      ),
    );
  }

  function removeNote(id: string) {
    persistNotes(notes.filter((item) => item.id !== id));
    setMessage("Note deleted.");
  }

  async function handleFiles(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    const nextAssets = await Promise.all(
      files.map(async (file) => {
        if (file.size > maxUploadSizeBytes) {
          return {
            id: createId(),
            name: file.name,
            type: "file" as const,
            content: `Skipped for safety: ${file.name} is larger than ${Math.round(maxUploadSizeBytes / 1024 / 1024)} MB. Paste the relevant text into the input box instead.`,
          };
        }

        if (file.type.startsWith("image/")) {
          return {
            id: createId(),
            name: file.name,
            type: "image" as const,
            content: await readAsDataUrl(file),
          };
        }

        const isDocx = /\.docx$/i.test(file.name);
        const isPptx = /\.pptx$/i.test(file.name);
        const canReadText = file.type.startsWith("text/") || /\.(md|txt|csv|json|html)$/i.test(file.name);
        const content = isDocx
          ? await readDocxText(file)
          : isPptx
            ? await readPptxText(file)
          : canReadText
            ? await file.text()
            : "This file is attached but cannot be read in the browser. For PDF or scanned files, paste the important text into the input box.";
        const safeContent = content.length > maxTextAssetCharacters
          ? `${content.slice(0, maxTextAssetCharacters)}\n\n[Content truncated at ${maxTextAssetCharacters.toLocaleString()} characters for browser performance.]`
          : content;

        return {
          id: createId(),
          name: file.name,
          type: canReadText || isDocx || isPptx ? ("text" as const) : ("file" as const),
          content: safeContent,
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

    const prompt = buildFullLlmPrompt(project.name, project.context, task.label, activeWorkTask?.checklist ?? [], input, assets, requirements);
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
    const confirmed = window.confirm(
      "Clear the current AI workspace? This removes the source input, uploaded assets, prompt, output, and output requirements for the currently selected task. It will not delete tasks, notes, reminders, or saved history.",
    );
    if (!confirmed) return;

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
      <aside className="app-rail" aria-label="Primary navigation">
        <div className="brand-mark">
          <span>AI</span>
          <div>
            <strong>WorkOS</strong>
            <small>Tasks + notes</small>
          </div>
        </div>
        <nav className="rail-nav">
          <button className={viewMode === "home" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("home")} type="button">
            <LayoutDashboard size={17} />
            Home
          </button>
          <button className={viewMode === "work" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("work")} type="button">
            <ListTodo size={17} />
            Work
          </button>
          <button className={viewMode === "notes" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("notes")} type="button">
            <StickyNote size={17} />
            Notes
          </button>
          <button className={viewMode === "reminders" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("reminders")} type="button">
            <CalendarClock size={17} />
            Reminders
          </button>
          <button className={viewMode === "projects" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("projects")} type="button">
            <BarChart3 size={17} />
            Projects
          </button>
          <button className={viewMode === "mobile" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("mobile")} type="button">
            <Smartphone size={17} />
            Mobile
          </button>
        </nav>
        <div className="rail-actions">
          <button className="ghost-button" disabled={syncing} onClick={() => void refreshData()} type="button" title="Pull latest data from Supabase">
            <History size={16} />
            {syncing ? "Syncing" : "Sync"}
          </button>
          <button className="danger-button rail-clear-button" onClick={clearWorkspace} type="button" title="Clear only the selected task AI workspace">
            <Trash2 size={16} />
            Clear AI workspace
          </button>
        </div>
      </aside>

      <section className="app-main">
      <section className="topbar">
        <div>
          <p className="eyebrow">AI-powered work command center</p>
          <h1>{viewMode === "home" ? "Today, tasks, notes, and AI drafts" : "Project work assistant"}</h1>
        </div>
        <div className="topbar-actions hero-actions">
          <button className="primary-button ai-action" onClick={() => setViewMode("work")} type="button">
            <WandSparkles size={16} />
            Open AI studio
          </button>
          <button className="ghost-button" onClick={() => setViewMode("notes")} type="button">
            <Plus size={16} />
            Capture note
          </button>
        </div>
      </section>

      {message && (
        <div className="sync-banner" role="status">
          {message}
        </div>
      )}

      {viewMode === "home" && (
        <section className="home-page">
          <section className="hero-panel">
            <div>
              <p className="eyebrow">Live workspace</p>
              <h2>Move the day forward</h2>
              <p className="context-copy">
                A faster launchpad for urgent work, reminders, notes, project status, and AI-ready drafting.
              </p>
            </div>
            <div className="hero-stat-grid">
              <button className="hero-stat urgent" onClick={() => setViewMode("reminders")} type="button">
                <strong>{reminderPlanner.overdue.length + reminderPlanner.today.length}</strong>
                <span>Due now</span>
              </button>
              <button className="hero-stat progress" onClick={() => setViewMode("work")} type="button">
                <strong>{summary.inProgress}</strong>
                <span>In progress</span>
              </button>
              <button className="hero-stat notes" onClick={() => setViewMode("notes")} type="button">
                <strong>{noteEntryCount}</strong>
                <span>Note entries</span>
              </button>
              <button className="hero-stat ai" onClick={() => setViewMode("work")} type="button">
                <strong>{savedOutputs.length}</strong>
                <span>AI outputs</span>
              </button>
            </div>
          </section>

          <section className="home-grid">
            <section className="panel focus-panel">
              <div className="result-header">
                <h2>
                  <ListTodo size={18} />
                  Focus queue
                </h2>
                <button className="ghost-button" onClick={() => setViewMode("work")} type="button">
                  View all
                </button>
              </div>
              <div className="all-task-list">
                {activeFocusTasks.map((item) => (
                  <button key={item.id} className={taskClassName(item, activeWorkTaskId)} onClick={() => openWorkTask(item)} type="button">
                    <span>
                      <strong>{item.title}</strong>
                      <small>{projects[item.projectId].name} - {item.category} - {item.status}</small>
                    </span>
                    <small>{taskDateLabel(item)}</small>
                  </button>
                ))}
                {activeFocusTasks.length === 0 && <p className="empty">No active tasks. Lovely clean slate.</p>}
              </div>
            </section>

            <section className="panel pulse-panel">
              <div className="result-header">
                <h2>
                  <CalendarClock size={18} />
                  Today
                </h2>
                <button className="ghost-button" onClick={() => setViewMode("reminders")} type="button">
                  Planner
                </button>
              </div>
              <div className="reminder-list">
                {[...reminderPlanner.overdue, ...reminderPlanner.today].slice(0, 5).map((item) => (
                  <button className={`reminder-title home-reminder ${reminderState(item, now)}`} key={item.id} onClick={() => openWorkTask(item)} type="button">
                    <strong>{item.title}</strong>
                    <span>{projects[item.projectId].name} - {reminderLabel(item, now)}</span>
                  </button>
                ))}
                {reminderPlanner.overdue.length + reminderPlanner.today.length === 0 && <p className="empty">No reminders due today.</p>}
              </div>
            </section>

            <section className="panel notes-snapshot">
              <div className="result-header">
                <h2>
                  <StickyNote size={18} />
                  Recent notes
                </h2>
                <button className="ghost-button" onClick={() => setViewMode("notes")} type="button">
                  Open notes
                </button>
              </div>
              <div className="note-snapshot-list">
                {recentNotes.map((note) => (
                  <button className="note-snapshot" key={note.id} onClick={() => setViewMode("notes")} type="button">
                    <span className={`project-chip project-${note.projectId}`}>{projects[note.projectId].name}</span>
                    <strong>{note.title}</strong>
                    <small>{note.entries[0]?.content || "No entries yet"}</small>
                  </button>
                ))}
                {recentNotes.length === 0 && <p className="empty">Capture notes and reference information here.</p>}
              </div>
            </section>

            <section className="panel ai-panel">
              <div className="result-header">
                <h2>
                  <Sparkles size={18} />
                  AI studio
                </h2>
              </div>
              <p className="context-copy">Turn task notes, uploaded documents, PowerPoint text, and checklists into document drafts, summaries, reports, and emails.</p>
              <div className="ai-studio-actions">
                <button className="primary-button ai-action" onClick={() => setViewMode("work")} type="button">
                  <WandSparkles size={16} />
                  Prepare prompt
                </button>
                <button className="ghost-button" onClick={() => setViewMode("projects")} type="button">
                  <BarChart3 size={16} />
                  Project view
                </button>
              </div>
            </section>
          </section>
        </section>
      )}

      {viewMode === "mobile" && (
        <section className="mobile-shell">
          <section className="mobile-tabs" aria-label="Mobile views">
            <button className={mobileSection === "tasks" ? "active" : ""} onClick={() => setMobileSection("tasks")} type="button">
              <ListTodo size={16} />
              Tasks
            </button>
            <button className={mobileSection === "capture" ? "active" : ""} onClick={() => setMobileSection("capture")} type="button">
              <Plus size={16} />
              Capture
            </button>
            <button className={mobileSection === "reminders" ? "active" : ""} onClick={() => setMobileSection("reminders")} type="button">
              <CalendarClock size={16} />
              Today
            </button>
          </section>

          {mobileSection === "capture" && (
          <section className="panel mobile-capture">
            <h2>
              <Smartphone size={18} />
              Capture
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
            <div className="mobile-checklist-capture">
              <label>
                Checklist items
                <div className="checklist-add">
                  <input
                    value={quickChecklistDraft}
                    onChange={(event) => setQuickChecklistDraft(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") addQuickChecklistItem();
                    }}
                    placeholder="Add item to tick off"
                  />
                  <button className="primary-button" onClick={addQuickChecklistItem} type="button" title="Add checklist item">
                    <Plus size={16} />
                  </button>
                </div>
              </label>
              {quickChecklistItems.length > 0 && (
                <div className="checklist-list">
                  {quickChecklistItems.map((item) => (
                    <div className="checklist-item" key={item.id}>
                      <input checked={false} readOnly type="checkbox" />
                      <input
                        value={item.text}
                        onChange={(event) =>
                          setQuickChecklistItems((current) =>
                            current.map((currentItem) => currentItem.id === item.id ? { ...currentItem, text: event.target.value } : currentItem),
                          )
                        }
                      />
                      <button className="ghost-button" onClick={() => removeQuickChecklistItem(item.id)} type="button" title="Remove checklist item">
                        <Trash2 size={15} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <button className="primary-button" onClick={createQuickTask} type="button">
              <Save size={16} />
              Save mobile task
            </button>
          </section>
          )}

          {mobileSection === "tasks" && (
          <section className="mobile-task-view">
            <div className="mobile-section-header">
              <h2>
                <ListTodo size={18} />
                Tasks
              </h2>
              <button className="primary-button" onClick={() => setMobileSection("capture")} type="button">
                <Plus size={15} />
                New
              </button>
            </div>
            <div className="mobile-project-strip" aria-label="Mobile project selector">
              {mobileProjectGroups.map((item) => (
                <button
                  className={projectId === item.projectId ? "active" : ""}
                  key={item.projectId}
                  onClick={() => updateProject(item.projectId)}
                  type="button"
                >
                  <span>{projects[item.projectId].name}</span>
                  <strong>{item.total}</strong>
                </button>
              ))}
            </div>
            {mobileCurrentProjectGroup && (
              <div className="mobile-status-groups">
                <div className="mobile-project-title">
                  <strong>{projects[mobileCurrentProjectGroup.projectId].name}</strong>
                  <span>{mobileCurrentProjectGroup.total} tasks</span>
                </div>
                {mobileStatusOrder.map((status) => (
                  <MobileStatusSection
                    activeWorkTaskId={activeWorkTaskId}
                    items={mobileCurrentProjectGroup.byStatus[status]}
                    key={status}
                    now={now}
                    onOpen={openMobileTask}
                    status={status}
                  />
                ))}
              </div>
            )}
          </section>
          )}

          {mobileSection === "reminders" && (
          <section className="panel">
            <div className="mobile-section-header">
              <h2>
                <CalendarClock size={18} />
                Today
              </h2>
              <div className="mobile-section-actions">
                <button className="ghost-button" onClick={() => void enableNotifications()} type="button">
                  <BellRing size={15} />
                  Alerts
                </button>
                <button className="ghost-button" onClick={() => setViewMode("reminders")} type="button">
                  Planner
                </button>
              </div>
            </div>
            <div className="mobile-task-list">
              {[...reminderPlanner.overdue, ...reminderPlanner.today].map((item) => (
                <article key={item.id} className={`${taskClassName(item, activeWorkTaskId)} mobile-task-card`}>
                  <div>
                    <strong>{item.title}</strong>
                    <small>{projects[item.projectId].name}</small>
                    <small>{reminderLabel(item, now)}</small>
                  </div>
                  <div className="mobile-task-actions">
                    <button className="ghost-button" onClick={() => openMobileTask(item, "details")} type="button">
                      <ListTodo size={15} />
                      Details
                    </button>
                    <button className="primary-button" onClick={() => scheduleReminderToday(item.id)} type="button">
                      <CalendarClock size={15} />
                      Replan
                    </button>
                  </div>
                </article>
              ))}
              {reminderPlanner.overdue.length + reminderPlanner.today.length === 0 && (
                <p className="empty">No overdue or today reminders.</p>
              )}
            </div>
          </section>
          )}

          <section className="panel mobile-full-panel">
            <button className="ghost-button mobile-full-toggle" onClick={() => setMobileFullOpen((current) => !current)} type="button">
              <BriefcaseBusiness size={16} />
              {mobileFullOpen ? "Close full functionality" : "Open full functionality"}
            </button>
            {mobileFullOpen && (
              <div className="mobile-full-content">
                <p className="context-copy">Use the main work area for templates, full task editing, prompts, uploads, output saving, and project history.</p>
                <button className="primary-button" onClick={() => setViewMode("work")} type="button">
                  <ArrowRight size={16} />
                  Go to full work view
                </button>
                <button className="ghost-button" onClick={() => setViewMode("reminders")} type="button">
                  <CalendarClock size={16} />
                  Open reminder planner
                </button>
              </div>
            )}
          </section>
        </section>
      )}

      {viewMode === "reminders" && (
        <section className="reminder-page">
          <section className="panel reminder-hero">
            <div>
              <p className="eyebrow">Reminder planner</p>
              <h2>
                <CalendarClock size={18} />
                Plan today and upcoming work
              </h2>
              <p className="context-copy">
                Review reminders here and enable browser notifications on each device that should alert you.
              </p>
            </div>
            <div className="reminder-hero-actions">
              <div className="reminder-metrics">
                <span><strong>{reminderPlanner.overdue.length}</strong> overdue</span>
                <span><strong>{reminderPlanner.today.length}</strong> today</span>
                <span><strong>{reminderPlanner.upcoming.length}</strong> upcoming</span>
              </div>
              <button
                className={notificationPermission === "granted" ? "ghost-button success-button" : "primary-button"}
                disabled={notificationPermission === "unsupported"}
                onClick={() => void enableNotifications()}
                type="button"
              >
                <BellRing size={16} />
                {notificationPermission === "granted" ? "Notifications on" : "Enable notifications"}
              </button>
            </div>
          </section>

          <section className="reminder-planner-layout">
            <section className="reminder-main-board">
              <ReminderColumn
                emptyText="No overdue reminders."
                items={reminderPlanner.overdue}
                now={now}
                onClear={clearReminder}
                onOpen={(item) => {
                  openWorkTask(item);
                  setViewMode("work");
                }}
                onSave={saveReminder}
                onSchedule={updateReminder}
                reminderValue={reminderValue}
                title="Overdue"
              />
              <ReminderColumn
                emptyText="Nothing planned for today."
                items={reminderPlanner.today}
                now={now}
                onClear={clearReminder}
                onOpen={(item) => {
                  openWorkTask(item);
                  setViewMode("work");
                }}
                onSave={saveReminder}
                onSchedule={updateReminder}
                reminderValue={reminderValue}
                title="Today"
              />
              <ReminderColumn
                emptyText="No upcoming reminders."
                items={reminderPlanner.upcoming}
                now={now}
                onClear={clearReminder}
                onOpen={(item) => {
                  openWorkTask(item);
                  setViewMode("work");
                }}
                onSave={saveReminder}
                onSchedule={updateReminder}
                reminderValue={reminderValue}
                title="Upcoming"
              />
            </section>

            <section className="panel reminder-unscheduled-panel">
              <div className="result-header">
                <h2>
                  <ListTodo size={18} />
                  Planning queue
                </h2>
                <span className="subtle-count">{reminderPlanner.unscheduled.length}</span>
              </div>
              <div className="reminder-list">
                {reminderPlanner.unscheduled.map((item) => (
                  <div className="reminder-row compact" key={item.id}>
                    <button
                      className="reminder-title"
                      onClick={() => {
                        openWorkTask(item);
                        setViewMode("work");
                      }}
                      type="button"
                    >
                      <strong>{item.title}</strong>
                      <span>{projects[item.projectId].name} - {item.category} - {item.priority}</span>
                    </button>
                    <div className="reminder-actions">
                      <button className="ghost-button" onClick={() => scheduleReminderToday(item.id)} type="button">
                        <CalendarClock size={15} />
                        Today
                      </button>
                      <button className="primary-button" onClick={() => saveReminder(item.id)} type="button">
                        <Save size={15} />
                        Save
                      </button>
                    </div>
                    <label>
                      Reminder
                      <input value={reminderValue(item)} onChange={(event) => updateReminder(item.id, event.target.value)} type="datetime-local" />
                    </label>
                  </div>
                ))}
                {reminderPlanner.unscheduled.length === 0 && <p className="empty">Every open task has a reminder, or there are no open tasks.</p>}
              </div>
            </section>
          </section>
        </section>
      )}

      {viewMode === "notes" && (
        <section className="notes-page">
          <section className="panel notes-compose">
            <div className="result-header">
              <h2>
                <StickyNote size={18} />
                Important notes
              </h2>
              <span className="subtle-count">{notes.reduce((total, note) => total + note.entries.length, 0)} entries</span>
            </div>
            <div className="note-compose-grid">
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
                Title
                <input
                  value={noteDraft.title}
                  onChange={(event) => setNoteDraft((current) => ({ ...current, title: event.target.value }))}
                  placeholder="Reference, decision, contact detail..."
                />
              </label>
            </div>
            <label>
              Information
              <textarea
                className="small-textarea"
                value={noteDraft.content}
                onChange={(event) => setNoteDraft((current) => ({ ...current, content: event.target.value }))}
                placeholder="Store important information that is not yet a task."
              />
            </label>
            <button className="primary-button" onClick={createNote} type="button">
              <Save size={16} />
              Save note
            </button>
          </section>

          <section className="notes-grid">
            {[...notes]
              .sort((a, b) => Number(b.pinned) - Number(a.pinned) || dateValue(b.updatedAt) - dateValue(a.updatedAt))
              .map((note) => (
                <article className={note.pinned ? "note-card pinned" : "note-card"} key={note.id}>
                  <div className="note-card-header">
                    <span className={`project-chip project-${note.projectId}`}>{projects[note.projectId].name}</span>
                    <button className="ghost-button icon-button" onClick={() => updateNote(note.id, { pinned: !note.pinned })} type="button" title={note.pinned ? "Unpin note" : "Pin note"}>
                      <Pin size={15} />
                    </button>
                  </div>
                  <input
                    className="note-title-input"
                    value={note.title}
                    onChange={(event) => updateNote(note.id, { title: event.target.value })}
                  />
                  <div className="note-entry-list">
                    {note.entries.map((entry) => (
                      <div className="note-entry" key={entry.id}>
                        <textarea
                          className="note-content-input"
                          value={entry.content}
                          onChange={(event) => updateNoteEntry(note.id, entry.id, event.target.value)}
                          placeholder="Add note detail..."
                        />
                        <div className="note-entry-footer">
                          <small>{new Date(entry.updatedAt).toLocaleString()}</small>
                          <button className="ghost-button icon-button" onClick={() => removeNoteEntry(note.id, entry.id)} type="button" title="Delete this note entry">
                            <Trash2 size={15} />
                          </button>
                        </div>
                      </div>
                    ))}
                    {note.entries.length === 0 && <p className="empty compact-empty">No entries under this title yet.</p>}
                  </div>
                  <div className="note-footer">
                    <small>{note.entries.length} {note.entries.length === 1 ? "entry" : "entries"}</small>
                    <button className="ghost-button" onClick={() => addNoteEntry(note.id)} type="button">
                      <Plus size={15} />
                      Add entry
                    </button>
                    <button className="ghost-button icon-button" onClick={() => removeNote(note.id)} type="button" title="Delete note">
                      <Trash2 size={15} />
                    </button>
                  </div>
                </article>
              ))}
            {notes.length === 0 && <p className="empty">Save useful reference information here. Notes sync with the same database as your tasks.</p>}
          </section>
        </section>
      )}

      {viewMode === "projects" && (
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
        <div className="summary-card summary-later">
          <strong>{summary.toDoLater}</strong>
          <span>To do later</span>
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
            <button
              key={item.projectId}
              className={projectId === item.projectId ? "project-dashboard-card active" : "project-dashboard-card"}
              onClick={() => {
                updateProject(item.projectId);
              }}
              type="button"
            >
              <strong>{projects[item.projectId].name}</strong>
              <span>{item.total} tasks</span>
              <div className="mini-status-grid">
                <small>Open {item.open}</small>
                <small>Progress {item.inProgress}</small>
                <small>Blocked {item.blocked}</small>
                <small>Later {item.toDoLater}</small>
                <small>Closed {item.closed}</small>
              </div>
            </button>
          ))}
        </div>
      </section>

      <section className="panel dashboard-panel">
        <div className="result-header">
          <h2>
            <BriefcaseBusiness size={18} />
            {project.name} details
          </h2>
          <button className="primary-button" onClick={() => setViewMode("work")} type="button">
            <ArrowRight size={16} />
            Open project work
          </button>
        </div>
        <p className="context-copy">{project.context}</p>
        {selectedProjectStats && (
          <div className="project-stat-list project-stat-columns">
            {taskStatuses.map((status) => (
              <div className={`project-stat-row status-${statusSlug(status)}`} key={status}>
                <span>{status}</span>
                <strong>{selectedProjectStats[statusKey(status)]}</strong>
              </div>
            ))}
          </div>
        )}
        <div className="project-detail-list">
          {sortedProjectTasks.slice(0, 8).map((item) => (
            <button key={item.id} className={taskClassName(item, activeWorkTaskId)} onClick={() => openWorkTask(item)} type="button">
              <span>
                <strong>{item.title}</strong>
                <small>{item.category} - {item.priority} - {item.status}</small>
              </span>
              <small>{taskDateLabel(item)}</small>
            </button>
          ))}
          {sortedProjectTasks.length === 0 && <p className="empty">No tasks have been created for this project yet.</p>}
        </div>
      </section>
        </>
      )}

      {viewMode === "work" && (
        <>

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
              Active project
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
              <ListTodo size={18} />
              Project tasks
            </h2>
            <div className="project-stat-list">
              <div className="project-stat-row status-open"><span>Open</span><strong>{selectedProjectStats?.open ?? 0}</strong></div>
              <div className="project-stat-row status-in-progress"><span>In progress</span><strong>{selectedProjectStats?.inProgress ?? 0}</strong></div>
              <div className="project-stat-row status-blocked"><span>Blocked</span><strong>{selectedProjectStats?.blocked ?? 0}</strong></div>
              <div className="project-stat-row status-to-do-later"><span>To do later</span><strong>{selectedProjectStats?.toDoLater ?? 0}</strong></div>
              <div className="project-stat-row status-closed"><span>Closed</span><strong>{selectedProjectStats?.closed ?? 0}</strong></div>
            </div>
          </section>

          <details className="panel template-drawer">
            <summary>
              <span>
                <Sparkles size={18} />
                Task templates
              </span>
              <small>{task.label}</small>
            </summary>
            <div className="task-list">
              {project.tasks.map((item) => (
                <button key={item.id} className={taskId === item.id ? "task-card selected" : "task-card"} onClick={() => selectTemplate(item.id)} type="button">
                  <strong>{item.label}</strong>
                  <span>{item.category} - {item.description}</span>
                </button>
              ))}
            </div>
          </details>

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
                {project.name} tasks
              </h2>
              <div className="new-task-control">
                <label>
                  Template
                  <select value={taskId} onChange={(event) => selectTemplate(event.target.value)}>
                    {project.tasks.map((item) => (
                      <option key={item.id} value={item.id}>
                        {item.label}
                      </option>
                    ))}
                  </select>
                </label>
                <button className="primary-button" onClick={createWorkTask} type="button">
                  <ListTodo size={16} />
                  New task
                </button>
              </div>
            </div>
            {sortedProjectTasks.length === 0 ? (
              <p className="empty">Create your first task for {project.name}. It can be a document, summary, email, or a checklist-style task.</p>
            ) : (
              <div className="all-task-list">
                {sortedProjectTasks.map((item) => (
                  <button key={item.id} className={taskClassName(item, activeWorkTaskId)} onClick={() => openWorkTask(item)} type="button">
                    <span>
                      <strong>{item.title}</strong>
                      <small>{item.category} - {item.priority} - {item.status}</small>
                    </span>
                    <small>{taskDateLabel(item)}</small>
                  </button>
                ))}
              </div>
            )}
          </section>

          {activeWorkTask ? (
            <section className="panel" id="task-details-section">
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
                <span className={`status-pill status-${statusSlug(activeWorkTask.status)}`}>{activeWorkTask.status}</span>
                <span className={`priority-pill priority-${activeWorkTask.priority.toLowerCase()}`}>{activeWorkTask.priority}</span>
                {activeWorkTask.reminderAt && (
                  <span className="reminder-pill">
                    <CalendarClock size={14} />
                    {new Date(activeWorkTask.reminderAt).toLocaleString()}
                  </span>
                )}
              </div>
              <TaskEditor
                task={activeWorkTask}
                onChange={updateWorkTask}
                onReminderChange={updateReminder}
                onReminderSave={saveReminder}
                reminderValue={reminderValue(activeWorkTask)}
                onTemplateChange={(nextProjectId, nextTemplateId) => {
                  const nextRequirements = taskRequirements(nextProjectId, nextTemplateId);
                  const nextTemplate = projects[nextProjectId].tasks.find((item) => item.id === nextTemplateId) ?? projects[nextProjectId].tasks[0];
                  setTaskId(nextTemplateId);
                  setRequirements(nextRequirements);
                  updateWorkTask(activeWorkTask.id, "templateId", nextTemplateId);
                  updateWorkTask(activeWorkTask.id, "category", nextTemplate.category);
                  updateWorkTask(activeWorkTask.id, "requirements", nextRequirements);
                }}
              />
              <label>
                Task notes and details
                <textarea
                  className="small-textarea"
                  value={activeWorkTask.details}
                  onChange={(event) => updateWorkTask(activeWorkTask.id, "details", event.target.value)}
                  placeholder="Record task background, decisions, follow-up notes, or status details here."
                />
              </label>
              <section className="checklist-panel">
                <div className="result-header">
                  <h3>
                    <CheckSquare size={17} />
                    Checklist
                  </h3>
                  <span>{activeWorkTask.checklist.filter((item) => item.done).length}/{activeWorkTask.checklist.length} done</span>
                </div>
                <div className="checklist-add">
                  <input
                    value={quickChecklistItem}
                    onChange={(event) => setQuickChecklistItem(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") addChecklistItem();
                    }}
                    placeholder="Add shopping item, follow-up, or subtask"
                  />
                  <button className="primary-button" onClick={addChecklistItem} type="button" title="Add checklist item">
                    <Plus size={16} />
                  </button>
                </div>
                <div className="checklist-list">
                  {activeWorkTask.checklist.map((item) => (
                    <div className={item.done ? "checklist-item done" : "checklist-item"} key={item.id}>
                      <input
                        aria-label={`Complete ${item.text}`}
                        checked={item.done}
                        onChange={(event) => updateChecklistItem(item.id, { done: event.target.checked })}
                        type="checkbox"
                      />
                      <input value={item.text} onChange={(event) => updateChecklistItem(item.id, { text: event.target.value })} />
                      <button className="ghost-button" onClick={() => removeChecklistItem(item.id)} type="button" title="Remove checklist item">
                        <Trash2 size={15} />
                      </button>
                    </div>
                  ))}
                  {activeWorkTask.checklist.length === 0 && <p className="empty">Add items here for shopping lists, subtasks, or simple tick-off work.</p>}
                </div>
              </section>
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
              <section className="panel" id="ai-workspace-section">
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
                  Upload documents, presentations, or images
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
      </section>
    </main>
  );
}

function TaskEditor({
  task,
  onChange,
  onReminderChange,
  onReminderSave,
  onTemplateChange,
  reminderValue,
}: {
  task: WorkTask;
  onChange: <K extends keyof WorkTask>(id: string, key: K, value: WorkTask[K]) => void;
  onReminderChange: (id: string, value: string) => void;
  onReminderSave: (id: string) => void;
  onTemplateChange: (projectId: ProjectId, templateId: string) => void;
  reminderValue: string;
}) {
  const availableTemplates = projects[task.projectId].tasks;

  return (
    <div className="task-editor">
      <label>
        Project
        <select
          value={task.projectId}
          onChange={(event) => {
            const nextProjectId = event.target.value as ProjectId;
            const nextTemplateId = projects[nextProjectId].tasks[0].id;
            onChange(task.id, "projectId", nextProjectId);
            onTemplateChange(nextProjectId, nextTemplateId);
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
        <select value={task.templateId} onChange={(event) => onTemplateChange(task.projectId, event.target.value)}>
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
        <div className="inline-save-field">
          <input value={reminderValue} onChange={(event) => onReminderChange(task.id, event.target.value)} type="datetime-local" />
          <button className="primary-button" onClick={() => onReminderSave(task.id)} type="button">
            <Save size={15} />
            Save
          </button>
        </div>
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

function ReminderColumn({
  emptyText,
  items,
  now,
  onClear,
  onOpen,
  onSave,
  onSchedule,
  reminderValue,
  title,
}: {
  emptyText: string;
  items: WorkTask[];
  now: number;
  onClear: (id: string) => void;
  onOpen: (task: WorkTask) => void;
  onSave: (id: string) => void;
  onSchedule: (id: string, value: string) => void;
  reminderValue: (task: WorkTask) => string;
  title: string;
}) {
  return (
    <section className="panel reminder-column">
      <div className="result-header">
        <h2>
          <CalendarClock size={18} />
          {title}
        </h2>
        <span className="subtle-count">{items.length}</span>
      </div>
      <div className="reminder-list">
        {items.map((item) => (
          <div className={`reminder-row ${reminderState(item, now)}`} key={item.id}>
            <button className="reminder-title" onClick={() => onOpen(item)} type="button">
              <strong>{item.title}</strong>
              <span>{projects[item.projectId].name} - {item.category} - {item.priority}</span>
              <small>{reminderLabel(item, now)}</small>
            </button>
            <label>
              Reminder
              <input value={reminderValue(item)} onChange={(event) => onSchedule(item.id, event.target.value)} type="datetime-local" />
            </label>
            <button className="primary-button" onClick={() => onSave(item.id)} type="button">
              <Save size={15} />
              Save
            </button>
            <button className="ghost-button" onClick={() => onClear(item.id)} type="button">
              Clear
            </button>
          </div>
        ))}
        {items.length === 0 && <p className="empty">{emptyText}</p>}
      </div>
    </section>
  );
}

function MobileStatusSection({
  activeWorkTaskId,
  items,
  now,
  onOpen,
  status,
}: {
  activeWorkTaskId: string;
  items: WorkTask[];
  now: number;
  onOpen: (task: WorkTask, target: "details" | "ai") => void;
  status: TaskStatus;
}) {
  const defaultOpen = status === "In Progress" || status === "Open" || status === "Blocked";

  return (
    <details className={`mobile-status-section status-${statusSlug(status)}`} open={defaultOpen}>
      <summary>
        <span>{status}</span>
        <strong>{items.length}</strong>
      </summary>
      <div className="mobile-task-list">
        {items.map((item) => (
          <article key={item.id} className={`${taskClassName(item, activeWorkTaskId)} mobile-task-card`}>
            <div>
              <strong>{item.title}</strong>
              <small>{item.category} - {item.priority}</small>
              <small>{taskDateLabel(item)}</small>
              {isValidDateTime(item.reminderAt) && <small>{reminderLabel(item, now)}</small>}
            </div>
            <div className="mobile-task-actions">
              <button className="ghost-button" onClick={() => onOpen(item, "details")} type="button">
                <ListTodo size={15} />
                Task details
              </button>
              <button className="primary-button" onClick={() => onOpen(item, "ai")} type="button">
                <Sparkles size={15} />
                AI prompt
              </button>
            </div>
          </article>
        ))}
        {items.length === 0 && <p className="empty compact-empty">No {status.toLowerCase()} tasks.</p>}
      </div>
    </details>
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
    checklist: task.checklist?.map((item) => ({ ...item, done: Boolean(item.done) })) ?? [],
    input: task.input ?? "",
    assets: task.assets ?? [],
    requirements: task.requirements ?? taskRequirements(fallbackProjectId, fallbackTemplateId),
    gptPrompt: task.gptPrompt ?? "",
    result: task.result ?? "",
    createdAt: task.createdAt || new Date().toISOString(),
    updatedAt: task.updatedAt || task.createdAt || new Date().toISOString(),
  };
}

function normalizeNote(note: AppNote): AppNote {
  const fallbackProjectId = note.projectId && projects[note.projectId] ? note.projectId : "avbob";
  const timestamp = note.updatedAt || note.createdAt || new Date().toISOString();
  const legacyContent = typeof note.content === "string" ? note.content.trim() : "";
  const entries = Array.isArray(note.entries)
    ? note.entries.map(normalizeNoteEntry)
    : legacyContent
      ? [{ id: createId(), content: legacyContent, createdAt: timestamp, updatedAt: timestamp }]
      : [];
  return {
    ...note,
    projectId: fallbackProjectId,
    title: note.title || "Untitled note",
    entries,
    pinned: Boolean(note.pinned),
    createdAt: note.createdAt || new Date().toISOString(),
    updatedAt: timestamp,
  };
}

function normalizeNoteEntry(entry: AppNoteEntry): AppNoteEntry {
  return {
    id: entry.id || createId(),
    content: entry.content ?? "",
    createdAt: entry.createdAt || new Date().toISOString(),
    updatedAt: entry.updatedAt || entry.createdAt || new Date().toISOString(),
  };
}

function reminderTriggerId(task: WorkTask) {
  return `${task.id}:${task.reminderAt}`;
}

async function notifyReminder(task: WorkTask) {
  if (typeof Notification === "undefined" || Notification.permission !== "granted") return;

  const title = `Reminder: ${task.title}`;
  const options: NotificationOptions = {
    body: `${projects[task.projectId].name} - ${reminderLabel(task, Date.now())}`,
    badge: "/pwa-icon.svg",
    icon: "/pwa-icon.svg",
    tag: `task-reminder-${task.id}`,
  };

  try {
    if ("serviceWorker" in navigator) {
      const registration = await navigator.serviceWorker.getRegistration();
      if (registration?.showNotification) {
        await registration.showNotification(title, options);
        return;
      }
    }
    new Notification(title, options);
  } catch {
    new Notification(title, options);
  }
}

function taskRequirements(projectId: ProjectId, templateId: string): Requirements {
  const template = projects[projectId].tasks.find((item) => item.id === templateId);
  return { ...(template?.requirements ?? defaultRequirements) };
}

function priorityRank(priority: Priority) {
  return { Low: 1, Normal: 2, High: 3, Urgent: 4 }[priority] ?? 0;
}

function statusRank(status: TaskStatus) {
  return { "In Progress": 1, Open: 2, Blocked: 3, "To Do Later": 4, Closed: 5 }[status] ?? 99;
}

function dateValue(date: string) {
  if (!date) return Number.MAX_SAFE_INTEGER;
  const value = new Date(date).getTime();
  return Number.isFinite(value) ? value : Number.MAX_SAFE_INTEGER;
}

function buildTaskSummary(tasks: WorkTask[]) {
  const now = Date.now();
  return {
    total: tasks.length,
    open: tasks.filter((item) => item.status === "Open").length,
    inProgress: tasks.filter((item) => item.status === "In Progress").length,
    blocked: tasks.filter((item) => item.status === "Blocked").length,
    toDoLater: tasks.filter((item) => item.status === "To Do Later").length,
    closed: tasks.filter((item) => item.status === "Closed").length,
    urgent: tasks.filter((item) => item.status !== "Closed" && item.priority === "Urgent").length,
    reminders: tasks.filter((item) => item.status !== "Closed" && isValidDateTime(item.reminderAt) && new Date(item.reminderAt).getTime() <= now).length,
  };
}

function buildReminderPlanner(tasks: WorkTask[], now: number) {
  const openTasks = tasks.filter((item) => item.status !== "Closed");
  const planned = openTasks
    .filter((item) => isValidDateTime(item.reminderAt))
    .sort((a, b) => dateValue(a.reminderAt) - dateValue(b.reminderAt));

  return {
    overdue: planned.filter((item) => dateValue(item.reminderAt) < startOfToday(now)),
    today: planned.filter((item) => {
      const value = dateValue(item.reminderAt);
      return value >= startOfToday(now) && value < startOfTomorrow(now);
    }),
    upcoming: planned.filter((item) => dateValue(item.reminderAt) >= startOfTomorrow(now)),
    unscheduled: openTasks
      .filter((item) => !isValidDateTime(item.reminderAt))
      .sort((a, b) => priorityRank(b.priority) - priorityRank(a.priority) || dateValue(a.dueDate) - dateValue(b.dueDate)),
  };
}

function buildProjectDashboard(tasks: WorkTask[]) {
  return (Object.keys(projects) as ProjectId[]).map((projectId) => {
    const projectTasks = tasks.filter((task) => task.projectId === projectId);
    const summary = buildTaskSummary(projectTasks);
    return { projectId, ...summary };
  });
}

function buildMobileProjectGroups(tasks: WorkTask[]) {
  return (Object.keys(projects) as ProjectId[]).map((projectId) => {
    const projectTasks = tasks
      .filter((task) => task.projectId === projectId)
      .sort((a, b) => statusRank(a.status) - statusRank(b.status) || priorityRank(b.priority) - priorityRank(a.priority) || dateValue(a.dueDate) - dateValue(b.dueDate));

    return {
      projectId,
      total: projectTasks.length,
      byStatus: Object.fromEntries(
        mobileStatusOrder.map((status) => [status, projectTasks.filter((task) => task.status === status)]),
      ) as Record<TaskStatus, WorkTask[]>,
    };
  });
}

function statusKey(status: TaskStatus): "open" | "inProgress" | "blocked" | "toDoLater" | "closed" {
  if (status === "In Progress") return "inProgress";
  if (status === "To Do Later") return "toDoLater";
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

function reminderLabel(task: WorkTask, now: number) {
  if (!isValidDateTime(task.reminderAt)) return "No reminder";
  const value = dateValue(task.reminderAt);
  const dateText = new Date(value).toLocaleString();
  if (value < startOfToday(now)) return `Overdue: ${dateText}`;
  if (value < startOfTomorrow(now)) return `Today: ${dateText}`;
  return `Upcoming: ${dateText}`;
}

function reminderState(task: WorkTask, now: number) {
  if (!isValidDateTime(task.reminderAt)) return "unscheduled";
  const value = dateValue(task.reminderAt);
  if (value < startOfToday(now)) return "overdue";
  if (value < startOfTomorrow(now)) return "today";
  return "upcoming";
}

function isOverdue(task: WorkTask) {
  if (!task.dueDate || task.status === "Closed") return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(task.dueDate).getTime() < today.getTime();
}

function isReminderDue(task: WorkTask) {
  return Boolean(isValidDateTime(task.reminderAt) && task.status !== "Closed" && new Date(task.reminderAt).getTime() <= Date.now());
}

function isValidDateTime(value: string) {
  return Boolean(value && Number.isFinite(new Date(value).getTime()));
}

function startOfToday(now: number) {
  const date = new Date(now);
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function startOfTomorrow(now: number) {
  const date = new Date(startOfToday(now));
  date.setDate(date.getDate() + 1);
  return date.getTime();
}

function nextPlanningHour(now: number) {
  const date = new Date(now);
  date.setMinutes(0, 0, 0);
  date.setHours(Math.min(date.getHours() + 1, 17));
  return date;
}

function localDatetimeValue(date: Date) {
  const offsetDate = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return offsetDate.toISOString().slice(0, 16);
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

async function readPptxText(file: File) {
  const { default: JSZip } = await import("jszip");
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slideFiles = Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort(comparePptxPartNames);
  const noteFiles = Object.keys(zip.files)
    .filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path))
    .sort(comparePptxPartNames);

  const slideSections = await Promise.all(
    slideFiles.map(async (path, index) => {
      const xml = await zip.files[path].async("text");
      const text = extractPresentationXmlText(xml);
      return text ? `Slide ${index + 1}\n${text}` : "";
    }),
  );
  const noteSections = await Promise.all(
    noteFiles.map(async (path, index) => {
      const xml = await zip.files[path].async("text");
      const text = extractPresentationXmlText(xml);
      return text ? `Speaker notes ${index + 1}\n${text}` : "";
    }),
  );
  const extracted = [...slideSections, ...noteSections].filter(Boolean).join("\n\n").trim();

  return extracted || "The PPTX file was read, but no slide text or speaker notes were found.";
}

function comparePptxPartNames(left: string, right: string) {
  return pptxPartNumber(left) - pptxPartNumber(right);
}

function pptxPartNumber(path: string) {
  return Number(path.match(/(\d+)\.xml$/)?.[1] ?? 0);
}

function extractPresentationXmlText(xml: string) {
  const matches = [...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)];
  return matches
    .map((match) => decodeXmlText(match[1]).trim())
    .filter(Boolean)
    .join("\n");
}

function decodeXmlText(value: string) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
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
  taskLabel: string,
  checklist: ChecklistItem[],
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
    `Task template: ${taskLabel}`,
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
    "TASK CHECKLIST OR SHOPPING LIST",
    checklist.length
      ? checklist.map((item) => `- [${item.done ? "x" : " "}] ${item.text}`).join("\n")
      : "No checklist items added.",
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
    "Create the requested output as a new, business-ready document, summary, or email draft according to the selected output type.",
    "If the requested output is an email, include a usable subject line and email body.",
    "If the requested output is a summary, distinguish confirmed information, assumptions, risks, gaps, and action items.",
    "If the task checklist contains open items, include them as action items or a tickable checklist when relevant.",
    "Do not summarize the prompt. Use all relevant source material. Ask clarifying questions only if required information is genuinely missing.",
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
