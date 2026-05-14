import {
  ArrowRight,
  BriefcaseBusiness,
  BarChart3,
  BellRing,
  CalendarClock,
  ChevronLeft,
  ChevronRight,
  Check,
  CheckSquare,
  Clipboard,
  Download,
  FileImage,
  FileText,
  History,
  Image,
  LayoutDashboard,
  ListTodo,
  Maximize2,
  Minimize2,
  Pin,
  Plus,
  Repeat,
  Save,
  ShieldCheck,
  Smartphone,
  Sparkles,
  Star,
  Settings,
  StickyNote,
  Trash2,
  Upload,
  WandSparkles,
  X,
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useState } from "react";
import { AlignmentType, Document, Footer, Header, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import {
  deleteTask,
  getCachedMeta,
  getCachedNotes,
  getCachedOutputs,
  getCachedTasks,
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
import type {
  AppNote,
  AppNoteEntry,
  ChecklistItem,
  Format,
  InputAsset,
  OutputTemplate,
  OutputTemplateFormat,
  Priority,
  ProjectId,
  Requirements,
  SavedOutput,
  TaskStatus,
  WorkTask,
} from "./types";
import { Home } from "./components/Home";
import { AIPromptBuilder } from "./components/AIPromptBuilder";

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
    sections: "Executive summary, Key facts, Key assumptions, Decisions or implications, Risks or gaps, Confidence level, Action items",
    constraints: "Separate confirmed facts from assumptions. Keep the wording easy to reuse in a document.",
    imageRequirements: "",
  },
  document: {
    outputType: "New business document",
    format: "Markdown",
    tone: "Professional and clear",
    audience: "Internal team",
    length: "Detailed",
    sections: "Title, Purpose, Background, Problem statement, Financial or operational impact, Main content, Recommendations, Next steps",
    constraints: "Produce a complete document, not notes about the document. Use headings and concise paragraphs.",
    imageRequirements: "",
  },
  email: {
    outputType: "Email draft",
    format: "Markdown",
    tone: "Warm, concise, and professional",
    audience: "Recipient named in the source notes",
    length: "Short",
    sections: "Subject line, Email body, Desired outcome, Optional follow-up note",
    constraints: "Write ready-to-send email copy. Include a clear ask or next step where appropriate.",
    imageRequirements: "",
  },
  report: {
    outputType: "Structured report",
    format: "Markdown",
    tone: "Evidence-led and businesslike",
    audience: "Decision makers",
    length: "Detailed",
    sections: "Executive summary, Context, Methodology, Findings, Evidence, Risks, Limitations, Recommendations, Next steps",
    constraints: "Make the report useful for decisions. Flag missing evidence clearly.",
    imageRequirements: "",
  },
  meetingBrief: {
    outputType: "Meeting or forum briefing pack",
    format: "Markdown",
    tone: "Concise, structured, and decision-focused",
    audience: "Meeting attendees and decision makers",
    length: "Medium",
    sections: "Purpose, Desired outcomes, Pre-read required, Agenda, Current status, Key discussion points, Risks, Decisions needed, Actions",
    constraints: "Use the source notes to prepare a practical meeting pack. Separate discussion points from decisions needed.",
    imageRequirements: "",
  },
  marketResearch: {
    outputType: "Market research and opportunity note",
    format: "Markdown",
    tone: "Analytical, balanced, and commercially useful",
    audience: "Business development and leadership team",
    length: "Detailed",
    sections: "Executive summary, Company profile, Market context, Competitive positioning, Opportunity fit, Commercial relevance to us, Risks and unknowns, Recommended next steps",
    constraints: "Distinguish source-provided facts from assumptions. Include gaps that need verification.",
    imageRequirements: "",
  },
  proposal: {
    outputType: "Proposal-ready document",
    format: "Markdown",
    tone: "Confident, useful, and client-ready",
    audience: "Client or sponsor",
    length: "Detailed",
    sections: "Overview, Client need, Proposed approach, Deliverables, Value / ROI case, Success metrics, Timeline, Assumptions, Next steps",
    constraints: "Avoid generic sales language. Convert rough notes into concrete proposed work.",
    imageRequirements: "",
  },
  presentation: {
    outputType: "Presentation content",
    format: "Markdown",
    tone: "Clear, structured, and presentation-ready",
    audience: "Meeting audience",
    length: "Medium",
    sections: "Slide title, Key message, Slide bullets, Speaker notes, Visual direction, Closing action",
    constraints: "Keep slide bullets short. Put detail in speaker notes.",
    imageRequirements: "Suggest useful visuals, charts, or screenshots where they would improve the presentation.",
  },
  process: {
    outputType: "Process document",
    format: "Markdown",
    tone: "Plain, precise, and operational",
    audience: "People following the process",
    length: "Detailed",
    sections: "Purpose, Scope, Roles, SLA / response expectations, Steps, Exceptions, Checklist, Owner and review date",
    constraints: "Write actionable steps in order. Include a checklist that can be ticked off.",
    imageRequirements: "",
  },
  clientUpdate: {
    outputType: "Client or stakeholder update",
    format: "Markdown",
    tone: "Clear, calm, and accountable",
    audience: "Client, sponsor, or internal stakeholder",
    length: "Short",
    sections: "Subject line, Overall status indicator (Green / Amber / Red), Status summary, Progress, Risks or blockers, Impact if risks materialise, Decisions needed, Next steps",
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
  decisionBrief: {
    outputType: "Decision brief",
    format: "Markdown",
    tone: "Concise, executive, and decision-focused",
    audience: "Decision makers",
    length: "Medium",
    sections: "Context, Problem, Options, Recommendation, Risks, Decision required",
    constraints: "Frame the decision clearly and make the recommendation easy to accept, reject, or revise.",
    imageRequirements: "",
  },
  strategyVision: {
    outputType: "Strategy or vision document",
    format: "Markdown",
    tone: "Strategic, clear, and commercially grounded",
    audience: "Leadership, partners, or project team",
    length: "Detailed",
    sections: "Vision, Problem space, Strategic pillars, Differentiation, Execution roadmap",
    constraints: "Keep the strategy actionable. Connect vision to practical execution steps.",
    imageRequirements: "",
  },
  productSpec: {
    outputType: "Product or feature specification",
    format: "Markdown",
    tone: "Precise, practical, and implementation-ready",
    audience: "Product, design, and engineering team",
    length: "Detailed",
    sections: "Objective, User roles, Functional requirements, User flows, Data requirements, Success criteria",
    constraints: "Write requirements clearly enough for implementation planning. Flag unresolved product decisions.",
    imageRequirements: "",
  },
  promptGenerator: {
    outputType: "AI prompt generator",
    format: "Markdown",
    tone: "Direct, structured, and reusable",
    audience: "AI assistant or prompt user",
    length: "Medium",
    sections: "Task objective, Input description, Output requirements, Formatting instructions, Constraints",
    constraints: "Produce a ready-to-use prompt that another AI can follow without extra context.",
    imageRequirements: "",
  },
  riskIssueLog: {
    outputType: "Risk or issue log",
    format: "Markdown",
    tone: "Clear, accountable, and operational",
    audience: "Project team and stakeholders",
    length: "Medium",
    sections: "Issue, Impact, Likelihood, Owner, Mitigation, Status",
    constraints: "Make each risk or issue trackable with clear ownership and next action.",
    imageRequirements: "",
  },
  adHoc: {
    outputType: "Ad hoc structured output",
    format: "Markdown",
    tone: "Professional, clear, and practical",
    audience: "Internal team",
    length: "Medium",
    sections: "Title, Context, Main content, Decisions, Next steps",
    constraints: "Use the section list exactly as the working structure. Keep headings clear and omit sections only when they are genuinely not relevant.",
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
  { id: "decision-brief", label: "Decision brief", description: "Prepare executive decision support with options, recommendation, risks, and decision needed.", category: "Strategy", requirements: requirementPresets.decisionBrief },
  { id: "strategy-vision", label: "Strategy/vision document", description: "Shape vision, strategic pillars, differentiation, and a practical execution roadmap.", category: "Strategy", requirements: requirementPresets.strategyVision },
  { id: "product-feature-spec", label: "Product/feature spec", description: "Define objectives, users, requirements, flows, data needs, and success criteria for build work.", category: "Product", requirements: requirementPresets.productSpec },
  { id: "ai-prompt-generator", label: "AI prompt generator", description: "Create a reusable prompt with objective, inputs, output rules, formatting, and constraints.", category: "AI", requirements: requirementPresets.promptGenerator },
  { id: "risk-issue-log", label: "Risk/issue log", description: "Track issues, impact, likelihood, ownership, mitigation, and status.", category: "Planning", requirements: requirementPresets.riskIssueLog },
  { id: "ad-hoc-template", label: "Ad hoc template", description: "Define the content goal, audience, and section list from the frontend for one-off outputs.", category: "Custom", requirements: requirementPresets.adHoc },
];

const legacyTemplateMap: Record<string, string> = {
  "client-communication": "draft-email",
  "summarize-documents": "summarize",
  "advisory-forum": "meeting-brief",
  "support-process": "process-document",
  "banking-documents": "draft-document",
  "process-notes": "process-document",
  "banking-client-update": "client-update",
  "bd-opportunity-note": "market-research",
  "customer-success-update": "client-update",
};

const projects: Record<ProjectId, { name: string; context: string; tasks: TaskTemplate[] }> = {
  avbob: {
    name: "AVBOB",
    context: "Client communication, document preparation, reports, presentations, and polished business content.",
    tasks: commonTasks,
  },
  naha: {
    name: "Naha Banking",
    context: "Banking-related drafts, client summaries, process notes, reports, proposals, and product copy.",
    tasks: commonTasks,
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

type ProjectSettings = Record<ProjectId, { name: string; context: string }>;

const defaultProjectSettings: ProjectSettings = Object.fromEntries(
  (Object.keys(projects) as ProjectId[]).map((id) => [id, { name: projects[id].name, context: projects[id].context }]),
) as ProjectSettings;
const defaultTaskTemplates = commonTasks;
const defaultMasterStatuses = {
  task: ["Open", "In Progress", "Blocked", "To Do Later", "Closed"] as TaskStatus[],
  note: ["Active", "Closed"] as Array<"Active" | "Closed">,
};
const taskStatuses: TaskStatus[] = defaultMasterStatuses.task;

const seedTimestamp = "2026-04-30T00:00:00.000Z";
const defaultBrand = {
  brandName: "TaskOS AI",
  primaryColor: "1f3a5f",
  secondaryColor: "eef3ff",
  accentColor: "3157ff",
  logoDataUrl: "",
  sourceTemplateName: "",
  sourceTemplateType: "",
  sourceTemplateDataUrl: "",
  sourceTemplateText: "",
};
const defaultOutputTemplates: OutputTemplate[] = [
  {
    id: "business-document",
    name: "Business Document",
    description: "General business document with clear sections and reusable prose.",
    group: "Document",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["draft-document", "summarize", "create-report", "market-research", "strategy-vision"],
    slots: ["Title", "Executive summary", "Background", "Main content", "Recommendations", "Next steps"],
    style: "Use H1 for the title, H2 for major sections, concise paragraphs, and action-oriented recommendations.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "ad-hoc-output",
    name: "Ad hoc Output",
    description: "Flexible output profile that follows the sections defined in the task brief.",
    group: "Custom",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["ad-hoc-template"],
    slots: ["Use the task-defined sections"],
    style: "Follow the sections supplied in the content brief. Use clean headings and practical wording.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "checklist-output",
    name: "Checklist",
    description: "Tickable checklist output for shopping lists, subtasks, follow-ups, and simple action tracking.",
    group: "Checklist",
    format: "Markdown",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["checklist"],
    slots: ["Checklist", "Notes", "Next steps"],
    style: "Return a concise tickable Markdown checklist using - [ ] items. Keep each item specific, short, and easy to complete.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "meeting-brief-output",
    name: "Meeting Brief",
    description: "Decision-focused meeting pack with agenda, pre-read, risks, and actions.",
    group: "Brief",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["meeting-brief"],
    slots: ["Purpose", "Desired outcomes", "Pre-read required", "Agenda", "Current status", "Decisions needed", "Actions"],
    style: "Make the brief practical for a live meeting. Separate discussion points, decisions, and follow-up actions.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "executive-decision-brief",
    name: "Executive Decision Brief",
    description: "Short decision-support format for leadership review.",
    group: "Brief",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["decision-brief"],
    slots: ["Context", "Problem", "Options", "Recommendation", "Risks", "Decision required"],
    style: "Keep it executive-ready, direct, and easy to scan. Highlight the recommended option clearly.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "process-document-output",
    name: "Process Document",
    description: "Operational process output with roles, steps, exceptions, and review points.",
    group: "Document",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["process-document"],
    slots: ["Purpose", "Scope", "Roles", "SLA", "Steps", "Exceptions", "Checklist", "Owner and review date"],
    style: "Write clear numbered process steps. Include exceptions and a short checklist for repeat execution.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "client-update",
    name: "Client Update",
    description: "Status update format for clients, sponsors, or internal stakeholders.",
    group: "Email",
    format: "Markdown",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["client-update", "draft-email"],
    slots: ["Subject line", "Overall status", "Progress", "Risks", "Decisions needed", "Next steps"],
    style: "Write in send-ready language. Use calm, accountable wording and make decisions or asks explicit.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "product-spec-output",
    name: "Product / Feature Spec",
    description: "Implementation-ready product specification for product, design, and engineering work.",
    group: "Specification",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["product-feature-spec"],
    slots: ["Objective", "User roles", "Functional requirements", "User flows", "Data requirements", "Success criteria", "Open questions"],
    style: "Write requirements clearly enough for implementation planning. Flag unresolved product decisions explicitly.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "ai-prompt-output",
    name: "Reusable AI Prompt",
    description: "Clean prompt block output with objective, inputs, constraints, and formatting rules.",
    group: "AI",
    format: "Markdown",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["ai-prompt-generator"],
    slots: ["Task objective", "Input description", "Output requirements", "Formatting instructions", "Constraints"],
    style: "Return only the reusable prompt text. Avoid surrounding commentary so it can be copied directly.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "proposal",
    name: "Proposal",
    description: "Proposal-ready structure with value, deliverables, timeline, and metrics.",
    group: "Proposal",
    format: "DOCX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["proposal-copy"],
    slots: ["Overview", "Client need", "Proposed approach", "Deliverables", "Value / ROI case", "Success metrics", "Timeline"],
    style: "Use confident, specific language. Avoid filler sales copy and connect deliverables to measurable value.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "presentation-deck",
    name: "Presentation Deck",
    description: "Slide-by-slide structure for AI-assisted presentation creation.",
    group: "Presentation",
    format: "PPTX",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["presentation-text"],
    slots: ["Title slide", "Agenda", "Context", "Key message slides", "Recommendation", "Closing action"],
    style: "Return strict slide JSON only. Each slide object must include title, layout, background, keyMessage, bullets, speakerNotes, visualDirection, and images.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "risk-issue-log",
    name: "Risk / Issue Log",
    description: "Trackable risk and issue format with ownership and mitigation.",
    group: "Checklist",
    format: "Markdown",
    ...defaultBrand,
    scope: "global",
    compatibleTaskIds: ["risk-issue-log"],
    slots: ["Issue", "Impact", "Likelihood", "Owner", "Mitigation", "Status"],
    style: "Use a table when there are multiple risks or issues. Make ownership and next action visible.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "rainfin-word-doc",
    name: "Rainfin Word Doc",
    description: "Rainfin-branded Word document export for formal reports, proposals, and briefs.",
    group: "Document",
    format: "DOCX",
    ...defaultBrand,
    brandName: "Rainfin",
    primaryColor: "17324d",
    secondaryColor: "eef4f8",
    accentColor: "2c80b8",
    scope: "global",
    compatibleTaskIds: ["draft-document", "create-report", "proposal-copy", "decision-brief", "market-research", "meeting-brief"],
    slots: ["Title", "Executive summary", "Background", "Main content", "Recommendations", "Next steps"],
    style: "Use a formal consulting-document style with clear headings, concise paragraphs, and executive-ready recommendations.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
  {
    id: "supplysync360-powerpoint",
    name: "SupplySync360 PowerPoint",
    description: "SupplySync360-branded presentation export for slide-ready AI outputs.",
    group: "Presentation",
    format: "PPTX",
    ...defaultBrand,
    brandName: "SupplySync360",
    primaryColor: "0f3f3c",
    secondaryColor: "e9f6f3",
    accentColor: "2fbf9b",
    scope: "global",
    compatibleTaskIds: ["presentation-text", "strategy-vision", "proposal-copy"],
    slots: ["Title slide", "Agenda", "Context", "Key message slides", "Recommendation", "Closing action"],
    style: "Return strict slide JSON only. Use concise slide language, strong key messages, and practical visual direction for each slide.",
    createdAt: seedTimestamp,
    updatedAt: seedTimestamp,
  },
];

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
const outputTemplateMetaKey = "outputTemplates";
const projectSettingsMetaKey = "projectSettings";
const taskTemplatesMetaKey = "taskTemplates";
const masterStatusesMetaKey = "masterStatuses";
const reminderDeviceIdStorageKey = "ai-workbench-reminder-device-id";
const deviceTriggeredReminderMetaKey = `triggeredReminderIds:${getReminderDeviceId()}`;
const mobileStatusOrder: TaskStatus[] = ["In Progress", "Open", "Blocked", "To Do Later", "Closed"];
const maxUploadSizeBytes = 8 * 1024 * 1024;
const maxTextAssetCharacters = 120_000;

function createId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function getReminderDeviceId() {
  if (typeof localStorage === "undefined") return "browser";

  const existing = localStorage.getItem(reminderDeviceIdStorageKey);
  if (existing) return existing;

  const next = typeof crypto !== "undefined" && "randomUUID" in crypto ? crypto.randomUUID() : createId();
  localStorage.setItem(reminderDeviceIdStorageKey, next);
  return next;
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
  const [outputTemplates, setOutputTemplates] = useState<OutputTemplate[]>(defaultOutputTemplates);
  const [projectSettings, setProjectSettings] = useState<ProjectSettings>(defaultProjectSettings);
  const [taskTemplates, setTaskTemplates] = useState<TaskTemplate[]>(defaultTaskTemplates);
  const [masterStatuses, setMasterStatuses] = useState(defaultMasterStatuses);
  const [selectedOutputTemplateId, setSelectedOutputTemplateId] = useState(defaultOutputTemplates[0].id);
  const [workTasks, setWorkTasks] = useState<WorkTask[]>([]);
  const [notes, setNotes] = useState<AppNote[]>([]);
  const [activeWorkTaskId, setActiveWorkTaskId] = useState("");
  const [viewMode, setViewMode] = useState<"home" | "work" | "favorites" | "ai" | "ai-engine" | "reminders" | "notes" | "settings" | "mobile">("home");
  const [quickTaskTitle, setQuickTaskTitle] = useState("");
  const [quickTaskDetails, setQuickTaskDetails] = useState("");
  const [quickChecklistDraft, setQuickChecklistDraft] = useState("");
  const [quickChecklistItems, setQuickChecklistItems] = useState<ChecklistItem[]>([]);
  const [quickChecklistItem, setQuickChecklistItem] = useState("");
  const [noteDraft, setNoteDraft] = useState({ title: "", content: "" });
  const [workQuickTaskDraft, setWorkQuickTaskDraft] = useState({ title: "", details: "" });
  const [workQuickNoteDraft, setWorkQuickNoteDraft] = useState({ title: "", content: "" });
  const [captureTaskOpen, setCaptureTaskOpen] = useState(false);
  const [captureNoteOpen, setCaptureNoteOpen] = useState(false);
  const [noteStatusFilter, setNoteStatusFilter] = useState<"Active" | "Closed" | "All">("Active");
  const [taskBrowserOpen, setTaskBrowserOpen] = useState(true);
  const [mobileFullOpen, setMobileFullOpen] = useState(false);
  const [mobileSection, setMobileSection] = useState<"capture" | "tasks" | "reminders">("tasks");
  const [taskListFilter, setTaskListFilter] = useState<"Active" | "Favorites" | TaskStatus | "All">("Active");
  const [previewSlideIndex, setPreviewSlideIndex] = useState(0);
  const [fullScreenPreviewOpen, setFullScreenPreviewOpen] = useState(false);
  const [reminderDrafts, setReminderDrafts] = useState<Record<string, string>>({});
  const [triggeredReminderIds, setTriggeredReminderIds] = useState<string[]>([]);
  const [notificationPermission, setNotificationPermission] = useState<NotificationPermission | "unsupported">(() =>
    typeof Notification === "undefined" ? "unsupported" : Notification.permission,
  );
  const [now, setNow] = useState(() => Date.now());
  const [syncing, setSyncing] = useState(false);
  const [message, setMessage] = useState("");

  const masterProjects = projectSettings;
  const masterTaskStatuses = masterStatuses.task.length ? masterStatuses.task : defaultMasterStatuses.task;
  const masterNoteStatuses = masterStatuses.note.length ? masterStatuses.note : defaultMasterStatuses.note;
  const projectIds = Object.keys(masterProjects) as ProjectId[];
  const activeProjectId = masterProjects[projectId] ? projectId : projectIds[0] ?? "avbob";
  const project = {
    ...(masterProjects[activeProjectId] ?? defaultProjectSettings.avbob),
    tasks: taskTemplates.length ? taskTemplates : defaultTaskTemplates,
  };
  const task = project.tasks.find((item) => item.id === taskId) ?? project.tasks[0];
  const availableOutputTemplates = outputTemplates.filter((item) => isOutputTemplateAvailable(item, projectId));
  const linkedOutputTemplates = compatibleOutputTemplates(availableOutputTemplates, taskId);
  const selectedOutputTemplate =
    linkedOutputTemplates.find((item) => item.id === selectedOutputTemplateId) ??
    preferredOutputTemplateForTask(projectId, taskId, availableOutputTemplates) ??
    linkedOutputTemplates[0] ??
    defaultOutputTemplates[0];
  const projectHistory = savedOutputs.filter((item) => item.projectId === projectId);
  const activeWorkTask = workTasks.find((item) => item.id === activeWorkTaskId);
  const activeTaskHistory = activeWorkTask ? savedOutputs.filter((item) => item.workTaskId === activeWorkTask.id) : [];
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
  const favoriteTasks = sortedWorkTasks.filter((item) => item.favorite).slice(0, 8);
  const favoriteNotes = [...notes]
    .filter((note) => note.favorite)
    .sort((a, b) => dateValue(b.updatedAt) - dateValue(a.updatedAt))
    .slice(0, 6);
  const visibleNotes = [...notes]
    .filter((note) => noteStatusFilter === "All" || (note.status ?? "Active") === noteStatusFilter)
    .sort((a, b) => Number(b.pinned) - Number(a.pinned) || dateValue(b.updatedAt) - dateValue(a.updatedAt));
  const visibleProjectTasks = filterTasksByListMode(sortedProjectTasks, taskListFilter);
  const recurringTasks = sortedProjectTasks.filter((item) => (item.recurrence ?? "None") !== "None");
  const mobileActiveProjectTasks = sortedProjectTasks.filter((item) => item.status !== "Closed");
  const mobileFavoriteProjectTasks = sortedProjectTasks.filter((item) => item.favorite);
  const recentNotes = [...notes].sort((a, b) => dateValue(b.updatedAt) - dateValue(a.updatedAt)).slice(0, 4);

  const missingDetails = useMemo(() => {
    const missing: string[] = [];
    if (input.trim().length < 20 && assets.length === 0) missing.push("Add source input, notes, or a readable file.");
    if (!requirements.outputType.trim()) missing.push("Choose the output type you need.");
    if (!requirements.audience.trim()) missing.push("Describe who the output is for.");
    if (!requirements.sections.trim()) missing.push("List the required sections or structure.");
    return missing;
  }, [assets.length, input, requirements.audience, requirements.outputType, requirements.sections]);
  const inputQuality = useMemo(() => buildInputQuality(missingDetails, assets), [assets, missingDetails]);
  const isPresentationOutput = selectedOutputTemplate.format === "PPTX";
  const slideDeck = useMemo(() => buildSlideDeck(result, assets, selectedOutputTemplate), [assets, result, selectedOutputTemplate]);
  const slidePreview = slideDeck.slides;
  const activePreviewSlideIndex = Math.min(previewSlideIndex, Math.max(slidePreview.length - 1, 0));
  const activePreviewSlide = slidePreview[activePreviewSlideIndex];

  useEffect(() => {
    if (linkedOutputTemplates.some((item) => item.id === selectedOutputTemplateId)) return;
    setSelectedOutputTemplateId(selectedOutputTemplate.id);
  }, [linkedOutputTemplates, selectedOutputTemplate.id, selectedOutputTemplateId]);

  useEffect(() => {
    if (previewSlideIndex < slidePreview.length) return;
    setPreviewSlideIndex(Math.max(slidePreview.length - 1, 0));
  }, [previewSlideIndex, slidePreview.length]);

  useEffect(() => {
    if (!fullScreenPreviewOpen) return;
    function handleKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") setFullScreenPreviewOpen(false);
      if (event.key === "ArrowLeft") setPreviewSlideIndex((current) => Math.max(current - 1, 0));
      if (event.key === "ArrowRight") setPreviewSlideIndex((current) => Math.min(current + 1, Math.max(slidePreview.length - 1, 0)));
    }
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [fullScreenPreviewOpen, slidePreview.length]);

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
      const triggeredResult = await getMeta<string[]>(deviceTriggeredReminderMetaKey, []);
      const outputTemplateResult = await getMeta<OutputTemplate[]>(outputTemplateMetaKey, defaultOutputTemplates);
      const projectSettingsResult = await getMeta<ProjectSettings>(projectSettingsMetaKey, defaultProjectSettings);
      const taskTemplatesResult = await getMeta<TaskTemplate[]>(taskTemplatesMetaKey, defaultTaskTemplates);
      const masterStatusesResult = await getMeta<typeof defaultMasterStatuses>(masterStatusesMetaKey, defaultMasterStatuses);

      setWorkTasks(taskResult.map(normalizeWorkTask));
      setNotes(noteResult.map(normalizeNote));
      setSavedOutputs(outputResult);
      setOutputTemplates(normalizeOutputTemplates(outputTemplateResult));
      setProjectSettings(normalizeProjectSettings(projectSettingsResult));
      setTaskTemplates(normalizeTaskTemplates(taskTemplatesResult));
      setMasterStatuses(normalizeMasterStatuses(masterStatusesResult));
      setTriggeredReminderIds(triggeredResult);
      setNow(Date.now());
      setMessage(`Synced ${taskResult.length} tasks, ${noteResult.length} notes, and ${outputResult.length} saved outputs from ${getStorageBackendLabel()}.${outputWarning}`);
    } catch (error) {
      setMessage(error instanceof Error ? `Database sync failed: ${error.message}` : "Database sync failed.");
    } finally {
      setSyncing(false);
    }
  }

  async function loadCachedData() {
    try {
      const [
        taskResult,
        outputResult,
        noteResult,
        triggeredResult,
        outputTemplateResult,
        projectSettingsResult,
        taskTemplatesResult,
        masterStatusesResult,
      ] = await Promise.all([
        getCachedTasks(),
        getCachedOutputs(),
        getCachedNotes(),
        getCachedMeta<string[]>(deviceTriggeredReminderMetaKey, []),
        getCachedMeta<OutputTemplate[]>(outputTemplateMetaKey, defaultOutputTemplates),
        getCachedMeta<ProjectSettings>(projectSettingsMetaKey, defaultProjectSettings),
        getCachedMeta<TaskTemplate[]>(taskTemplatesMetaKey, defaultTaskTemplates),
        getCachedMeta<typeof defaultMasterStatuses>(masterStatusesMetaKey, defaultMasterStatuses),
      ]);

      setWorkTasks(taskResult.map(normalizeWorkTask));
      setNotes(noteResult.map(normalizeNote));
      setSavedOutputs(outputResult);
      setOutputTemplates(normalizeOutputTemplates(outputTemplateResult));
      setProjectSettings(normalizeProjectSettings(projectSettingsResult));
      setTaskTemplates(normalizeTaskTemplates(taskTemplatesResult));
      setMasterStatuses(normalizeMasterStatuses(masterStatusesResult));
      setTriggeredReminderIds(triggeredResult);
      setNow(Date.now());
      setMessage(
        taskResult.length || noteResult.length || outputResult.length
          ? `Loaded local cache: ${taskResult.length} tasks, ${noteResult.length} notes, and ${outputResult.length} saved outputs. Use Sync to refresh.`
          : "No local cache found yet. Use Sync to pull your saved data.",
      );
    } catch (error) {
      setMessage(error instanceof Error ? `Local cache failed: ${error.message}` : "Local cache failed.");
    }
  }

  useEffect(() => {
    void (async () => {
      await loadCachedData();
      await refreshData();
    })();
  }, []);

  useEffect(() => {
    const timer = window.setInterval(() => {
      setNow(Date.now());
    }, 30000);

    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    function checkRemindersNow() {
      setNow(Date.now());
    }

    function handleVisibilityChange() {
      if (document.visibilityState === "visible") checkRemindersNow();
    }

    window.addEventListener("focus", checkRemindersNow);
    window.addEventListener("pageshow", checkRemindersNow);
    document.addEventListener("visibilitychange", handleVisibilityChange);

    return () => {
      window.removeEventListener("focus", checkRemindersNow);
      window.removeEventListener("pageshow", checkRemindersNow);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, []);

  useEffect(() => {
    const nextReminderAt = workTasks.reduce<number | undefined>((earliest, item) => {
      if (!item.reminderAt || item.status === "Closed") return earliest;
      const dueAt = new Date(item.reminderAt).getTime();
      if (!Number.isFinite(dueAt) || dueAt <= now || triggeredReminderIds.includes(reminderTriggerId(item))) return earliest;
      return earliest === undefined || dueAt < earliest ? dueAt : earliest;
    }, undefined);

    if (nextReminderAt === undefined) return;

    const delay = Math.max(0, Math.min(nextReminderAt - Date.now() + 1000, 2_147_483_647));
    const timer = window.setTimeout(() => {
      setNow(Date.now());
    }, delay);

    return () => window.clearTimeout(timer);
  }, [now, triggeredReminderIds, workTasks]);

  useEffect(() => {
    const dueTasks = workTasks.filter((item) => {
      if (!item.reminderAt || item.status === "Closed") return false;
      const dueAt = new Date(item.reminderAt).getTime();
      return Number.isFinite(dueAt) && dueAt <= now && !triggeredReminderIds.includes(reminderTriggerId(item));
    });

    if (dueTasks.length === 0) return;

    const nextTriggeredIds = [...triggeredReminderIds, ...dueTasks.map(reminderTriggerId)];
    setTriggeredReminderIds(nextTriggeredIds);
    void setMeta(deviceTriggeredReminderMetaKey, nextTriggeredIds);
    dueTasks.forEach((item) => {
      void notifyReminder(item);
    });
  }, [now, triggeredReminderIds, workTasks]);

  function updateProject(nextProjectId: ProjectId) {
    setProjectId(nextProjectId);
    const nextTaskId = taskTemplates[0]?.id ?? defaultTaskTemplates[0].id;
    const nextOutputTemplate = preferredOutputTemplateForTask(nextProjectId, nextTaskId, outputTemplates) ?? defaultOutputTemplates[0];
    setTaskId(nextTaskId);
    setSelectedOutputTemplateId(nextOutputTemplate.id);
    setRequirements(requirementsLinkedToOutputTemplate(taskRequirements(nextProjectId, nextTaskId, taskTemplates), nextOutputTemplate));
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
    const nextOutputTemplate = preferredOutputTemplateForTask(projectId, nextTaskId, availableOutputTemplates) ?? selectedOutputTemplate;
    const nextRequirements = requirementsLinkedToOutputTemplate(taskRequirements(projectId, nextTaskId, taskTemplates), nextOutputTemplate);
    setSelectedOutputTemplateId(nextOutputTemplate.id);
    setRequirements(nextRequirements);
    if (activeWorkTask) {
      updateWorkTask(activeWorkTask.id, "templateId", nextTaskId);
      updateWorkTask(activeWorkTask.id, "category", nextTemplate.category);
      updateWorkTask(activeWorkTask.id, "outputTemplateId", nextOutputTemplate.id);
      updateWorkTask(activeWorkTask.id, "requirements", nextRequirements);
    }
  }

  function selectOutputTemplate(nextOutputTemplateId: string) {
    const nextTemplate = outputTemplates.find((item) => item.id === nextOutputTemplateId) ?? defaultOutputTemplates[0];
    if (!isOutputTemplateCompatibleWithTask(nextTemplate, taskId)) {
      setMessage(`${nextTemplate.name} is not linked to ${task.label}. Choose a linked output template or edit the template compatibility.`);
      return;
    }
    const nextRequirements = requirementsLinkedToOutputTemplate(requirements, nextTemplate);
    setSelectedOutputTemplateId(nextOutputTemplateId);
    setRequirements(nextRequirements);
    if (activeWorkTask) {
      updateWorkTask(activeWorkTask.id, "outputTemplateId", nextOutputTemplateId);
      updateWorkTask(activeWorkTask.id, "requirements", nextRequirements);
    }
  }

  function saveOutputTemplates(nextTemplates: OutputTemplate[]) {
    const normalized = normalizeOutputTemplates(nextTemplates);
    setOutputTemplates(normalized);
    void setMeta(outputTemplateMetaKey, normalized);
  }

  function saveProjectSettings(nextSettings: ProjectSettings) {
    const normalized = normalizeProjectSettings(nextSettings);
    setProjectSettings(normalized);
    void setMeta(projectSettingsMetaKey, normalized);
  }

  function updateProjectSetting<K extends keyof ProjectSettings[ProjectId]>(id: ProjectId, key: K, value: ProjectSettings[ProjectId][K]) {
    saveProjectSettings({
      ...projectSettings,
      [id]: {
        ...projectSettings[id],
        [key]: value,
      },
    });
  }

  function saveTaskTemplates(nextTemplates: TaskTemplate[]) {
    const normalized = normalizeTaskTemplates(nextTemplates);
    setTaskTemplates(normalized);
    void setMeta(taskTemplatesMetaKey, normalized);
  }

  function updateTaskTemplate(templateId: string, patch: Partial<TaskTemplate>) {
    saveTaskTemplates(taskTemplates.map((template) => (template.id === templateId ? { ...template, ...patch } : template)));
  }

  function addProjectSetting() {
    const id = `project-${Date.now().toString(36)}`;
    saveProjectSettings({
      ...projectSettings,
      [id]: {
        name: "New Project",
        context: "",
      },
    });
    setProjectId(id);
  }

  function deleteProjectSetting(id: ProjectId) {
    const ids = Object.keys(projectSettings);
    if (ids.length <= 1) return;
    const { [id]: _deleted, ...next } = projectSettings;
    saveProjectSettings(next as ProjectSettings);
    if (projectId === id) setProjectId(Object.keys(next)[0] ?? "avbob");
  }

  function addTaskTemplate() {
    const template: TaskTemplate = {
      id: `task-${Date.now().toString(36)}`,
      label: "New task template",
      description: "",
      category: "Custom",
      requirements: requirementPresets.adHoc,
    };
    saveTaskTemplates([...taskTemplates, template]);
  }

  function deleteTaskTemplate(templateId: string) {
    const nextTemplates = taskTemplates.filter((template) => template.id !== templateId);
    if (nextTemplates.length) saveTaskTemplates(nextTemplates);
  }

  function addTaskStatus() {
    saveMasterStatuses({ ...masterStatuses, task: [...masterTaskStatuses, "Open"] });
  }

  function deleteTaskStatus(status: TaskStatus) {
    const nextStatuses = masterTaskStatuses.filter((item) => item !== status);
    if (nextStatuses.length) saveMasterStatuses({ ...masterStatuses, task: nextStatuses });
  }

  function addNoteStatus() {
    saveMasterStatuses({ ...masterStatuses, note: [...masterNoteStatuses, "Active"] });
  }

  function deleteNoteStatus(status: "Active" | "Closed") {
    const nextStatuses = masterNoteStatuses.filter((item) => item !== status);
    if (nextStatuses.length) saveMasterStatuses({ ...masterStatuses, note: nextStatuses });
  }

  function createOutputTemplate() {
    const timestamp = new Date().toISOString();
    const template: OutputTemplate = {
      ...defaultOutputTemplates[0],
      id: `output-${Date.now().toString(36)}`,
      name: "New Output Template",
      description: "",
      createdAt: timestamp,
      updatedAt: timestamp,
    };
    saveOutputTemplates([template, ...outputTemplates]);
    setSelectedOutputTemplateId(template.id);
  }

  function saveMasterStatuses(nextStatuses: typeof defaultMasterStatuses) {
    const normalized = normalizeMasterStatuses(nextStatuses);
    setMasterStatuses(normalized);
    void setMeta(masterStatusesMetaKey, normalized);
  }

  function updateSelectedOutputTemplate<K extends keyof OutputTemplate>(key: K, value: OutputTemplate[K]) {
    const timestamp = new Date().toISOString();
    const nextTemplate = {
      ...selectedOutputTemplate,
      [key]: value,
      updatedAt: timestamp,
    };
    saveOutputTemplates(
      outputTemplates.map((item) => (item.id === selectedOutputTemplate.id ? nextTemplate : item)),
    );
    if (key === "format" || key === "slots") {
      const nextRequirements = requirementsLinkedToOutputTemplate(requirements, nextTemplate);
      setRequirements(nextRequirements);
      if (activeWorkTask) updateWorkTask(activeWorkTask.id, "requirements", nextRequirements);
    }
  }

  function duplicateOutputTemplate() {
    const timestamp = new Date().toISOString();
    const copy: OutputTemplate = {
      ...selectedOutputTemplate,
      id: createId(),
      name: `${selectedOutputTemplate.name} copy`,
      createdAt: timestamp,
      updatedAt: timestamp,
    };
    saveOutputTemplates([copy, ...outputTemplates]);
    setSelectedOutputTemplateId(copy.id);
  }

  function deleteOutputTemplate(id: string) {
    if (outputTemplates.length <= 1) return;
    const next = outputTemplates.filter((item) => item.id !== id);
    saveOutputTemplates(next);
    setSelectedOutputTemplateId((next[0] ?? defaultOutputTemplates[0]).id);
  }

  function resetOutputTemplates() {
    saveOutputTemplates(defaultOutputTemplates);
    setSelectedOutputTemplateId(defaultOutputTemplates[0].id);
    setMessage("Output templates reset to defaults.");
  }

  async function uploadTemplateLogo(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setMessage("Choose an image file for the brand logo.");
      event.target.value = "";
      return;
    }
    updateSelectedOutputTemplate("logoDataUrl", await readAsDataUrl(file));
    event.target.value = "";
  }

  async function uploadSourceTemplate(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const templateText = await extractTemplateReferenceText(file);
    updateSelectedOutputTemplate("sourceTemplateName", file.name);
    updateSelectedOutputTemplate("sourceTemplateType", file.type || file.name.split(".").pop() || "");
    updateSelectedOutputTemplate("sourceTemplateDataUrl", await readAsDataUrl(file));
    updateSelectedOutputTemplate("sourceTemplateText", templateText);
    setMessage(`Stored ${file.name} as a maintained template reference.`);
    event.target.value = "";
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

  function createQuickWorkTask() {
    const title = workQuickTaskDraft.title.trim() || task.label;
    const details = workQuickTaskDraft.details.trim();
    createTask({
      title,
      details,
      taskProjectId: projectId,
      templateId: task.id,
      category: task.category,
    });
    setWorkQuickTaskDraft({ title: "", details: "" });
    setCaptureTaskOpen(false);
    window.setTimeout(() => {
      document.getElementById("task-details-section")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 80);
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
      templateId: taskTemplates[0]?.id ?? defaultTaskTemplates[0].id,
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
    const outputTemplate = preferredOutputTemplateForTask(taskProjectId, templateId, outputTemplates) ?? selectedOutputTemplate;
    const newTask: WorkTask = {
      id: createId(),
      projectId: taskProjectId,
      templateId,
      outputTemplateId: outputTemplate.id,
      title,
      details,
      category,
      priority: "Normal",
      dueDate: "",
      reminderAt: "",
      recurrence: "None",
      recurrenceNote: "",
      status: "Open",
      favorite: false,
      statusHistory: [{ status: "Open", changedAt: new Date().toISOString() }],
      checklist,
      input: "",
      assets: [],
      requirements: requirementsLinkedToOutputTemplate(taskRequirements(taskProjectId, templateId, taskTemplates), outputTemplate),
      gptPrompt: "",
      result: "",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    setWorkTasks((current) => [newTask, ...current]);
    void saveTask(newTask);
    openWorkTask(newTask);
    setMessage("Task created. Add dates or reminders if needed.");
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
    if (key === "outputTemplateId") setSelectedOutputTemplateId(value as string);
    if (key === "reminderAt") forgetTriggeredReminder(id);
  }

  function updateActiveWorkTask<K extends keyof WorkTask>(key: K, value: WorkTask[K]) {
    if (!activeWorkTaskId) return;
    updateWorkTask(activeWorkTaskId, key, value);
  }

  function openWorkTask(item: WorkTask) {
    const normalized = normalizeWorkTask(item);
    setTaskBrowserOpen(true);
    setActiveWorkTaskId(normalized.id);
    setProjectId(normalized.projectId);
    setTaskId(normalized.templateId);
    const taskAvailableOutputTemplates = outputTemplates.filter((template) => isOutputTemplateAvailable(template, normalized.projectId));
    const savedOutputTemplate = taskAvailableOutputTemplates.find((template) => template.id === normalized.outputTemplateId);
    const nextOutputTemplate =
      savedOutputTemplate && isOutputTemplateCompatibleWithTask(savedOutputTemplate, normalized.templateId)
        ? savedOutputTemplate
        : preferredOutputTemplateForTask(normalized.projectId, normalized.templateId, taskAvailableOutputTemplates) ?? defaultOutputTemplates[0];
    setSelectedOutputTemplateId(nextOutputTemplate.id);
    setInput(normalized.input);
    setAssets(normalized.assets);
    setRequirements(requirementsLinkedToOutputTemplate(normalized.requirements, nextOutputTemplate));
    setGptPrompt(normalized.gptPrompt);
    setResult(normalized.result);
    setMessage("Task opened in AI mode.");
  }

  function openMobileTask(item: WorkTask, target: "details" | "ai") {
    openWorkTask(item);
    setViewMode(target === "ai" ? "ai" : "work");
    window.setTimeout(() => {
      document.getElementById(target === "details" ? "task-details-section" : "ai-workspace-section")?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    }, 80);
  }

  function openFavoriteTask(item: WorkTask) {
    openWorkTask(item);
    setViewMode("work");
    window.setTimeout(() => {
      document.getElementById("task-details-section")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 80);
  }

  function openNote(note: AppNote) {
    setProjectId(note.projectId);
    setViewMode("notes");
    window.setTimeout(() => {
      document.getElementById(`note-${note.id}`)?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 80);
  }

  function toggleTaskFavorite(item: WorkTask) {
    updateWorkTask(item.id, "favorite", !item.favorite);
  }

  function openCaptureTask() {
    setViewMode("work");
    setCaptureTaskOpen(true);
  }

  function openAiStudio() {
    const linkedTask = activeWorkTask ?? sortedProjectTasks.find((item) => item.status !== "Closed") ?? sortedProjectTasks[0];
    if (linkedTask) openWorkTask(linkedTask);
    setViewMode("ai-engine");
    window.setTimeout(() => {
      document.querySelector(".ai-engine-builder")?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    }, 80);
  }

  function openHomeProjectDashboard() {
    setViewMode("home");
    window.setTimeout(() => {
      document.getElementById("home-project-dashboard")?.scrollIntoView({ behavior: "smooth", block: "start" });
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

    if (permission === "granted" && "serviceWorker" in navigator) {
      await navigator.serviceWorker.ready.catch(() => undefined);
    }

    setNow(Date.now());
    setMessage(permission === "granted" ? "Reminder notifications enabled for this device." : "Reminder notifications were not enabled.");
  }

  function forgetTriggeredReminder(taskId: string) {
    setTriggeredReminderIds((current) => {
      const next = current.filter((item) => !item.startsWith(`${taskId}:`));
      void setMeta(deviceTriggeredReminderMetaKey, next);
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
      status: "Active",
      pinned: false,
      favorite: false,
      createdAt: timestamp,
      updatedAt: timestamp,
    };
    persistNotes([newNote, ...notes]);
    setNoteDraft({ title: "", content: "" });
    setMessage("Note saved.");
  }

  function createQuickWorkNote() {
    if (!workQuickNoteDraft.title.trim() && !workQuickNoteDraft.content.trim()) {
      setMessage("Add a quick note title or detail first.");
      return;
    }

    const timestamp = new Date().toISOString();
    const newNote: AppNote = {
      id: createId(),
      projectId,
      title: workQuickNoteDraft.title.trim() || "Untitled note",
      entries: workQuickNoteDraft.content.trim()
        ? [{ id: createId(), content: workQuickNoteDraft.content.trim(), createdAt: timestamp, updatedAt: timestamp }]
        : [],
      status: "Active",
      pinned: false,
      favorite: false,
      createdAt: timestamp,
      updatedAt: timestamp,
    };
    persistNotes([newNote, ...notes]);
    setWorkQuickNoteDraft({ title: "", content: "" });
    setCaptureNoteOpen(false);
    setMessage("Quick note saved.");
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
        const isPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);
        const canReadText = file.type.startsWith("text/") || /\.(md|txt|csv|json|html)$/i.test(file.name);
        const content = isDocx
          ? await readDocxText(file)
          : isPptx
            ? await readPptxText(file)
          : isPdf
            ? await readPdfText(file)
          : canReadText
            ? await file.text()
            : "This file is attached but cannot be read in the browser. Paste the important text into the input box.";
        const safeContent = content.length > maxTextAssetCharacters
          ? `${content.slice(0, maxTextAssetCharacters)}\n\n[Content truncated at ${maxTextAssetCharacters.toLocaleString()} characters for browser performance.]`
          : content;

        return {
          id: createId(),
          name: file.name,
          type: canReadText || isDocx || isPptx || isPdf ? ("text" as const) : ("file" as const),
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

    const prompt = buildFullLlmPrompt(
      project.name,
      project.context,
      task.label,
      activeWorkTask?.title ?? task.label,
      activeWorkTask?.details ?? "",
      activeWorkTask?.checklist ?? [],
      input,
      assets,
      requirements,
      selectedOutputTemplate,
      inputQuality,
    );
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

  function updateResultValue(value: string) {
    setResult(value);
    updateActiveWorkTask("result", value);
  }

  function replaceSlideJson(index: number, value: string) {
    try {
      const nextSlide = normalizeSlideModel(JSON.parse(value), index, assets, selectedOutputTemplate);
      const nextSlides = slidePreview.map((slide, slideIndex) => (slideIndex === index ? nextSlide : slide));
      updateResultValue(formatSlideDeckJson(nextSlides));
      setMessage(`Updated slide ${index + 1} preview from JSON.`);
    } catch (error) {
      updateResultValue(value);
      setMessage(error instanceof Error ? `Slide ${index + 1} JSON is invalid: ${error.message}` : `Slide ${index + 1} JSON is invalid.`);
    }
  }

  function addStarterSlideJson() {
    const slides = slidePreview.length
      ? slidePreview
      : [
          createFallbackSlide(0, selectedOutputTemplate),
        ];
    updateResultValue(formatSlideDeckJson(slides));
    setMessage("Created editable slide JSON. Change any slide JSON to update the preview.");
  }

  function movePreviewSlide(direction: -1 | 1) {
    setPreviewSlideIndex((current) => Math.min(Math.max(current + direction, 0), Math.max(slidePreview.length - 1, 0)));
  }

  async function downloadResult(format: OutputTemplateFormat) {
    const extension = format === "Markdown" ? "md" : format.toLowerCase();
    const templatedResult = format === "PPTX" ? result : applyOutputTemplate(result, selectedOutputTemplate);
    const blob =
      format === "DOCX"
        ? await toDocxBlob(templatedResult, selectedOutputTemplate)
        : format === "PDF"
          ? await toPdfBlob(templatedResult, selectedOutputTemplate)
        : format === "PPTX"
          ? await toPptxBlob(templatedResult, selectedOutputTemplate, assets)
        : new Blob([templatedResult], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${project.name}-${task.label}-${selectedOutputTemplate.name}`.replace(/[^a-z0-9]+/gi, "-").toLowerCase() + `.${extension}`;
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
      outputTemplateId: selectedOutputTemplate.id,
      title: `${activeWorkTask?.title ?? task.label} - ${new Date().toLocaleString()}`,
      createdAt: new Date().toISOString(),
      input,
      requirements,
      gptPrompt,
      result,
      renderedOutput: applyOutputTemplate(result, selectedOutputTemplate),
    };
    setSavedOutputs((current) => [saved, ...current]);
    void saveOutput(saved);
    setMessage("Saved to project history.");
  }

  function loadSaved(saved: SavedOutput) {
    setTaskId(legacyTemplateMap[saved.taskId] ?? saved.taskId);
    const nextOutputTemplateId = saved.outputTemplateId ?? defaultOutputTemplates[0].id;
    const nextOutputTemplate = outputTemplates.find((item) => item.id === nextOutputTemplateId) ?? defaultOutputTemplates[0];
    setSelectedOutputTemplateId(nextOutputTemplateId);
    const savedTask = workTasks.find((item) => item.id === saved.workTaskId);
    if (savedTask) openWorkTask(savedTask);
    setInput(saved.input);
    setRequirements(requirementsLinkedToOutputTemplate(saved.requirements, nextOutputTemplate));
    setGptPrompt(saved.gptPrompt ?? "");
    setResult(saved.result);
    setMessage("Loaded saved output.");
  }

  function saveAiEngineActivity(activity: {
    prompt: string;
    output: string;
    sourceText: string;
    form: {
      outputTitle: string;
      outputType: string;
      audience: string;
      tone: string;
      outputLength: string;
      outputFormat: string;
      sections: string[];
      additionalInstructions: string;
    };
  }) {
    const linkedTask = activeWorkTask ?? workTasks.find((item) => item.status !== "Closed") ?? workTasks[0];
    if (!linkedTask) {
      setMessage("Create or select a task before saving AI activity.");
      return;
    }

    const linkedOutputTemplate = outputTemplates.find((item) => item.id === linkedTask.outputTemplateId) ?? selectedOutputTemplate;
    const activityRequirements: Requirements = {
      outputType: activity.form.outputType,
      format: linkedOutputTemplate.format === "TXT" || linkedOutputTemplate.format === "DOCX" ? linkedOutputTemplate.format : "Markdown",
      tone: activity.form.tone,
      audience: activity.form.audience,
      length: activity.form.outputLength,
      sections: activity.form.sections.join(", "),
      constraints: activity.form.additionalInstructions,
      imageRequirements: activity.form.outputFormat === "Image prompt" ? activity.form.additionalInstructions : "",
    };
    const saved: SavedOutput = {
      id: createId(),
      projectId: linkedTask.projectId,
      workTaskId: linkedTask.id,
      taskId: linkedTask.templateId,
      outputTemplateId: linkedOutputTemplate.id,
      title: `${activity.form.outputTitle.trim() || linkedTask.title} - ${new Date().toLocaleString()}`,
      createdAt: new Date().toISOString(),
      input: activity.sourceText,
      requirements: activityRequirements,
      gptPrompt: activity.prompt,
      result: activity.output,
      renderedOutput: activity.output ? applyOutputTemplate(activity.output, linkedOutputTemplate) : "",
    };

    setSavedOutputs((current) => [saved, ...current]);
    void saveOutput(saved);
    updateWorkTask(linkedTask.id, "gptPrompt", activity.prompt);
    updateWorkTask(linkedTask.id, "input", activity.sourceText);
    updateWorkTask(linkedTask.id, "requirements", activityRequirements);
    updateWorkTask(linkedTask.id, "result", activity.output);
    setActiveWorkTaskId(linkedTask.id);
    setMessage("AI activity saved to the linked task history.");
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

  function openTasksNavigation() {
    if (typeof window !== "undefined" && window.matchMedia("(max-width: 700px)").matches) {
      setViewMode("mobile");
      setMobileSection("tasks");
      return;
    }

    setViewMode("work");
  }

  return (
    <main className="app-shell">
      <aside className="app-rail" aria-label="Primary navigation">
        <div className="brand-mark">
          <span>AI</span>
          <div>
            <strong>TaskOS AI</strong>
            <small>Notes, tasks, drafts</small>
          </div>
        </div>
        <nav className="rail-nav">
          <button className={viewMode === "home" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("home")} type="button">
            <LayoutDashboard size={17} />
            Home
          </button>
          <button className={viewMode === "favorites" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("favorites")} type="button">
            <Star size={17} />
            Favorites
          </button>
          <button className={viewMode === "work" || viewMode === "mobile" ? "nav-button active" : "nav-button"} onClick={openTasksNavigation} type="button">
            <ListTodo size={17} />
            Tasks
          </button>
          <button className={viewMode === "ai-engine" ? "nav-button active" : "nav-button"} onClick={openAiStudio} type="button">
            <WandSparkles size={17} />
            AI Engine
          </button>
          <button className="nav-button mobile-sync-nav" disabled={syncing} onClick={() => void refreshData()} type="button">
            <History size={17} />
            {syncing ? "Syncing" : "Sync"}
          </button>
          <button className={viewMode === "notes" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("notes")} type="button">
            <StickyNote size={17} />
            Notes
          </button>
          <button className={viewMode === "reminders" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("reminders")} type="button">
            <CalendarClock size={17} />
            Reminders
          </button>
          <button className={viewMode === "settings" ? "nav-button active" : "nav-button"} onClick={() => setViewMode("settings")} type="button">
            <Settings size={17} />
            Settings
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

      <section className={`app-main app-main-${viewMode}`}>
      <section className="topbar">
        <div>
          <p className="eyebrow">AI-enabled task command center</p>
          <h1>{viewMode === "home" ? "Today, tasks, notes, and AI drafts" : viewMode === "ai-engine" ? "AI Generation Engine" : "Tasks, notes, and AI drafting"}</h1>
          <div className="topbar-meta" aria-label="Workspace readiness">
            <span>
              <ShieldCheck size={14} />
              Global-ready UI
            </span>
            <span>{getStorageBackendLabel()}</span>
          </div>
        </div>
        <div className="topbar-actions hero-actions">
          <button className="primary-button" onClick={openCaptureTask} type="button">
            <Plus size={16} />
            Capture task
          </button>
          <button className="ai-action" onClick={openAiStudio} type="button">
            <WandSparkles size={16} />
            Open AI Engine
          </button>
          <button
            className="ghost-button"
            onClick={() => {
              setViewMode("notes");
              setCaptureNoteOpen(true);
            }}
            type="button"
          >
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

      {viewMode === "ai-engine" && (
        <AIPromptBuilder
          activeTask={activeWorkTask}
          onOutputChange={(value) => {
            setResult(value);
            updateActiveWorkTask("result", value);
          }}
          onPromptGenerated={(prompt, sourceText) => {
            setGptPrompt(prompt);
            updateActiveWorkTask("gptPrompt", prompt);
            updateActiveWorkTask("input", sourceText);
          }}
          onSaveActivity={saveAiEngineActivity}
          onTaskSelect={(id) => {
            const selectedTask = workTasks.find((item) => item.id === id);
            if (selectedTask) openWorkTask(selectedTask);
          }}
          taskHistory={activeTaskHistory}
          tasks={sortedWorkTasks}
        />
      )}

      {viewMode === "home" && (
        <Home
          activeWorkTaskId={activeWorkTaskId}
          favoriteNotes={favoriteNotes}
          favoriteTasks={favoriteTasks}
          noteEntryCount={noteEntryCount}
          now={now}
          openAiStudio={openAiStudio}
          openFavoriteTask={openFavoriteTask}
          openHomeProjectDashboard={openHomeProjectDashboard}
          openNote={openNote}
          openWorkTask={openWorkTask}
          projects={projects}
          recentNotes={recentNotes}
          reminderPlanner={reminderPlanner}
          savedOutputs={savedOutputs}
          setViewMode={setViewMode}
          summary={summary}
          toggleTaskFavorite={toggleTaskFavorite}
          activeFocusTasks={activeFocusTasks}
          projectDashboard={projectDashboard}
          projectId={projectId}
          updateProject={updateProject}
        />
      )}

      {viewMode === "favorites" && (
        <section className="favorites-page">
          <section className="panel favorites-section">
            <div className="result-header">
              <h2>
                <Star size={18} />
                Favorites
              </h2>
              <span className="subtle-count">{favoriteTasks.length + favoriteNotes.length}</span>
            </div>
            <div className="favorites-split">
              <section>
                <h3>Tasks</h3>
                <div className="favorite-list">
                  {favoriteTasks.map((item) => (
                    <button className="favorite-list-item" key={item.id} onClick={() => openFavoriteTask(item)} type="button">
                      <strong>{item.title}</strong>
                      <span>{projects[item.projectId].name} - {item.status}</span>
                    </button>
                  ))}
                  {favoriteTasks.length === 0 && <p className="empty compact-empty">No favorite tasks.</p>}
                </div>
              </section>
              <section>
                <h3>Notes</h3>
                <div className="favorite-list">
                  {favoriteNotes.map((note) => (
                    <button className="favorite-list-item note-favorite" key={note.id} onClick={() => openNote(note)} type="button">
                      <strong>{note.title}</strong>
                      <span>{projects[note.projectId].name} - {note.entries.length} entries</span>
                    </button>
                  ))}
                  {favoriteNotes.length === 0 && <p className="empty compact-empty">No favorite notes.</p>}
                </div>
              </section>
            </div>
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
            <button disabled={syncing} onClick={() => void refreshData()} type="button" title="Pull latest data from Supabase">
              <History size={16} />
              {syncing ? "Syncing" : "Sync"}
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
            <div className="mobile-project-strip" aria-label="Mobile project selector">
              {mobileProjectGroups.map((item) => (
                <button
                  className={projectId === item.projectId ? "active" : ""}
                  key={item.projectId}
                  onClick={() => updateProject(item.projectId)}
                  type="button"
                >
                  <span>{projects[item.projectId].name}</span>
                  <strong>{item.total - item.closed}</strong>
                  <small>{item.favorite ? `${item.favorite} fav` : `${item.closed} closed`}</small>
                </button>
              ))}
            </div>
            {mobileCurrentProjectGroup && (
              <div className="mobile-status-groups">
                <section className="mobile-project-title">
                  <div>
                    <strong>{projects[mobileCurrentProjectGroup.projectId].name}</strong>
                    <span>{mobileActiveProjectTasks.length} active - {mobileFavoriteProjectTasks.length} favorites</span>
                  </div>
                  <button className="ghost-button" onClick={() => setViewMode("work")} type="button">
                    Full view
                  </button>
                </section>

                <section className="mobile-task-section">
                  <div className="result-header">
                    <h2>
                      <ListTodo size={18} />
                      Active tasks
                    </h2>
                    <span className="subtle-count">{mobileActiveProjectTasks.length}</span>
                  </div>
                  <div className="mobile-task-list">
                    {mobileActiveProjectTasks.map((item) => (
                      <MobileTaskCard
                        activeWorkTaskId={activeWorkTaskId}
                        item={item}
                        key={item.id}
                        now={now}
                        onOpen={openMobileTask}
                        onToggleFavorite={toggleTaskFavorite}
                      />
                    ))}
                    {mobileActiveProjectTasks.length === 0 && <p className="empty compact-empty">No active tasks in this project.</p>}
                  </div>
                </section>

                <details className="mobile-more-status">
                  <summary>
                    <span>Browse by status</span>
                    <strong>{mobileCurrentProjectGroup.total}</strong>
                  </summary>
                  {mobileStatusOrder.map((status) => (
                  <MobileStatusSection
                    activeWorkTaskId={activeWorkTaskId}
                    items={mobileCurrentProjectGroup.byStatus[status]}
                    key={status}
                    now={now}
                    onOpen={openMobileTask}
                    onToggleFavorite={toggleTaskFavorite}
                    status={status}
                  />
                  ))}
                </details>
              </div>
            )}

            <section className="mobile-overview-panel">
              <div className="mobile-section-header">
                <div>
                  <p className="eyebrow">Mobile task workspace</p>
                  <h2>
                    <Smartphone size={18} />
                    Today
                  </h2>
                </div>
                <button className="primary-button" onClick={() => setMobileSection("capture")} type="button">
                  <Plus size={15} />
                  New
                </button>
              </div>
              <div className="mobile-metric-grid">
                <button onClick={() => setViewMode("work")} type="button">
                  <strong>{summary.total - summary.closed}</strong>
                  <span>Active</span>
                </button>
                <button onClick={() => setViewMode("work")} type="button">
                  <strong>{favoriteTasks.length + favoriteNotes.length}</strong>
                  <span>Favorites</span>
                </button>
                <button onClick={() => setMobileSection("reminders")} type="button">
                  <strong>{reminderPlanner.overdue.length + reminderPlanner.today.length}</strong>
                  <span>Due now</span>
                </button>
                <button onClick={openHomeProjectDashboard} type="button">
                  <strong>{summary.closed}</strong>
                  <span>Closed</span>
                </button>
              </div>
            </section>

            {(favoriteTasks.length > 0 || favoriteNotes.length > 0) && (
              <section className="mobile-favorites">
                <div className="mobile-section-header">
                  <h2>
                    <Star size={18} />
                    Favorites
                  </h2>
                </div>
                <div className="mobile-favorite-strip">
                  {favoriteTasks.map((item) => (
                    <button className="mobile-favorite-card" key={item.id} onClick={() => openMobileTask(item, "details")} type="button">
                      <Star size={15} />
                      <span>
                        <strong>{item.title}</strong>
                        <small>{projects[item.projectId].name} - task</small>
                      </span>
                    </button>
                  ))}
                  {favoriteNotes.map((note) => (
                    <button className="mobile-favorite-card note-favorite" key={note.id} onClick={() => openNote(note)} type="button">
                      <StickyNote size={15} />
                      <span>
                        <strong>{note.title}</strong>
                        <small>{projects[note.projectId].name} - note</small>
                      </span>
                    </button>
                  ))}
                </div>
              </section>
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
              {[...reminderPlanner.overdue, ...reminderPlanner.today, ...reminderPlanner.tomorrow].map((item) => (
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
              {reminderPlanner.overdue.length + reminderPlanner.today.length + reminderPlanner.tomorrow.length === 0 && (
                <p className="empty">No overdue, today, or tomorrow reminders.</p>
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
                <p className="context-copy">Use the main task area for templates, full task editing, prompts, uploads, output saving, and project history.</p>
                <button className="primary-button" onClick={() => setViewMode("work")} type="button">
                  <ArrowRight size={16} />
                  Go to full tasks view
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
                Plan today and upcoming tasks
              </h2>
              <p className="context-copy">
                Review reminders here and enable browser notifications on each device that should alert you.
              </p>
            </div>
            <div className="reminder-hero-actions">
              <div className="reminder-metrics">
                <span><strong>{reminderPlanner.overdue.length}</strong> overdue</span>
                <span><strong>{reminderPlanner.today.length}</strong> today</span>
                <span><strong>{reminderPlanner.tomorrow.length}</strong> tomorrow</span>
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
                emptyText="Nothing planned for tomorrow."
                items={reminderPlanner.tomorrow}
                now={now}
                onClear={clearReminder}
                onOpen={(item) => {
                  openWorkTask(item);
                  setViewMode("work");
                }}
                onSave={saveReminder}
                onSchedule={updateReminder}
                reminderValue={reminderValue}
                title="Tomorrow"
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
                <span className="subtle-count">{reminderPlanner.planningQueue.length}</span>
              </div>
              <div className="reminder-list">
                {reminderPlanner.planningQueue.map((item) => (
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
                {reminderPlanner.planningQueue.length === 0 && <p className="empty">Every open task has a reminder or due date, or there are no open tasks.</p>}
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
              <div className="notes-toolbar">
                {(["Active", "Closed", "All"] as const).map((status) => (
                  <button className={noteStatusFilter === status ? "active" : ""} key={status} onClick={() => setNoteStatusFilter(status)} type="button">
                    {status}
                  </button>
                ))}
                <button className="primary-button" onClick={() => setCaptureNoteOpen(true)} type="button">
                  <Plus size={16} />
                  Capture note
                </button>
              </div>
            </div>
            <p className="context-copy compact-empty">
              Showing {visibleNotes.length} {noteStatusFilter.toLowerCase()} notes and {notes.reduce((total, note) => total + note.entries.length, 0)} total entries.
            </p>
          </section>

          <section className="notes-grid">
            {visibleNotes.map((note) => (
                <article className={`${note.pinned ? "note-card pinned" : "note-card"} ${note.favorite ? "favorite" : ""}`} id={`note-${note.id}`} key={note.id}>
                  <div className="note-card-header">
                    <span className={`project-chip project-${note.projectId}`}>{projects[note.projectId].name}</span>
                    <div className="note-card-actions">
                      <button className={note.favorite ? "ghost-button icon-button favorite-toggle active" : "ghost-button icon-button favorite-toggle"} onClick={() => updateNote(note.id, { favorite: !note.favorite })} type="button" title={note.favorite ? "Remove favorite" : "Add favorite"}>
                        <Star size={15} />
                      </button>
                      <button className="ghost-button icon-button" onClick={() => updateNote(note.id, { pinned: !note.pinned })} type="button" title={note.pinned ? "Unpin note" : "Pin note"}>
                        <Pin size={15} />
                      </button>
                    </div>
                  </div>
                  <input
                    className="note-title-input"
                    value={note.title}
                    onChange={(event) => updateNote(note.id, { title: event.target.value })}
                  />
                  <label className="note-status-field">
                    Status
                    <select value={note.status ?? "Active"} onChange={(event) => updateNote(note.id, { status: event.target.value as AppNote["status"] })}>
                      <option>Active</option>
                      <option>Closed</option>
                    </select>
                  </label>
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
            {visibleNotes.length === 0 && <p className="empty">No {noteStatusFilter.toLowerCase()} notes to show.</p>}
          </section>
        </section>
      )}

      {viewMode === "settings" && (
        <section className="settings-page">
          <section className="panel">
            <div className="result-header">
              <h2>
                <Settings size={18} />
                Master settings
              </h2>
              <button className="ghost-button" onClick={() => saveProjectSettings(defaultProjectSettings)} type="button">
                Reset projects
              </button>
              <button className="primary-button" onClick={addProjectSetting} type="button">
                <Plus size={16} />
                Add project
              </button>
            </div>
            <p className="context-copy">
              Maintain project names, global task templates, statuses, and output templates here. App screens read these master values.
            </p>
          </section>

          <section className="panel settings-section">
            <h2>Projects</h2>
            <div className="settings-grid">
              {(Object.keys(masterProjects) as ProjectId[]).map((id) => (
                <article className="settings-card" key={id}>
                  <span className="project-chip">{id}</span>
                  <label>
                    Project name
                    <input value={masterProjects[id].name} onChange={(event) => updateProjectSetting(id, "name", event.target.value)} />
                  </label>
                  <label>
                    Context
                    <textarea className="small-textarea" value={masterProjects[id].context} onChange={(event) => updateProjectSetting(id, "context", event.target.value)} />
                  </label>
                  <button className="danger-button" disabled={projectIds.length <= 1} onClick={() => deleteProjectSetting(id)} type="button">
                    <Trash2 size={15} />
                    Delete project
                  </button>
                </article>
              ))}
            </div>
          </section>

          <section className="panel settings-section">
            <div className="result-header">
              <h2>Task templates</h2>
              <button className="primary-button" onClick={addTaskTemplate} type="button">
                <Plus size={16} />
                Add task template
              </button>
            </div>
            <div className="settings-list">
              {taskTemplates.map((template) => (
                <article className="settings-template-row" key={template.id}>
                  <strong>{template.id}</strong>
                  <label>
                    Name
                    <input value={template.label} onChange={(event) => updateTaskTemplate(template.id, { label: event.target.value })} />
                  </label>
                  <label>
                    Category
                    <input value={template.category} onChange={(event) => updateTaskTemplate(template.id, { category: event.target.value })} />
                  </label>
                  <label>
                    Description
                    <textarea className="small-textarea" value={template.description} onChange={(event) => updateTaskTemplate(template.id, { description: event.target.value })} />
                  </label>
                  <label>
                    Content goal
                    <input
                      value={template.requirements.outputType}
                      onChange={(event) =>
                        updateTaskTemplate(template.id, {
                          requirements: { ...template.requirements, outputType: event.target.value },
                        })
                      }
                    />
                  </label>
                  <label>
                    Sections
                    <textarea
                      className="small-textarea"
                      placeholder="One section per line, or comma-separated"
                      value={template.requirements.sections}
                      onChange={(event) =>
                        updateTaskTemplate(template.id, {
                          requirements: { ...template.requirements, sections: event.target.value },
                        })
                      }
                    />
                  </label>
                  <label>
                    Constraints
                    <textarea
                      className="small-textarea"
                      value={template.requirements.constraints}
                      onChange={(event) =>
                        updateTaskTemplate(template.id, {
                          requirements: { ...template.requirements, constraints: event.target.value },
                        })
                      }
                    />
                  </label>
                  <button className="danger-button icon-button" disabled={taskTemplates.length <= 1} onClick={() => deleteTaskTemplate(template.id)} type="button" title="Delete task template">
                    <Trash2 size={15} />
                  </button>
                </article>
              ))}
            </div>
          </section>

          <section className="panel settings-section">
            <h2>Statuses</h2>
            <div className="settings-status-grid">
              <section>
                <div className="result-header">
                  <h3>Task statuses</h3>
                  <button className="ghost-button" onClick={addTaskStatus} type="button"><Plus size={15} />Add</button>
                </div>
                <div className="settings-list">
                  {masterTaskStatuses.map((status, index) => (
                    <div className="settings-status-row" key={`${status}-${index}`}>
                      <input
                        value={status}
                        onChange={(event) => {
                          const next = [...masterTaskStatuses];
                          next[index] = event.target.value as TaskStatus;
                          saveMasterStatuses({ ...masterStatuses, task: next });
                        }}
                      />
                      <button className="danger-button icon-button" disabled={masterTaskStatuses.length <= 1} onClick={() => deleteTaskStatus(status)} type="button"><Trash2 size={15} /></button>
                    </div>
                  ))}
                </div>
              </section>
              <section>
                <div className="result-header">
                  <h3>Note statuses</h3>
                  <button className="ghost-button" onClick={addNoteStatus} type="button"><Plus size={15} />Add</button>
                </div>
                <div className="settings-list">
                  {masterNoteStatuses.map((status, index) => (
                    <div className="settings-status-row" key={`${status}-${index}`}>
                      <input
                        value={status}
                        onChange={(event) => {
                          const next = [...masterNoteStatuses];
                          next[index] = event.target.value as "Active" | "Closed";
                          saveMasterStatuses({ ...masterStatuses, note: next });
                        }}
                      />
                      <button className="danger-button icon-button" disabled={masterNoteStatuses.length <= 1} onClick={() => deleteNoteStatus(status)} type="button"><Trash2 size={15} /></button>
                    </div>
                  ))}
                </div>
              </section>
            </div>
          </section>

          <section className="panel settings-section">
            <div className="result-header">
              <h2>Output templates</h2>
              <div className="template-actions">
                <button className="primary-button" onClick={createOutputTemplate} type="button"><Plus size={15} />Add output template</button>
                <button className="ghost-button" onClick={resetOutputTemplates} type="button">Reset defaults</button>
              </div>
            </div>
            <div className="settings-list">
              {outputTemplates.map((template) => (
                <article className="settings-output-row" key={template.id}>
                  <strong>{template.id}</strong>
                  <input value={template.name} onChange={(event) => saveOutputTemplates(outputTemplates.map((item) => item.id === template.id ? { ...item, name: event.target.value, updatedAt: new Date().toISOString() } : item))} />
                  <input value={template.group} onChange={(event) => saveOutputTemplates(outputTemplates.map((item) => item.id === template.id ? { ...item, group: event.target.value, updatedAt: new Date().toISOString() } : item))} />
                  <select value={template.format} onChange={(event) => saveOutputTemplates(outputTemplates.map((item) => item.id === template.id ? { ...item, format: event.target.value as OutputTemplateFormat, updatedAt: new Date().toISOString() } : item))}>
                    <option>Markdown</option>
                    <option>TXT</option>
                    <option value="DOCX">DOCX-ready content</option>
                    <option value="PDF">PDF-ready content</option>
                    <option value="PPTX">PPTX-ready slide content</option>
                  </select>
                  <textarea className="small-textarea" value={template.description} onChange={(event) => saveOutputTemplates(outputTemplates.map((item) => item.id === template.id ? { ...item, description: event.target.value, updatedAt: new Date().toISOString() } : item))} />
                  <button
                    className="ghost-button icon-button"
                    onClick={() => {
                      const timestamp = new Date().toISOString();
                      const copy = { ...template, id: createId(), name: `${template.name} copy`, createdAt: timestamp, updatedAt: timestamp };
                      saveOutputTemplates([copy, ...outputTemplates]);
                      setSelectedOutputTemplateId(copy.id);
                    }}
                    type="button"
                    title="Duplicate output template"
                  >
                    <Plus size={15} />
                  </button>
                  <button className="danger-button icon-button" disabled={outputTemplates.length <= 1} onClick={() => deleteOutputTemplate(template.id)} type="button" title="Delete output template"><Trash2 size={15} /></button>
                </article>
              ))}
            </div>
          </section>
        </section>
      )}

      {(viewMode === "work" || viewMode === "ai") && (
        <>

      <section className="workflow">
        <div className="step done">View tasks</div>
        <ArrowRight size={16} />
        <div className="step done">Open task</div>
        <ArrowRight size={16} />
        <div className={viewMode === "work" ? "step done" : "step"}>Save details</div>
        <ArrowRight size={16} />
        <div className={viewMode === "ai" ? "step done" : "step"}>AI mode</div>
      </section>

      <div className={viewMode === "ai" || viewMode === "work" ? "layout ai-layout" : "layout"}>
        {false && viewMode === "work" && (
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

          <details className="panel template-drawer">
            <summary>
              <span>
                <Clipboard size={18} />
                Output templates
              </span>
              <small>{selectedOutputTemplate.name}</small>
            </summary>
            <div className="output-template-editor">
              <label>
                Template
                <select value={selectedOutputTemplate.id} onChange={(event) => selectOutputTemplate(event.target.value)}>
                  {linkedOutputTemplates.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.name} ({item.group})
                    </option>
                  ))}
                </select>
              </label>
              <p className="context-copy compact-empty">
                Showing output templates linked to the selected input template: {task.label}.
              </p>
              <label>
                Name
                <input value={selectedOutputTemplate.name} onChange={(event) => updateSelectedOutputTemplate("name", event.target.value)} />
              </label>
              <label>
                Description
                <textarea className="small-textarea" value={selectedOutputTemplate.description} onChange={(event) => updateSelectedOutputTemplate("description", event.target.value)} />
              </label>
              <div className="form-grid compact-grid">
                <label>
                  Group
                  <input value={selectedOutputTemplate.group} onChange={(event) => updateSelectedOutputTemplate("group", event.target.value)} />
                </label>
                <label>
                  Format
                  <select value={selectedOutputTemplate.format} onChange={(event) => updateSelectedOutputTemplate("format", event.target.value as OutputTemplateFormat)}>
                    <option>Markdown</option>
                    <option>TXT</option>
                    <option value="DOCX">DOCX-ready content</option>
                    <option value="PDF">PDF-ready content</option>
                    <option value="PPTX">PPTX-ready slide content</option>
                  </select>
                </label>
              </div>
              <div className="form-grid compact-grid">
                <label>
                  Brand name
                  <input value={selectedOutputTemplate.brandName} onChange={(event) => updateSelectedOutputTemplate("brandName", event.target.value)} />
                </label>
                <label>
                  Primary color
                  <input type="color" value={`#${selectedOutputTemplate.primaryColor}`} onChange={(event) => updateSelectedOutputTemplate("primaryColor", event.target.value.replace("#", ""))} />
                </label>
                <label>
                  Secondary color
                  <input type="color" value={`#${selectedOutputTemplate.secondaryColor}`} onChange={(event) => updateSelectedOutputTemplate("secondaryColor", event.target.value.replace("#", ""))} />
                </label>
                <label>
                  Accent color
                  <input type="color" value={`#${selectedOutputTemplate.accentColor}`} onChange={(event) => updateSelectedOutputTemplate("accentColor", event.target.value.replace("#", ""))} />
                </label>
              </div>
              <label className="upload-control compact-upload">
                <Upload size={15} />
                Upload logo
                <input accept="image/*" onChange={(event) => void uploadTemplateLogo(event)} type="file" />
              </label>
              <label className="upload-control compact-upload">
                <Upload size={15} />
                Upload maintained template source
                <input accept=".docx,.pptx,.pdf,.potx,.dotx" onChange={(event) => void uploadSourceTemplate(event)} type="file" />
              </label>
              {(selectedOutputTemplate.logoDataUrl || selectedOutputTemplate.sourceTemplateName) && (
                <div className="template-source-summary">
                  {selectedOutputTemplate.logoDataUrl && <span>Logo attached</span>}
                  {selectedOutputTemplate.sourceTemplateName && <span>Source: {selectedOutputTemplate.sourceTemplateName}</span>}
                  {selectedOutputTemplate.sourceTemplateText && <span>Template text extracted for prompt guidance</span>}
                </div>
              )}
              <label>
                Scope
                <select
                  value={selectedOutputTemplate.scope === "global" ? "global" : "project"}
                  onChange={(event) => updateSelectedOutputTemplate("scope", event.target.value === "global" ? "global" : [projectId])}
                >
                  <option value="global">Global</option>
                  <option value="project">Current project only</option>
                </select>
              </label>
              <label>
                Compatible task IDs
                <input
                  value={selectedOutputTemplate.compatibleTaskIds.join(", ")}
                  onChange={(event) =>
                    updateSelectedOutputTemplate(
                      "compatibleTaskIds",
                      event.target.value
                        .split(",")
                        .map((item) => item.trim())
                        .filter(Boolean),
                    )
                  }
                />
              </label>
              <label>
                Slots
                <textarea
                  className="small-textarea"
                  value={selectedOutputTemplate.slots.join("\n")}
                  onChange={(event) =>
                    updateSelectedOutputTemplate(
                      "slots",
                      event.target.value
                        .split("\n")
                        .map((item) => item.trim())
                        .filter(Boolean),
                    )
                  }
                />
              </label>
              <label>
                Style rules
                <textarea className="small-textarea" value={selectedOutputTemplate.style} onChange={(event) => updateSelectedOutputTemplate("style", event.target.value)} />
              </label>
              <div className="template-actions">
                <button className="ghost-button" onClick={duplicateOutputTemplate} type="button">
                  <Plus size={15} />
                  Duplicate
                </button>
                <button className="ghost-button" onClick={resetOutputTemplates} type="button">
                  Reset defaults
                </button>
                <button className="danger-button" disabled={outputTemplates.length <= 1} onClick={() => deleteOutputTemplate(selectedOutputTemplate.id)} type="button">
                  <Trash2 size={15} />
                  Delete
                </button>
              </div>
            </div>
          </details>

        </aside>
        )}

        <section className="work-area">
          {viewMode === "ai" && activeWorkTask && (
            <section className="panel ai-task-link">
              <div>
                <p className="eyebrow">AI linked to task</p>
                <h2>{activeWorkTask.title}</h2>
                <p className="context-copy">{project.name} - {activeWorkTask.category} - {activeWorkTask.status}</p>
              </div>
              <button className="ghost-button" onClick={() => setViewMode("work")} type="button">
                <ListTodo size={16} />
                Back to task
              </button>
            </section>
          )}

          {viewMode === "ai" && !activeWorkTask && (
            <section className="panel">
              <p className="empty">Select a task first, then open the AI workspace for that task.</p>
              <button className="primary-button" onClick={() => setViewMode("work")} type="button">
                <ListTodo size={16} />
                View tasks
              </button>
            </section>
          )}

          {viewMode === "work" && (
          <>
          <section className="panel task-top-panel">
            <div className="result-header">
              <h2>
                <BarChart3 size={18} />
                {project.name} status
              </h2>
              <span className="subtle-count">{sortedProjectTasks.length}</span>
            </div>
            <div className="project-filter-strip" aria-label="Project filters">
              {(Object.keys(projects) as ProjectId[]).map((id) => {
                const stats = projectDashboard.find((item) => item.projectId === id);
                return (
                  <button className={projectId === id ? "active" : ""} key={id} onClick={() => updateProject(id)} type="button">
                    <strong>{projects[id].name}</strong>
                    <span>{stats?.total ?? 0} tasks</span>
                  </button>
                );
              })}
            </div>
            <div className="project-stat-columns task-status-summary">
              <div className="project-stat-row status-open"><span>Open</span><strong>{selectedProjectStats?.open ?? 0}</strong></div>
              <div className="project-stat-row status-in-progress"><span>In progress</span><strong>{selectedProjectStats?.inProgress ?? 0}</strong></div>
              <div className="project-stat-row status-blocked"><span>Blocked</span><strong>{selectedProjectStats?.blocked ?? 0}</strong></div>
              <div className="project-stat-row status-to-do-later"><span>To do later</span><strong>{selectedProjectStats?.toDoLater ?? 0}</strong></div>
              <div className="project-stat-row status-closed"><span>Closed</span><strong>{selectedProjectStats?.closed ?? 0}</strong></div>
            </div>
          </section>

          {/* Task Browser - Primary */}
          <details
            className="panel task-browser-primary"
            onToggle={(event) => setTaskBrowserOpen(event.currentTarget.open)}
            open={taskBrowserOpen}
          >
            <summary>
              <span>
                <ListTodo size={18} />
                Browse {project.name} tasks
              </span>
              <strong>{visibleProjectTasks.length}</strong>
            </summary>
            {sortedProjectTasks.length === 0 ? (
              <p className="empty">Create your first task for {project.name}.</p>
            ) : (
              <>
                <div className="task-list-toolbar" aria-label="Task list filters">
                  {(["Active", "Favorites", "Open", "In Progress", "Blocked", "To Do Later", "Closed", "All"] as const).map((filter) => (
                    <button
                      className={taskListFilter === filter ? "active" : ""}
                      key={filter}
                      onClick={() => setTaskListFilter(filter)}
                      type="button"
                    >
                      {filter}
                    </button>
                  ))}
                </div>
                <div className="all-task-list compact-task-list primary-task-list">
                  {visibleProjectTasks.map((item) => (
                    <TaskListRow
                      activeWorkTaskId={activeWorkTaskId}
                      key={item.id}
                      onOpen={openFavoriteTask}
                      onToggleFavorite={toggleTaskFavorite}
                      projectLabel={projects[item.projectId].name}
                      task={item}
                    />
                  ))}
                  {visibleProjectTasks.length === 0 && <p className="empty">No tasks match this view.</p>}
                </div>
              </>
            )}
          </details>

          {taskBrowserOpen && activeWorkTask ? (
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
                <button className={activeWorkTask.favorite ? "ghost-button favorite-toggle active" : "ghost-button favorite-toggle"} onClick={() => toggleTaskFavorite(activeWorkTask)} type="button">
                  <Star size={16} />
                  {activeWorkTask.favorite ? "Favorited" : "Favorite"}
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
                projects={masterProjects}
                templates={taskTemplates}
                reminderValue={reminderValue(activeWorkTask)}
                statuses={masterTaskStatuses}
                onTemplateChange={(nextProjectId, nextTemplateId) => {
                  const nextOutputTemplate = preferredOutputTemplateForTask(nextProjectId, nextTemplateId, outputTemplates) ?? selectedOutputTemplate;
                  const nextRequirements = requirementsLinkedToOutputTemplate(taskRequirements(nextProjectId, nextTemplateId, taskTemplates), nextOutputTemplate);
                  const nextTemplate = taskTemplates.find((item) => item.id === nextTemplateId) ?? taskTemplates[0] ?? defaultTaskTemplates[0];
                  setTaskId(nextTemplateId);
                  setSelectedOutputTemplateId(nextOutputTemplate.id);
                  setRequirements(nextRequirements);
                  updateWorkTask(activeWorkTask.id, "templateId", nextTemplateId);
                  updateWorkTask(activeWorkTask.id, "category", nextTemplate.category);
                  updateWorkTask(activeWorkTask.id, "outputTemplateId", nextOutputTemplate.id);
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
              <section className="history-panel">
                <div className="result-header">
                  <h3>
                    <History size={17} />
                    AI activity history
                  </h3>
                  <span className="subtle-count">{activeTaskHistory.length}</span>
                </div>
                <div className="history-list">
                  {activeTaskHistory.map((saved) => (
                    <button className="history-item" key={saved.id} onClick={() => loadSaved(saved)} type="button">
                      <strong>{saved.title}</strong>
                      <span>{new Date(saved.createdAt).toLocaleString()}</span>
                      <span>{saved.input ? "Source saved" : "No source saved"} - {saved.gptPrompt ? "Prompt saved" : "No prompt saved"} - {saved.result ? "Output saved" : "No output saved"}</span>
                    </button>
                  ))}
                  {activeTaskHistory.length === 0 && <p className="empty compact-empty">No AI prompt or pasted output has been saved for this task yet.</p>}
                </div>
              </section>
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
          ) : taskBrowserOpen ? (
            <section className="panel">
              <p className="empty">Select or create a task to expand it into AI mode.</p>
            </section>
          ) : null}

          </>
          )}

          {false && <details className="panel task-browse-panel" open={sortedProjectTasks.length <= 6}>
            <summary>
              <span>
                <ListTodo size={18} />
                Browse {project.name} tasks
              </span>
              <strong>{visibleProjectTasks.length}</strong>
            </summary>
            {sortedProjectTasks.length === 0 ? (
              <p className="empty">Create your first task for {project.name}. It can be a document, summary, email, or checklist-style task.</p>
            ) : (
              <>
                <div className="task-list-toolbar" aria-label="Task list filters">
                  {(["Active", "Favorites", "Open", "In Progress", "Blocked", "To Do Later", "Closed", "All"] as const).map((filter) => (
                    <button
                      className={taskListFilter === filter ? "active" : ""}
                      key={filter}
                      onClick={() => setTaskListFilter(filter)}
                      type="button"
                    >
                      {filter}
                    </button>
                  ))}
                </div>
                <div className="all-task-list compact-task-list">
                  {visibleProjectTasks.map((item) => (
                    <TaskListRow
                      activeWorkTaskId={activeWorkTaskId}
                      key={item.id}
                      onOpen={openFavoriteTask}
                      onToggleFavorite={toggleTaskFavorite}
                      projectLabel={projects[item.projectId].name}
                      task={item}
                    />
                  ))}
                  {visibleProjectTasks.length === 0 && <p className="empty compact-empty">No tasks match this view.</p>}
                </div>
              </>
            )}
          </details>}

          {activeWorkTask && viewMode === "ai" && (
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
                  Upload PDF, DOCX, PPTX, text, or images
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
                  Content brief
                </h2>
                <p className="context-copy">
                  The task template defines the content intent. The selected output template controls format, slots, and export structure.
                </p>
                <div className="form-grid">
                  <label>
                    Content goal
                    <input value={requirements.outputType} onChange={(event) => updateRequirement("outputType", event.target.value)} />
                  </label>
                  <label>
                    Output template
                    <select value={selectedOutputTemplate.id} onChange={(event) => selectOutputTemplate(event.target.value)}>
                      {linkedOutputTemplates.map((item) => (
                        <option key={item.id} value={item.id}>
                          {item.name}
                        </option>
                      ))}
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
                    <textarea
                      className="small-textarea"
                      value={requirements.sections}
                      onChange={(event) => updateRequirement("sections", event.target.value)}
                      placeholder="Example: Title, Context, Options, Recommendation, Next steps"
                    />
                  </label>
                  <div className="linked-template-field">
                    <span>Template format</span>
                    <strong>{manualPromptFormatLabel(selectedOutputTemplate.format)}</strong>
                  </div>
                  <div className="linked-template-field">
                    <span>Template group</span>
                    <strong>{selectedOutputTemplate.group}</strong>
                  </div>
                </div>
                {selectedOutputTemplate.format === "DOCX" && (
                  <p className="field-helper">
                    ChatGPT will return formatted text. The app must export the final content to .docx if a real Word file is required.
                  </p>
                )}
                <div className="linked-template-slots">
                  <span>Output template slots</span>
                  <p>{selectedOutputTemplate.slots.join(", ")}</p>
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

              <section className={`panel quality-panel ${inputQuality.status === "Ready" ? "ready" : "needs-detail"}`}>
                <h2>
                  <Check size={18} />
                  Input quality
                </h2>
                <div className="quality-status">
                  <strong>{inputQuality.status}</strong>
                  <span>{selectedOutputTemplate.name} - {selectedOutputTemplate.format}</span>
                </div>
                {inputQuality.checks.length > 0 && (
                  <ul>
                    {inputQuality.checks.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                )}
                {inputQuality.warnings.length > 0 && (
                  <ul>
                    {inputQuality.warnings.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ul>
                )}
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
                <p className="context-copy">
                  Export will use {selectedOutputTemplate.name}
                  {selectedOutputTemplate.sourceTemplateName ? ` with source template ${selectedOutputTemplate.sourceTemplateName}` : ""}. {isPresentationOutput ? "Paste or edit slide JSON; the preview and PPTX export use the same slide model." : "Paste clean ChatGPT content; markdown headings and image placeholders will be converted into formatted output."}
                </p>
                {!isPresentationOutput && (
                  <textarea
                    className="output-textarea"
                    value={result}
                    onChange={(event) => updateResultValue(event.target.value)}
                    placeholder="Paste the answer from ChatGPT Plus here. This is the final output you can save or export."
                  />
                )}
                {isPresentationOutput && (
                  <div className="builder-workspace">
                    <section className="slide-json-panel" aria-label="Slide JSON editor">
                      <div className="builder-toolbar">
                        <strong>Slide JSON</strong>
                        <button onClick={addStarterSlideJson} type="button">
                          <Plus size={16} />
                          {slidePreview.length ? "Normalize JSON" : "Create JSON"}
                        </button>
                      </div>
                      {slideDeck.error && <p className="json-warning">{slideDeck.error}</p>}
                      {slidePreview.length ? (
                        <div className="slide-json-list">
                          {slidePreview.map((slide, index) => (
                            <label className="slide-json-editor" key={slide.id}>
                              <span>Slide {index + 1}</span>
                              <textarea
                                value={JSON.stringify(serializeSlideForJson(slide), null, 2)}
                                onChange={(event) => replaceSlideJson(index, event.target.value)}
                                spellCheck={false}
                              />
                            </label>
                          ))}
                        </div>
                      ) : (
                        <textarea
                          className="output-textarea"
                          value={result}
                          onChange={(event) => updateResultValue(event.target.value)}
                          placeholder='Paste slide JSON here, for example: { "slides": [{ "title": "Opening", "layout": "title", "background": "#ffffff", "bullets": [], "images": [] }] }'
                          spellCheck={false}
                        />
                      )}
                    </section>
                    <section className="slide-preview-panel" aria-label="Slide preview">
                      <div className="result-header preview-toolbar">
                        <h3>Preview</h3>
                        <div className="preview-toolbar-actions">
                          <span className="subtle-count">{slidePreview.length}</span>
                          <button disabled={!slidePreview.length} onClick={() => setFullScreenPreviewOpen(true)} type="button">
                            <Maximize2 size={16} />
                            Full screen
                          </button>
                        </div>
                      </div>
                      <div className="slide-preview-grid">
                        {slidePreview.map((slide, index) => (
                          <button
                            className={index === activePreviewSlideIndex ? "slide-preview-item active" : "slide-preview-item"}
                            key={slide.id}
                            onClick={() => setPreviewSlideIndex(index)}
                            type="button"
                          >
                            <SlideCanvas slide={slide} slideNumber={index + 1} template={selectedOutputTemplate} />
                            <SlideMeta slide={slide} />
                          </button>
                        ))}
                      </div>
                    </section>
                  </div>
                )}
                {!isPresentationOutput && result && (
                  <section className="panel slide-preview-panel">
                    <div className="result-header">
                      <h3>Slide preview</h3>
                      <span className="subtle-count">{slidePreview.length}</span>
                    </div>
                    <div className="slide-preview-grid">
                      {slidePreview.map((slide, index) => (
                        <article className="slide-card" key={`${index}-${slide.title}`}>
                          <div className="slide-card-title">
                            <strong>{slide.title}</strong>
                            <span>Slide {index + 1}</span>
                          </div>
                          {slide.bullets.length > 0 && (
                            <ul className="slide-card-bullets">
                              {slide.bullets.slice(0, 4).map((line, lineIndex) => (
                                <li key={lineIndex}>{line}</li>
                              ))}
                            </ul>
                          )}
                          {slide.speakerNotes && <p className="slide-card-notes"><strong>Notes:</strong> {slide.speakerNotes}</p>}
                          {slide.images.length > 0 && (
                            <p className="slide-card-images">
                              <strong>Images:</strong> {slide.images.map((image) => image.label).join(", ")}
                            </p>
                          )}
                        </article>
                      ))}
                    </div>
                  </section>
                )}
                <div className="export-bar">
                  <button disabled={!result} onClick={copyResult} type="button">
                    <Clipboard size={16} />
                    Copy output
                  </button>
                  <button disabled={!result} onClick={() => void downloadResult("DOCX")} type="button">
                    <Download size={16} />
                    Export to DOCX
                  </button>
                  <button disabled={!result} onClick={() => void downloadResult("PDF")} type="button">
                    <Download size={16} />
                    PDF
                  </button>
                  <button disabled={!result} onClick={() => void downloadResult("PPTX")} type="button">
                    <Download size={16} />
                    PPTX
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
      {captureTaskOpen && (
        <div className="capture-modal" role="dialog" aria-modal="true" aria-label="Capture task">
          <section className="capture-dialog">
            <div className="result-header">
              <h2>
                <Plus size={18} />
                Capture task
              </h2>
              <button className="ghost-button icon-button" onClick={() => setCaptureTaskOpen(false)} type="button" title="Close capture task">
                <X size={17} />
              </button>
            </div>
            <div className="form-grid compact-grid">
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
                Task type
                <select value={taskId} onChange={(event) => selectTemplate(event.target.value)}>
                  {project.tasks.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.label}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <label>
              Title
              <input
                autoFocus
                value={workQuickTaskDraft.title}
                onChange={(event) => setWorkQuickTaskDraft((current) => ({ ...current, title: event.target.value }))}
                placeholder={task.label}
              />
            </label>
            <label>
              Details
              <textarea
                className="small-textarea"
                value={workQuickTaskDraft.details}
                onChange={(event) => setWorkQuickTaskDraft((current) => ({ ...current, details: event.target.value }))}
                placeholder="Add context, notes, links, or the outcome you need."
              />
            </label>
            <div className="export-bar">
              <button className="primary-button" onClick={createQuickWorkTask} type="button">
                <ListTodo size={16} />
                Create task
              </button>
              <button className="ghost-button" onClick={() => setCaptureTaskOpen(false)} type="button">
                Cancel
              </button>
            </div>
          </section>
        </div>
      )}
      {captureNoteOpen && (
        <div className="capture-modal" role="dialog" aria-modal="true" aria-label="Capture note">
          <section className="capture-dialog">
            <div className="result-header">
              <h2>
                <StickyNote size={18} />
                Capture note
              </h2>
              <button className="ghost-button icon-button" onClick={() => setCaptureNoteOpen(false)} type="button" title="Close capture note">
                <X size={17} />
              </button>
            </div>
            <div className="form-grid compact-grid">
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
                Status
                <select defaultValue="Active" disabled>
                  <option>Active</option>
                </select>
              </label>
            </div>
            <label>
              Title
              <input
                autoFocus
                value={workQuickNoteDraft.title}
                onChange={(event) => setWorkQuickNoteDraft((current) => ({ ...current, title: event.target.value }))}
                placeholder="Reference, decision, contact detail..."
              />
            </label>
            <label>
              Information
              <textarea
                className="small-textarea"
                value={workQuickNoteDraft.content}
                onChange={(event) => setWorkQuickNoteDraft((current) => ({ ...current, content: event.target.value }))}
                placeholder="Store important information that is not yet a task."
              />
            </label>
            <div className="export-bar">
              <button className="primary-button" onClick={createQuickWorkNote} type="button">
                <Save size={16} />
                Save note
              </button>
              <button className="ghost-button" onClick={() => setCaptureNoteOpen(false)} type="button">
                Cancel
              </button>
            </div>
          </section>
        </div>
      )}
      {fullScreenPreviewOpen && activePreviewSlide && (
        <div className="preview-modal" role="dialog" aria-modal="true" aria-label="Full screen slide preview">
          <div className="preview-modal-header">
            <div>
              <span>Slide {activePreviewSlideIndex + 1} of {slidePreview.length}</span>
              <strong>{activePreviewSlide.title}</strong>
            </div>
            <div className="preview-modal-actions">
              <button disabled={activePreviewSlideIndex === 0} onClick={() => movePreviewSlide(-1)} type="button" title="Previous slide">
                <ChevronLeft size={17} />
              </button>
              <button disabled={activePreviewSlideIndex >= slidePreview.length - 1} onClick={() => movePreviewSlide(1)} type="button" title="Next slide">
                <ChevronRight size={17} />
              </button>
              <button onClick={() => setFullScreenPreviewOpen(false)} type="button">
                <Minimize2 size={16} />
                Exit
              </button>
              <button onClick={() => setFullScreenPreviewOpen(false)} type="button" title="Close full screen preview">
                <X size={17} />
              </button>
            </div>
          </div>
          <div className="preview-modal-stage">
            <SlideCanvas isFullScreen slide={activePreviewSlide} slideNumber={activePreviewSlideIndex + 1} template={selectedOutputTemplate} />
          </div>
          <div className="preview-modal-meta">
            <SlideMeta slide={activePreviewSlide} />
          </div>
        </div>
      )}
    </main>
  );
}

function SlideCanvas({
  isFullScreen = false,
  slide,
  slideNumber,
  template,
}: {
  isFullScreen?: boolean;
  slide: SlideModel;
  slideNumber: number;
  template: OutputTemplate;
}) {
  return (
    <article
      className={`slide-canvas slide-layout-${slide.layout}${isFullScreen ? " full-screen-slide" : ""}`}
      style={{ background: slide.background }}
    >
      <div className="slide-brand-strip" style={{ backgroundColor: `#${template.primaryColor}` }} />
      <header className="slide-canvas-header">
        <div>
          <span>Slide {slideNumber} / {slide.layout}</span>
          <h4>{slide.title}</h4>
        </div>
        {template.logoDataUrl ? (
          <img alt={brandedTitle(template)} src={template.logoDataUrl} />
        ) : (
          <strong>{brandedTitle(template)}</strong>
        )}
      </header>
      <div className="slide-canvas-body">
        <div className="slide-copy">
          {slide.keyMessage && <p className="slide-key-message">{slide.keyMessage}</p>}
          {slide.bullets.length > 0 && (
            <ul>
              {slide.bullets.slice(0, 5).map((line, lineIndex) => (
                <li key={lineIndex}>{line}</li>
              ))}
            </ul>
          )}
        </div>
        <div className="slide-visual">
          {slide.imageAssets[0] ? (
            <img alt={slide.images[0]?.alt || slide.images[0]?.label || "Slide visual"} src={slide.imageAssets[0].content} />
          ) : (
            <div className="slide-visual-placeholder">
              <Image size={22} />
              <span>{slide.images[0]?.label || slide.visualDirection || "Visual direction"}</span>
            </div>
          )}
        </div>
      </div>
      <footer className="slide-canvas-footer">
        <span>{brandedTitle(template)}</span>
        <span>{slideNumber}</span>
      </footer>
    </article>
  );
}

function SlideMeta({ slide }: { slide: SlideModel }) {
  return (
    <div className="slide-meta">
      <span>{slide.background}</span>
      <span>{slide.layout}</span>
      <span>{slide.images.length ? `Images: ${slide.images.map((image) => image.label).join(", ")}` : "No images"}</span>
      <span>{slide.speakerNotes ? `Notes: ${slide.speakerNotes}` : "No speaker notes"}</span>
    </div>
  );
}

function TaskListRow({
  activeWorkTaskId,
  onOpen,
  onToggleFavorite,
  projectLabel,
  task,
}: {
  activeWorkTaskId: string;
  onOpen: (task: WorkTask) => void;
  onToggleFavorite: (task: WorkTask) => void;
  projectLabel: string;
  task: WorkTask;
}) {
  return (
    <article className={`${taskClassName(task, activeWorkTaskId)} task-row`}>
      <button className="task-row-main" onClick={() => onOpen(task)} type="button">
        <span className="task-row-title">
          <strong>{task.title}</strong>
        </span>
        <span className={`status-pill status-${statusSlug(task.status)}`}>{task.status}</span>
        <small>{task.category}</small>
        <small>{task.priority}</small>
        <small>{task.recurrence && task.recurrence !== "None" ? task.recurrence : taskDateLabel(task)}</small>
      </button>
      <button
        aria-label={task.favorite ? `Remove ${task.title} from favorites` : `Add ${task.title} to favorites`}
        className={task.favorite ? "ghost-button icon-button favorite-toggle active" : "ghost-button icon-button favorite-toggle"}
        onClick={() => onToggleFavorite(task)}
        type="button"
        title={task.favorite ? "Remove favorite" : "Add favorite"}
      >
        <Star size={15} />
      </button>
    </article>
  );
}

function TaskEditor({
  task,
  onChange,
  onReminderChange,
  onReminderSave,
  onTemplateChange,
  projects,
  reminderValue,
  statuses,
  templates,
}: {
  task: WorkTask;
  onChange: <K extends keyof WorkTask>(id: string, key: K, value: WorkTask[K]) => void;
  onReminderChange: (id: string, value: string) => void;
  onReminderSave: (id: string) => void;
  onTemplateChange: (projectId: ProjectId, templateId: string) => void;
  projects: ProjectSettings;
  reminderValue: string;
  statuses: TaskStatus[];
  templates: TaskTemplate[];
}) {
  const availableTemplates = templates.length ? templates : defaultTaskTemplates;

  return (
    <div className="task-editor">
      <label>
        Project
        <select
          value={task.projectId}
          onChange={(event) => {
            const nextProjectId = event.target.value as ProjectId;
            const nextTemplateId = availableTemplates[0]?.id ?? defaultTaskTemplates[0].id;
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
        Recurrence
        <select value={task.recurrence ?? "None"} onChange={(event) => onChange(task.id, "recurrence", event.target.value as WorkTask["recurrence"])}>
          <option>None</option>
          <option>Daily</option>
          <option>Weekly</option>
          <option>Monthly</option>
          <option>Quarterly</option>
          <option>Yearly</option>
        </select>
      </label>
      <label>
        Recurrence note
        <input
          value={task.recurrenceNote ?? ""}
          onChange={(event) => onChange(task.id, "recurrenceNote", event.target.value)}
          placeholder="How this task repeats or what to update"
        />
      </label>
      <label>
        Status
        <select value={task.status} onChange={(event) => onChange(task.id, "status", event.target.value as TaskStatus)}>
          {statuses.map((status) => (
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
  onToggleFavorite,
  status,
}: {
  activeWorkTaskId: string;
  items: WorkTask[];
  now: number;
  onOpen: (task: WorkTask, target: "details" | "ai") => void;
  onToggleFavorite: (task: WorkTask) => void;
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
          <MobileTaskCard
            activeWorkTaskId={activeWorkTaskId}
            item={item}
            key={item.id}
            now={now}
            onOpen={onOpen}
            onToggleFavorite={onToggleFavorite}
          />
        ))}
        {items.length === 0 && <p className="empty compact-empty">No {status.toLowerCase()} tasks.</p>}
      </div>
    </details>
  );
}

function MobileTaskCard({
  activeWorkTaskId,
  item,
  now,
  onOpen,
  onToggleFavorite,
}: {
  activeWorkTaskId: string;
  item: WorkTask;
  now: number;
  onOpen: (task: WorkTask, target: "details" | "ai") => void;
  onToggleFavorite: (task: WorkTask) => void;
}) {
  const dueLabel = taskDateLabel(item);

  return (
    <article className={`${taskClassName(item, activeWorkTaskId)} mobile-task-card`}>
      <div className="mobile-task-card-header">
        <div>
          <strong>{item.title}</strong>
          <div className="mobile-task-card-meta">
            <span className={`status-pill status-${statusSlug(item.status)}`}>{item.status}</span>
            <span className={`priority-pill priority-${item.priority.toLowerCase()}`}>{item.priority}</span>
            {dueLabel !== "No due date" && <small>{dueLabel}</small>}
            {isValidDateTime(item.reminderAt) && <small>{reminderLabel(item, now)}</small>}
          </div>
        </div>
        <button
          aria-label={item.favorite ? `Remove ${item.title} from favorites` : `Add ${item.title} to favorites`}
          className={item.favorite ? "ghost-button icon-button favorite-toggle active" : "ghost-button icon-button favorite-toggle"}
          onClick={() => onToggleFavorite(item)}
          type="button"
          title={item.favorite ? "Remove favorite" : "Add favorite"}
        >
          <Star size={15} />
        </button>
      </div>
      <div className="mobile-task-actions">
        <button className="ghost-button" onClick={() => onOpen(item, "details")} type="button">
          <ListTodo size={15} />
          Details
        </button>
        <button className="primary-button" onClick={() => onOpen(item, "ai")} type="button">
          <Sparkles size={15} />
          AI
        </button>
      </div>
    </article>
  );
}

type InputQuality = {
  status: "Ready" | "Needs detail";
  checks: string[];
  warnings: string[];
};

function buildInputQuality(missingDetails: string[], assets: InputAsset[]): InputQuality {
  const warnings: string[] = [];
  const checks: string[] = [];
  const imageAssets = assets.filter((asset) => asset.type === "image");
  const unreadableAssets = assets.filter((asset) => asset.type === "file");
  const truncatedAssets = assets.filter((asset) => asset.content.includes("[Content truncated at"));
  const readableAssets = assets.filter((asset) => asset.type === "text");

  if (readableAssets.length) checks.push(`${readableAssets.length} readable document${readableAssets.length === 1 ? "" : "s"} included in the prompt.`);
  if (imageAssets.length) warnings.push("Images are only listed by filename. Upload them directly to ChatGPT or describe the important content.");
  if (unreadableAssets.length) warnings.push("Some attached files cannot be read in the browser. Paste the important text into the input box.");
  if (readableAssets.some((asset) => asset.content.includes("no selectable text was found"))) {
    warnings.push("At least one PDF had no selectable text. For scanned PDFs, paste the important text manually.");
  }
  if (truncatedAssets.length) warnings.push("Some extracted source text was truncated for browser performance. Paste any missing critical content manually.");
  missingDetails.forEach((item) => warnings.push(item));

  if (!warnings.length) checks.push("Prompt has enough source material and output requirements to generate directly.");

  return {
    status: warnings.length ? "Needs detail" : "Ready",
    checks,
    warnings,
  };
}

function missingDetailsForTask(task: WorkTask) {
  const missing: string[] = [];
  if (task.input.trim().length < 20 && task.assets.length === 0) missing.push("Add source input, notes, or a readable file.");
  if (!task.requirements.outputType.trim()) missing.push("Choose the output type you need.");
  if (!task.requirements.audience.trim()) missing.push("Describe who the output is for.");
  if (!task.requirements.sections.trim()) missing.push("List the required sections or structure.");
  return missing;
}

function taskInputQualityStatus(task: WorkTask) {
  return `AI input: ${buildInputQuality(missingDetailsForTask(task), task.assets).status.toLowerCase()}`;
}

function isOutputTemplateAvailable(template: OutputTemplate, projectId: ProjectId) {
  return template.scope === "global" || template.scope.includes(projectId);
}

function isOutputTemplateCompatibleWithTask(template: OutputTemplate, templateId: string) {
  return template.compatibleTaskIds.length === 0 || template.compatibleTaskIds.includes(templateId);
}

function compatibleOutputTemplates(templates: OutputTemplate[], templateId: string) {
  const compatible = templates.filter((template) => isOutputTemplateCompatibleWithTask(template, templateId));
  return compatible.length ? compatible : templates;
}

function preferredOutputTemplateForTask(projectId: ProjectId, templateId: string, templates: OutputTemplate[]) {
  const available = templates.filter((template) => isOutputTemplateAvailable(template, projectId));
  return compatibleOutputTemplates(available, templateId)[0];
}

function normalizeOutputTemplates(templates: OutputTemplate[]) {
  const timestamp = new Date().toISOString();
  const mergedTemplates = [
    ...templates,
    ...defaultOutputTemplates.filter((defaultTemplate) => !templates.some((template) => template.id === defaultTemplate.id)),
  ];
  const normalized = mergedTemplates
    .filter((template) => template?.id && template.name)
    .map((template) => ({
      ...template,
      description: template.description ?? "",
      group: template.group || "Document",
      format: normalizeOutputTemplateFormat(template.format),
      brandName: template.brandName || template.name || defaultBrand.brandName,
      primaryColor: normalizeHexColor(template.primaryColor, defaultBrand.primaryColor),
      secondaryColor: normalizeHexColor(template.secondaryColor, defaultBrand.secondaryColor),
      accentColor: normalizeHexColor(template.accentColor, defaultBrand.accentColor),
      logoDataUrl: template.logoDataUrl ?? "",
      sourceTemplateName: template.sourceTemplateName ?? "",
      sourceTemplateType: template.sourceTemplateType ?? "",
      sourceTemplateDataUrl: template.sourceTemplateDataUrl ?? "",
      sourceTemplateText: template.sourceTemplateText ?? "",
      scope: normalizeOutputTemplateScope(template.scope),
      compatibleTaskIds: Array.isArray(template.compatibleTaskIds) ? template.compatibleTaskIds : [],
      slots: Array.isArray(template.slots) && template.slots.length ? template.slots : ["Title", "Main content", "Next steps"],
      style: template.style ?? "",
      createdAt: template.createdAt || timestamp,
      updatedAt: template.updatedAt || template.createdAt || timestamp,
    }));

  return normalized.length ? normalized : defaultOutputTemplates;
}

function normalizeHexColor(value: string | undefined, fallback: string) {
  const cleaned = (value ?? "").replace("#", "").trim();
  return /^[0-9a-f]{6}$/i.test(cleaned) ? cleaned : fallback;
}

function normalizeProjectSettings(settings: ProjectSettings): ProjectSettings {
  const next = {} as ProjectSettings;
  const ids = Array.from(new Set([...Object.keys(defaultProjectSettings), ...Object.keys(settings ?? {})])) as ProjectId[];
  for (const id of ids) {
    const project = settings?.[id] ?? defaultProjectSettings[id];
    if (!project) continue;
    next[id] = {
      name: project.name || defaultProjectSettings[id]?.name || id,
      context: project.context ?? defaultProjectSettings[id]?.context ?? "",
    };
  }
  return Object.keys(next).length ? next : defaultProjectSettings;
}

function normalizeTaskTemplates(templates: TaskTemplate[]) {
  const mergedTemplates = [
    ...(Array.isArray(templates) ? templates : []),
    ...defaultTaskTemplates.filter((defaultTemplate) => !templates?.some((template) => template.id === defaultTemplate.id)),
  ];
  const normalized = mergedTemplates
    .filter((template) => template?.id)
    .map((template, index) => {
      const fallback = defaultTaskTemplates[index % defaultTaskTemplates.length];
      return {
        ...fallback,
        ...template,
        label: template.label || fallback.label,
        category: template.category || fallback.category,
        description: template.description ?? "",
        requirements: normalizeRequirements(template.requirements, fallback.requirements),
      };
    });
  return normalized.length ? normalized : defaultTaskTemplates;
}

function normalizeRequirements(requirements: Requirements | undefined, fallback: Requirements): Requirements {
  return {
    ...fallback,
    ...(requirements ?? {}),
    format: requirements?.format === "Markdown" || requirements?.format === "TXT" || requirements?.format === "DOCX" ? requirements.format : fallback.format,
    sections: requirements?.sections ?? fallback.sections,
  };
}

function normalizeMasterStatuses(statuses: typeof defaultMasterStatuses) {
  return {
    task: statuses?.task?.length ? statuses.task : defaultMasterStatuses.task,
    note: statuses?.note?.length ? statuses.note : defaultMasterStatuses.note,
  };
}

function normalizeOutputTemplateFormat(format: OutputTemplate["format"]): OutputTemplateFormat {
  return format === "TXT" || format === "DOCX" || format === "PDF" || format === "PPTX" ? format : "Markdown";
}

function normalizeOutputTemplateScope(scope: OutputTemplate["scope"]): OutputTemplate["scope"] {
  if (scope === "global") return "global";
  if (!Array.isArray(scope)) return "global";
  return scope.filter((item): item is ProjectId => Boolean(projects[item]));
}

function applyOutputTemplate(output: string, template: OutputTemplate) {
  const trimmed = cleanChatGptOutput(output);
  if (!trimmed) return "";

  return [
    `# ${template.brandName || template.name}`,
    trimmed,
  ]
    .filter(Boolean)
    .join("\n");
}

function cleanChatGptOutput(output: string) {
  return output
    .replace(/^```(?:markdown|md|text)?\s*/i, "")
    .replace(/```\s*$/i, "")
    .replace(/^\s*(?:Here is|Sure, here is|Below is).{0,120}:\s*/i, "")
    .replace(/^#+\s*(?:Template Slots|Content|Format target|Template purpose|Style rules)\s*$/gim, "")
    .replace(/^(?:Format target|Template purpose|Style rules):.*$/gim, "")
    .replace(/<\/?[^>]+>/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeWorkTask(task: WorkTask): WorkTask {
  const fallbackProjectId = task.projectId && projects[task.projectId] ? task.projectId : "avbob";
  const migratedTemplateId = legacyTemplateMap[task.templateId] ?? task.templateId;
  const fallbackTemplateId = projects[fallbackProjectId].tasks.some((template) => template.id === migratedTemplateId)
    ? migratedTemplateId
    : projects[fallbackProjectId].tasks[0].id;
  const savedOutputTemplate = defaultOutputTemplates.find((template) => template.id === task.outputTemplateId);
  const fallbackOutputTemplate =
    savedOutputTemplate && isOutputTemplateCompatibleWithTask(savedOutputTemplate, fallbackTemplateId)
      ? savedOutputTemplate
      : preferredOutputTemplateForTask(fallbackProjectId, fallbackTemplateId, defaultOutputTemplates) ?? defaultOutputTemplates[0];
  const migratedFromLegacyTemplate = migratedTemplateId !== task.templateId;

  return {
    ...task,
    projectId: fallbackProjectId,
    templateId: fallbackTemplateId,
    outputTemplateId: fallbackOutputTemplate.id,
    title: task.title || "Untitled task",
    details: task.details ?? "",
    category: task.category || "General",
    priority: task.priority || "Normal",
    dueDate: task.dueDate ?? "",
    reminderAt: task.reminderAt ?? "",
    recurrence: task.recurrence ?? "None",
    recurrenceNote: task.recurrenceNote ?? "",
    status: normalizeStatus(task.status),
    favorite: Boolean(task.favorite),
    statusHistory: task.statusHistory?.length
      ? task.statusHistory.map((entry) => ({ ...entry, status: normalizeStatus(entry.status) }))
      : [{ status: normalizeStatus(task.status), changedAt: task.updatedAt || task.createdAt || new Date().toISOString() }],
    checklist: task.checklist?.map((item) => ({ ...item, done: Boolean(item.done) })) ?? [],
    input: task.input ?? "",
    assets: task.assets ?? [],
    requirements: migratedFromLegacyTemplate
      ? requirementsLinkedToOutputTemplate(taskRequirements(fallbackProjectId, fallbackTemplateId), fallbackOutputTemplate)
      : requirementsLinkedToOutputTemplate(task.requirements ?? taskRequirements(fallbackProjectId, fallbackTemplateId), fallbackOutputTemplate),
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
    status: note.status === "Closed" ? "Closed" : "Active",
    pinned: Boolean(note.pinned),
    favorite: Boolean(note.favorite),
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

function taskRequirements(projectId: ProjectId, templateId: string, templates = projects[projectId]?.tasks ?? defaultTaskTemplates): Requirements {
  const template = templates.find((item) => item.id === templateId) ?? projects[projectId]?.tasks.find((item) => item.id === templateId);
  return { ...(template?.requirements ?? defaultRequirements) };
}

function requirementsLinkedToOutputTemplate(requirements: Requirements, outputTemplate: OutputTemplate): Requirements {
  return {
    ...requirements,
    format: outputTemplate.format === "PPTX" || outputTemplate.format === "PDF" ? "Markdown" : outputTemplate.format,
    sections: requirements.sections.trim() || outputTemplate.slots.join(", "),
  };
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
  const scheduled = openTasks
    .filter((item) => scheduleValue(item) !== Number.MAX_SAFE_INTEGER)
    .sort((a, b) => scheduleValue(a) - scheduleValue(b));
  const todayStart = startOfToday(now);
  const tomorrowStart = startOfTomorrow(now);
  const dayAfterTomorrowStart = startOfDayOffset(now, 2);

  return {
    overdue: scheduled.filter((item) => scheduleValue(item) < todayStart),
    today: scheduled.filter((item) => {
      const value = scheduleValue(item);
      return value >= todayStart && value < tomorrowStart;
    }),
    tomorrow: scheduled.filter((item) => {
      const value = scheduleValue(item);
      return value >= tomorrowStart && value < dayAfterTomorrowStart;
    }),
    upcoming: scheduled.filter((item) => scheduleValue(item) >= dayAfterTomorrowStart),
    planningQueue: openTasks
      .filter((item) => scheduleValue(item) === Number.MAX_SAFE_INTEGER)
      .sort((a, b) => priorityRank(b.priority) - priorityRank(a.priority) || dateValue(a.dueDate) - dateValue(b.dueDate)),
  };
}

function buildProjectDashboard(tasks: WorkTask[]) {
  return (Object.keys(projects) as ProjectId[]).map((projectId) => {
    const projectTasks = tasks.filter((task) => task.projectId === projectId);
    const summary = buildTaskSummary(projectTasks);
    return { projectId, favorite: projectTasks.filter((task) => task.favorite).length, ...summary };
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
      closed: projectTasks.filter((task) => task.status === "Closed").length,
      favorite: projectTasks.filter((task) => task.favorite).length,
      byStatus: Object.fromEntries(
        mobileStatusOrder.map((status) => [status, projectTasks.filter((task) => task.status === status)]),
      ) as Record<TaskStatus, WorkTask[]>,
    };
  });
}

function filterTasksByListMode(tasks: WorkTask[], filter: "Active" | "Favorites" | TaskStatus | "All") {
  if (filter === "All") return tasks;
  if (filter === "Active") return tasks.filter((task) => task.status !== "Closed");
  if (filter === "Favorites") return tasks.filter((task) => task.favorite);
  return tasks.filter((task) => task.status === filter);
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
  const value = scheduleValue(task);
  if (value === Number.MAX_SAFE_INTEGER) return "No reminder or due date";
  const dateText = new Date(value).toLocaleString();
  const source = isValidDateTime(task.reminderAt) ? "Reminder" : "Due date";
  if (value < startOfToday(now)) return `Overdue: ${dateText}`;
  if (value < startOfTomorrow(now)) return `${source} today: ${dateText}`;
  if (value < startOfDayOffset(now, 2)) return `${source} tomorrow: ${dateText}`;
  return `${source} upcoming: ${dateText}`;
}

function reminderState(task: WorkTask, now: number) {
  const value = scheduleValue(task);
  if (value === Number.MAX_SAFE_INTEGER) return "unscheduled";
  if (value < startOfToday(now)) return "overdue";
  if (value < startOfTomorrow(now)) return "today";
  if (value < startOfDayOffset(now, 2)) return "tomorrow";
  return "upcoming";
}

function scheduleValue(task: WorkTask) {
  if (isValidDateTime(task.reminderAt)) return dateValue(task.reminderAt);
  if (task.dueDate) return new Date(`${task.dueDate}T00:00:00`).getTime();
  return Number.MAX_SAFE_INTEGER;
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
  return startOfDayOffset(now, 1);
}

function startOfDayOffset(now: number, offsetDays: number) {
  const date = new Date(startOfToday(now));
  date.setDate(date.getDate() + offsetDays);
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

async function readPdfText(file: File) {
  const pdfjs = await import("pdfjs-dist");
  const workerUrl = new URL("pdfjs-dist/build/pdf.worker.mjs", import.meta.url).toString();
  pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

  const document = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pages: string[] = [];

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const pageText = textContent.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();

    if (pageText) pages.push(`Page ${pageNumber}\n${pageText}`);
  }

  const extracted = pages.join("\n\n").trim();
  return extracted || "The PDF file was read, but no selectable text was found. If it is scanned or image-only, paste the important text into the input box.";
}

async function extractTemplateReferenceText(file: File) {
  const isDocx = /\.docx$/i.test(file.name);
  const isPptx = /\.pptx$/i.test(file.name);
  const isPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);
  const isText = file.type.startsWith("text/") || /\.(md|txt|csv|json|html)$/i.test(file.name);
  const extracted = isDocx
    ? await readDocxText(file)
    : isPptx
      ? await readPptxText(file)
    : isPdf
      ? await readPdfText(file)
    : isText
      ? await file.text()
      : "";
  const placeholders = extractLikelyPlaceholders(extracted);
  return [
    extracted ? extracted.slice(0, 30_000) : "No readable text could be extracted from this maintained template source.",
    placeholders.length ? `\nDetected placeholders:\n${placeholders.map((item) => `- ${item}`).join("\n")}` : "",
  ]
    .filter(Boolean)
    .join("\n")
    .trim();
}

function extractLikelyPlaceholders(text: string) {
  const matches = text.match(/\{\{[^}]+\}\}|\[[^\]]+\]|<[A-Z][^>]+>/g) ?? [];
  return Array.from(new Set(matches.map((item) => item.trim()))).slice(0, 80);
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

function normalizeSectionList(sections: string) {
  return sections
    .split(/\n|,/)
    .map((section) => section.trim())
    .filter(Boolean)
    .join(", ");
}

function manualPromptFormatLabel(format: OutputTemplateFormat) {
  if (format === "DOCX") return "DOCX-ready content";
  if (format === "PDF") return "PDF-ready content";
  if (format === "PPTX") return "PPTX-ready slide content";
  if (format === "TXT") return "Plain text";
  return "Markdown";
}

function manualPromptExportInstruction(format: OutputTemplateFormat) {
  if (format === "DOCX") {
    return [
      "Formatting/export target: DOCX-ready Markdown.",
      "Return the final answer directly in this chat.",
      "Do not claim to create, attach, export or download a DOCX file.",
      "Structure the content so it can be copied into Word or exported by my app with minimal editing.",
    ].join("\n");
  }

  if (format === "PDF") return "Formatting/export target: PDF-ready Markdown. Return the final answer directly in this chat as text; the app is responsible for PDF export.";
  if (format === "PPTX") return "Formatting/export target: PPTX-ready slide content. Return slide-ready text or JSON directly in this chat; the app is responsible for PPTX export.";
  if (format === "TXT") return "Formatting/export target: Plain text. Return the final answer directly in this chat.";
  return "Formatting/export target: Markdown. Return the final answer directly in this chat.";
}

function buildFullLlmPrompt(
  projectName: string,
  projectContext: string,
  taskLabel: string,
  taskTitle: string,
  taskDetails: string,
  checklist: ChecklistItem[],
  input: string,
  assets: InputAsset[],
  requirements: Requirements,
  outputTemplate: OutputTemplate,
  inputQuality: InputQuality,
) {
  const readableFiles = assets.filter((asset) => asset.type === "text");
  const imageFiles = assets.filter((asset) => asset.type === "image");
  const unreadableFiles = assets.filter((asset) => asset.type === "file");
  const sourceLimitations = [
    imageFiles.length ? "Image attachments are listed by filename only. Use them only if the user also provides visual details in the typed input." : "",
    unreadableFiles.length ? "Some attachments cannot be read in this browser workflow. Do not invent their contents." : "",
    readableFiles.some((asset) => asset.content.includes("no selectable text was found"))
      ? "At least one PDF had no selectable text, which usually means it is scanned or image-only. Do not invent its contents."
      : "",
    assets.some((asset) => asset.content.includes("[Content truncated at")) ? "Some source text was truncated. Flag any important gaps before relying on missing material." : "",
  ].filter(Boolean);
  const exportInstruction = manualPromptExportInstruction(outputTemplate.format);
  const promptFormatLabel = manualPromptFormatLabel(outputTemplate.format);

  return [
    "You are a practical AI tasks and notes assistant. Process the full source material and create new output content.",
    "I am using ChatGPT Plus manually, so produce the final answer directly in this chat.",
    "",
    "INTENT BRIEF",
    `Objective: Create ${requirements.outputType} for "${taskTitle}".`,
    `Desired outcome: A usable final output that can be pasted into the selected output template or exported with minimal editing.`,
    `Audience: ${requirements.audience}`,
    `Success criteria: Follow the required sections, use the selected output template slots, respect the tone and constraints, and clearly distinguish facts from assumptions where relevant.`,
    `Source material included: Typed input${readableFiles.length ? `, ${readableFiles.length} readable file(s)` : ""}${checklist.length ? ", task checklist" : ""}.`,
    `Source limitations: ${sourceLimitations.length ? sourceLimitations.join(" ") : "No major source limitations detected."}`,
    "Assumptions policy: State important assumptions explicitly. Do not invent missing facts, figures, dates, commitments, or document contents.",
    `Output structure: Use the task sections and map the content into the output template slots named below.`,
    exportInstruction,
    "",
    "PROJECT",
    `Name: ${projectName}`,
    `Context: ${projectContext}`,
    `Task template: ${taskLabel}`,
    `Task title: ${taskTitle}`,
    `Task details: ${taskDetails.trim() || "No task details provided."}`,
    "",
    "CONTENT BRIEF",
    `Content goal: ${requirements.outputType}`,
    `Linked export format: ${promptFormatLabel}`,
    `Tone: ${requirements.tone}`,
    `Audience: ${requirements.audience}`,
    `Length: ${requirements.length}`,
    `Required sections: ${normalizeSectionList(requirements.sections)}`,
    `Constraints: ${requirements.constraints || "None"}`,
    `Image requirements: ${requirements.imageRequirements || "None"}`,
    "",
    "OUTPUT TEMPLATE",
    `Name: ${outputTemplate.name}`,
    `Description: ${outputTemplate.description || "None"}`,
    `Group: ${outputTemplate.group}`,
    `Preferred format: ${promptFormatLabel}`,
    `Maintained source template: ${outputTemplate.sourceTemplateName || "None uploaded"}`,
    `Compatible task IDs: ${outputTemplate.compatibleTaskIds.length ? outputTemplate.compatibleTaskIds.join(", ") : "Any"}`,
    `Output template slots: ${outputTemplate.slots.join(", ")}`,
    `Style rules: ${outputTemplate.style || "None"}`,
    outputTemplate.sourceTemplateText
      ? `Extracted source template text and placeholders:\n${outputTemplate.sourceTemplateText}`
      : "Extracted source template text and placeholders: None. Use the output template slots and style rules.",
    "",
    "INPUT QUALITY CHECK",
    `Status: ${inputQuality.status}`,
    inputQuality.checks.length ? `Checks: ${inputQuality.checks.join(" ")}` : "Checks: None",
    inputQuality.warnings.length ? `Warnings: ${inputQuality.warnings.join(" ")}` : "Warnings: None",
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
    "Create the requested output as business-ready text according to the selected output type.",
    "The AI prompt should only request text content. The app is responsible for turning that content into DOCX, PDF, or PPTX files.",
    "Use headings that match the required sections from the content brief.",
    "Map the answer to the output template slots and any placeholders found in the maintained source template.",
    outputTemplate.format === "PPTX"
      ? "For presentations, return strict JSON only with this shape: { \"slides\": [{ \"id\": \"slide-1\", \"title\": \"\", \"layout\": \"title|section|content|two-column|image-left|image-right|image-full\", \"background\": \"#ffffff\", \"keyMessage\": \"\", \"bullets\": [], \"speakerNotes\": \"\", \"visualDirection\": \"\", \"images\": [{ \"label\": \"\", \"source\": \"matching uploaded file name or visual source\", \"alt\": \"\" }] }] }. Do not wrap the JSON in code fences."
      : "Return clean Markdown only. Do not wrap the answer in code fences. Do not include XML, HTML, template metadata, or commentary about how the answer was created.",
    "For image needs, presentations must use the images array; other outputs may use explicit placeholders in this format: [IMAGE PLACEHOLDER: short description | suggested source or visual direction].",
    "If the selected output template is AI prompt generator, return a clean reusable prompt block with no surrounding commentary.",
    "If the requested output is an email, include a usable subject line and email body.",
    "If the requested output is a summary, distinguish confirmed information, assumptions, risks, gaps, and action items.",
    "If the task checklist contains open items, include them as action items or a tickable checklist when relevant.",
    "Do not summarize the prompt. Use all relevant source material. Ask clarifying questions only if required information is genuinely missing.",
  ].join("\n");
}

type MarkdownBlock = {
  type: "heading" | "bullet" | "numbered" | "paragraph";
  level: number;
  text: string;
};

type SlideLayout = "title" | "section" | "content" | "two-column" | "image-left" | "image-right" | "image-full";

type SlideImageModel = {
  label: string;
  source: string;
  alt: string;
};

type SlideModel = {
  id: string;
  title: string;
  layout: SlideLayout;
  background: string;
  keyMessage: string;
  bullets: string[];
  speakerNotes: string;
  visualDirection: string;
  images: SlideImageModel[];
  imageAssets: InputAsset[];
  raw: string;
};

type SlideDeckBuild = {
  slides: SlideModel[];
  error: string;
};

function parseMarkdownBlocks(markdown: string): MarkdownBlock[] {
  return markdown
    .split("\n")
    .map((line) => {
      const trimmed = line.trim();
      if (!trimmed) return { type: "paragraph" as const, level: 0, text: "" };
      const headingMatch = trimmed.match(/^(#{1,4})\s+(.+)$/);
      if (headingMatch) return { type: "heading" as const, level: headingMatch[1].length, text: cleanInlineMarkdown(headingMatch[2]) };
      if (/^[-*]\s+/.test(trimmed)) return { type: "bullet" as const, level: 0, text: cleanInlineMarkdown(trimmed.replace(/^[-*]\s+/, "")) };
      if (/^\d+\.\s+/.test(trimmed)) return { type: "numbered" as const, level: 0, text: cleanInlineMarkdown(trimmed.replace(/^\d+\.\s+/, "")) };
      return { type: "paragraph" as const, level: 0, text: cleanInlineMarkdown(trimmed.replace(/^\*\*([^*]+):\*\*\s*/, "$1: ")) };
    });
}

function cleanInlineMarkdown(text: string) {
  return text
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .trim();
}

function dataUrlToUint8Array(dataUrl: string) {
  const base64 = dataUrl.split(",")[1] ?? "";
  const binary = atob(base64);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

function brandedTitle(template: OutputTemplate) {
  return template.brandName || template.name;
}

async function toDocxBlob(markdown: string, template: OutputTemplate) {
  const blocks = parseMarkdownBlocks(markdown);
  const children = markdown.split("\n").map((line) => {
    const trimmed = line.trim();
    const headingMatch = trimmed.match(/^(#{1,3})\s+(.+)$/);
    const isBullet = /^[-*]\s+/.test(trimmed);
    const isNumbered = /^\d+\.\s+/.test(trimmed);
    const text = trimmed
      .replace(/^#{1,6}\s*/, "")
      .replace(/^[-*]\s+/, "")
      .replace(/^\d+\.\s+/, "")
      .replace(/^\*\*([^*]+):\*\*\s*/, "$1: ");

    if (!trimmed) {
      return new Paragraph({ spacing: { after: 80 } });
    }

    return new Paragraph({
      heading: headingMatch?.[1].length === 1 ? HeadingLevel.HEADING_1 : headingMatch?.[1].length === 2 ? HeadingLevel.HEADING_2 : headingMatch ? HeadingLevel.HEADING_3 : undefined,
      bullet: isBullet ? { level: 0 } : undefined,
      numbering: isNumbered ? { reference: "ordered-list", level: 0 } : undefined,
      spacing: { after: headingMatch ? 180 : 90 },
      children: [
        new TextRun({
          text,
          bold: Boolean(headingMatch) || /^\w[\w\s]+:/.test(text),
          size: headingMatch ? 30 : 22,
        }),
      ],
    });
  });
  const headerChildren = [
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({
          text: brandedTitle(template),
          bold: true,
          color: template.primaryColor,
        }),
      ],
    }),
  ];

  const document = new Document({
    numbering: {
      config: [
        {
          reference: "ordered-list",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: "left",
            },
          ],
        },
      ],
    },
    sections: [
      {
        headers: { default: new Header({ children: headerChildren }) },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: brandedTitle(template), color: template.primaryColor, size: 18 })],
              }),
            ],
          }),
        },
        children,
      },
    ],
  });

  return Packer.toBlob(document);
}

async function toPdfBlob(markdown: string, template: OutputTemplate) {
  const { jsPDF } = await import("jspdf");
  const pdf = new jsPDF({ unit: "pt", format: "a4" });
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 48;
  let y = 56;

  pdf.setFillColor(`#${template.primaryColor}`);
  pdf.rect(0, 0, pageWidth, 14, "F");
  if (template.logoDataUrl) {
    pdf.addImage(template.logoDataUrl, "PNG", pageWidth - 138, 26, 90, 34);
  } else {
    pdf.setTextColor(`#${template.primaryColor}`);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(12);
    pdf.text(brandedTitle(template), pageWidth - margin, 36, { align: "right" });
  }

  for (const block of parseMarkdownBlocks(markdown)) {
    if (!block.text) {
      y += 8;
      continue;
    }
    const isHeading = block.type === "heading";
    const fontSize = isHeading ? (block.level === 1 ? 18 : 14) : 10.5;
    const lineHeight = isHeading ? fontSize + 8 : fontSize + 5;
    pdf.setFont("helvetica", isHeading ? "bold" : "normal");
    pdf.setFontSize(fontSize);
    pdf.setTextColor(isHeading ? `#${template.primaryColor}` : "#171827");
    const prefix = block.type === "bullet" ? "- " : "";
    const lines = pdf.splitTextToSize(`${prefix}${block.text}`, pageWidth - margin * 2);

    if (y + lines.length * lineHeight > pageHeight - 54) {
      pdf.addPage();
      pdf.setFillColor(`#${template.primaryColor}`);
      pdf.rect(0, 0, pageWidth, 14, "F");
      y = 48;
    }

    pdf.text(lines, margin, y);
    y += lines.length * lineHeight + (isHeading ? 6 : 2);
  }

  const pageCount = pdf.getNumberOfPages();
  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    pdf.setPage(pageNumber);
    pdf.setFontSize(8);
    pdf.setTextColor("#627168");
    pdf.text(`${brandedTitle(template)} | ${pageNumber}`, pageWidth - margin, pageHeight - 24, { align: "right" });
  }

  return pdf.output("blob");
}

function splitMarkdownIntoSlides(markdown: string) {
  const explicitSlides = markdown
    .split(/\n(?=#{1,3}\s+(?:Slide\s+\d+|Title slide|Agenda|Context|Recommendation|Closing action)\b)/i)
    .map((item) => item.trim())
    .filter(Boolean);
  if (explicitSlides.length > 1) return explicitSlides;

  const blocks = parseMarkdownBlocks(markdown).filter((block) => block.text);
  const chunks: string[] = [];
  for (let index = 0; index < blocks.length; index += 7) {
    chunks.push(blocks.slice(index, index + 7).map((block) => `${block.type === "heading" ? "#" : "-"} ${block.text}`).join("\n"));
  }
  return chunks.length ? chunks : [markdown];
}

function extractSlideImagePlaceholders(text: string) {
  const labels = new Set<string>();
  const markdownImage = /!\[([^\]]*)\]\(([^)]+)\)/g;
  const namedImage = /\[\[image:\s*([^\]]+)\]\]/gi;
  const placeholderImage = /\[IMAGE PLACEHOLDER:\s*([^\]]+)\]/gi;
  let match: RegExpExecArray | null;

  while ((match = markdownImage.exec(text))) {
    labels.add(match[1] || match[2]);
  }
  while ((match = namedImage.exec(text))) {
    labels.add(match[1]);
  }
  while ((match = placeholderImage.exec(text))) {
    labels.add(match[1]);
  }

  return Array.from(labels).map((item) => item.trim()).filter(Boolean);
}

function matchAssetForPlaceholder(label: string, assets: InputAsset[]) {
  const normalizedLabel = label.toLowerCase().trim();
  if (!normalizedLabel) return undefined;
  return assets.find((asset) =>
    asset.type === "image" &&
    [asset.name, asset.name.replace(/[-_]/g, " ")].some((assetName) =>
      assetName.toLowerCase().includes(normalizedLabel) || normalizedLabel.includes(assetName.toLowerCase()),
    ),
  );
}

function normalizeSlideLayout(value: unknown, index: number): SlideLayout {
  const normalized = String(value ?? "").toLowerCase().trim();
  const validLayouts: SlideLayout[] = ["title", "section", "content", "two-column", "image-left", "image-right", "image-full"];
  if (validLayouts.includes(normalized as SlideLayout)) return normalized as SlideLayout;
  if (index === 0) return "title";
  return "content";
}

function normalizeBackground(value: unknown, template: OutputTemplate) {
  const background = String(value ?? "").trim();
  if (/^#[0-9a-f]{6}$/i.test(background)) return background;
  if (/^[0-9a-f]{6}$/i.test(background)) return `#${background}`;
  return `#${template.secondaryColor || "ffffff"}`;
}

function normalizeStringList(value: unknown) {
  if (Array.isArray(value)) return value.map((item) => String(item).trim()).filter(Boolean);
  if (typeof value === "string") return value.split(/\n|;/).map((item) => item.trim()).filter(Boolean);
  return [];
}

function normalizeSlideImages(value: unknown, visualDirection: string) {
  const items = Array.isArray(value) ? value : value ? [value] : [];
  const images = items
    .map((item): SlideImageModel => {
      if (typeof item === "string") return { label: item.trim(), source: "", alt: item.trim() };
      const record = item && typeof item === "object" ? item as Record<string, unknown> : {};
      const label = String(record.label ?? record.description ?? record.name ?? record.placeholder ?? "").trim();
      const source = String(record.source ?? record.src ?? record.url ?? "").trim();
      const alt = String(record.alt ?? label).trim();
      return { label, source, alt };
    })
    .filter((item) => item.label || item.source);

  if (!images.length && visualDirection) {
    return [{ label: visualDirection, source: "", alt: visualDirection }];
  }

  return images;
}

function serializeSlideForJson(slide: SlideModel) {
  return {
    id: slide.id,
    title: slide.title,
    layout: slide.layout,
    background: slide.background,
    keyMessage: slide.keyMessage,
    bullets: slide.bullets,
    speakerNotes: slide.speakerNotes,
    visualDirection: slide.visualDirection,
    images: slide.images,
  };
}

function formatSlideDeckJson(slides: SlideModel[]) {
  return JSON.stringify({ slides: slides.map(serializeSlideForJson) }, null, 2);
}

function createFallbackSlide(index: number, template: OutputTemplate): SlideModel {
  return normalizeSlideModel(
    {
      id: `slide-${index + 1}`,
      title: index === 0 ? "Presentation title" : `Slide ${index + 1}`,
      layout: index === 0 ? "title" : "content",
      background: `#${template.secondaryColor || "ffffff"}`,
      keyMessage: "",
      bullets: [],
      speakerNotes: "",
      visualDirection: "",
      images: [],
    },
    index,
    [],
    template,
  );
}

function normalizeSlideModel(value: unknown, index: number, assets: InputAsset[], template: OutputTemplate): SlideModel {
  const record = value && typeof value === "object" ? value as Record<string, unknown> : {};
  const visualDirection = String(record.visualDirection ?? record.visual_direction ?? record.visual ?? "").trim();
  const images = normalizeSlideImages(record.images ?? record.image ?? record.imagePlaceholders, visualDirection);
  const imageAssets = images
    .map((image) => matchAssetForPlaceholder(image.source || image.label, assets))
    .filter(Boolean) as InputAsset[];

  return {
    id: String(record.id ?? `slide-${index + 1}`).trim() || `slide-${index + 1}`,
    title: String(record.title ?? record.slideTitle ?? `Slide ${index + 1}`).trim() || `Slide ${index + 1}`,
    layout: normalizeSlideLayout(record.layout, index),
    background: normalizeBackground(record.background, template),
    keyMessage: String(record.keyMessage ?? record.key_message ?? record.message ?? "").trim(),
    bullets: normalizeStringList(record.bullets ?? record.bodyLines ?? record.body),
    speakerNotes: String(record.speakerNotes ?? record.speaker_notes ?? record.notes ?? "").trim(),
    visualDirection,
    images,
    imageAssets,
    raw: JSON.stringify(record, null, 2),
  };
}

function parseSlideDeckJson(input: string, assets: InputAsset[], template: OutputTemplate): SlideDeckBuild {
  const cleaned = cleanChatGptOutput(input).replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/i, "").trim();
  if (!cleaned) return { slides: [], error: "" };

  try {
    const parsed = JSON.parse(cleaned);
    const slideValues = Array.isArray(parsed) ? parsed : Array.isArray(parsed?.slides) ? parsed.slides : [];
    if (!slideValues.length) {
      return { slides: [], error: "JSON parsed, but no slides array was found." };
    }
    return {
      slides: slideValues.map((slide: unknown, index: number) => normalizeSlideModel(slide, index, assets, template)),
      error: "",
    };
  } catch (error) {
    return {
      slides: [],
      error: error instanceof Error ? `JSON is invalid: ${error.message}` : "JSON is invalid.",
    };
  }
}

function markdownSlideToModel(slideContent: string, index: number, assets: InputAsset[], template: OutputTemplate): SlideModel {
    const blocks = parseMarkdownBlocks(slideContent).filter((block) => block.text);
    const title = blocks.find((block) => block.type === "heading")?.text ?? `Slide ${index + 1}`;
    const notesBlock = blocks.find((block) => /^speaker notes?:/i.test(block.text) || /^notes?:/i.test(block.text));
    const notes = notesBlock ? notesBlock.text.replace(/^(speaker notes?:|notes?:)\s*/i, "") : "";
    const imageLabels = extractSlideImagePlaceholders(slideContent);
    const imageAssets = imageLabels.map((label) => matchAssetForPlaceholder(label, assets)).filter(Boolean) as InputAsset[];
    const bodyLines = blocks
      .filter(
        (block) =>
          block.type !== "heading" &&
          !/^speaker notes?:/i.test(block.text) &&
          !/^notes?:/i.test(block.text),
      )
      .map((block) => block.text);

    return {
      id: `slide-${index + 1}`,
      title,
      layout: index === 0 ? "title" : imageAssets.length || imageLabels.length ? "image-right" : "content",
      background: `#${template.secondaryColor || "ffffff"}`,
      keyMessage: bodyLines.find((line) => /^key message:/i.test(line))?.replace(/^key message:\s*/i, "") ?? "",
      bullets: bodyLines.filter((line) => !/^key message:/i.test(line) && !/^visual direction:/i.test(line)),
      speakerNotes: notes,
      visualDirection: bodyLines.find((line) => /^visual direction:/i.test(line))?.replace(/^visual direction:\s*/i, "") ?? "",
      images: imageLabels.map((label) => ({ label, source: "", alt: label })),
      imageAssets,
      raw: slideContent,
    };
}

function buildSlideDeck(input: string, assets: InputAsset[], template: OutputTemplate): SlideDeckBuild {
  const trimmed = input.trim();
  if (!trimmed) return { slides: [], error: "" };
  if (/^\s*(?:```json\s*)?[\[{]/i.test(trimmed)) return parseSlideDeckJson(trimmed, assets, template);

  return {
    slides: splitMarkdownIntoSlides(trimmed).map((slideContent, index) => markdownSlideToModel(slideContent, index, assets, template)),
    error: "This is Markdown, so the preview is using a best-effort conversion. Normalize it to JSON before managing the deck.",
  };
}

async function toPptxBlob(markdown: string, template: OutputTemplate, assets: InputAsset[]) {
  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = brandedTitle(template);
  pptx.subject = template.name;
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
  };

  const slides = buildSlideDeck(markdown, assets, template).slides;
  slides.forEach((slidePreview, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: slidePreview.background.replace("#", "") || "FFFFFF" };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.12, fill: { color: template.primaryColor }, line: { color: template.primaryColor } });
    if (template.logoDataUrl) {
      slide.addImage({ data: template.logoDataUrl, x: 11.2, y: 0.25, w: 1.55, h: 0.55 });
    } else {
      slide.addText(brandedTitle(template), { x: 10.25, y: 0.3, w: 2.5, h: 0.25, fontSize: 8, bold: true, color: template.primaryColor, align: "right" });
    }

    slide.addText(slidePreview.title, {
      x: 0.65,
      y: 0.65,
      w: 11.2,
      h: 0.55,
      fontSize: 28,
      bold: true,
      color: template.primaryColor,
      margin: 0,
    });
    slide.addShape(pptx.ShapeType.line, { x: 0.65, y: 1.28, w: 2.2, h: 0, line: { color: template.accentColor, width: 2 } });

    const hasImage = slidePreview.imageAssets.length > 0;
    const imageLeft = slidePreview.layout === "image-left";
    const imageFull = slidePreview.layout === "image-full";
    const textWidth = hasImage && !imageFull ? 6.8 : 11.65;
    const textX = hasImage && imageLeft ? 5.75 : 0.8;
    const textY = 1.65;
    const textHeight = 5.1;

    const textLines = [
      slidePreview.keyMessage,
      ...slidePreview.bullets,
    ].filter(Boolean);

    if (textLines.length > 0 && !imageFull) {
      const bulletLines = textLines.map((line) => ({ text: line, options: { bullet: { type: "bullet" as const } } }));
      slide.addText(bulletLines, {
        x: textX,
        y: textY,
        w: textWidth,
        h: textHeight,
        fontSize: 16,
        color: "171827",
        breakLine: false,
        fit: "shrink",
        valign: "top",
      });
    }

    if (hasImage) {
      slidePreview.imageAssets.slice(0, 2).forEach((asset, assetIndex) => {
        const imageX = imageFull ? 0.85 : imageLeft ? 0.8 : 8.1;
        const imageY = imageFull ? 1.5 : 1.5 + assetIndex * 3.4;
        const imageW = imageFull ? 11.65 : 4.5;
        const imageH = imageFull ? 5.05 : 3.25;
        slide.addImage({ data: asset.content, x: imageX, y: imageY, w: imageW, h: imageH });
      });
    }

    if (slidePreview.speakerNotes) {
      const notesText = slidePreview.speakerNotes.trim();
      if (notesText && typeof (slide as any).addNotes === "function") {
        (slide as any).addNotes(notesText);
      }
    }

    slide.addText(`${brandedTitle(template)} | ${index + 1}`, { x: 0.65, y: 7.05, w: 12, h: 0.22, fontSize: 7, color: "627168", align: "right" });
  });

  return (await pptx.write({ outputType: "blob" })) as Blob;
}

export { App };
