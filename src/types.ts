export type ProjectId = "avbob" | "naha" | "personal" | "supplysync360" | "bma" | "thenga";
export type Format = "Markdown" | "TXT" | "DOCX";
export type OutputTemplateFormat = Format | "PDF" | "PPTX";
export type Priority = "Low" | "Normal" | "High" | "Urgent";
export type TaskStatus = "Open" | "In Progress" | "Blocked" | "To Do Later" | "Closed";

export type InputAsset = {
  id: string;
  name: string;
  type: "text" | "image" | "file";
  content: string;
};

export type Requirements = {
  outputType: string;
  format: Format;
  tone: string;
  audience: string;
  length: string;
  sections: string;
  constraints: string;
  imageRequirements: string;
};

export type ChecklistItem = {
  id: string;
  text: string;
  done: boolean;
};

export type OutputTemplate = {
  id: string;
  name: string;
  description: string;
  group: string;
  format: OutputTemplateFormat;
  brandName: string;
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  logoDataUrl: string;
  sourceTemplateName: string;
  sourceTemplateType: string;
  sourceTemplateDataUrl: string;
  sourceTemplateText: string;
  scope: "global" | ProjectId[];
  compatibleTaskIds: string[];
  slots: string[];
  style: string;
  createdAt: string;
  updatedAt: string;
};

export type WorkTask = {
  id: string;
  projectId: ProjectId;
  templateId: string;
  outputTemplateId?: string;
  title: string;
  details: string;
  category: string;
  priority: Priority;
  dueDate: string;
  reminderAt: string;
  status: TaskStatus;
  statusHistory: Array<{
    status: TaskStatus;
    changedAt: string;
  }>;
  checklist: ChecklistItem[];
  input: string;
  assets: InputAsset[];
  requirements: Requirements;
  gptPrompt: string;
  result: string;
  createdAt: string;
  updatedAt: string;
};

export type SavedOutput = {
  id: string;
  projectId: ProjectId;
  workTaskId: string;
  taskId: string;
  outputTemplateId?: string;
  title: string;
  createdAt: string;
  input: string;
  requirements: Requirements;
  result: string;
  renderedOutput?: string;
};

export type AppNoteEntry = {
  id: string;
  content: string;
  createdAt: string;
  updatedAt: string;
};

export type AppNote = {
  id: string;
  projectId: ProjectId;
  title: string;
  content?: string;
  entries: AppNoteEntry[];
  pinned: boolean;
  createdAt: string;
  updatedAt: string;
};
