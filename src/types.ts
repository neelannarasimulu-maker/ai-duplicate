export type ProjectId = "avbob" | "naha" | "personal" | "supplysync360" | "bma" | "thenga";
export type Format = "Markdown" | "TXT" | "DOCX";
export type Priority = "Low" | "Normal" | "High" | "Urgent";
export type TaskStatus = "Open" | "In Progress" | "Blocked" | "Closed";

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

export type WorkTask = {
  id: string;
  projectId: ProjectId;
  templateId: string;
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
  title: string;
  createdAt: string;
  input: string;
  requirements: Requirements;
  result: string;
};
