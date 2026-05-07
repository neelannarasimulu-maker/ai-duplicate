import type { InputAsset, ProjectId, TaskStatus, WorkTask } from "../types";

export function isValidDateTime(value: string) {
  return Boolean(value && Number.isFinite(new Date(value).getTime()));
}

export function startOfToday(now: number) {
  const date = new Date(now);
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

export function startOfDayOffset(now: number, offsetDays: number) {
  const date = new Date(startOfToday(now));
  date.setDate(date.getDate() + offsetDays);
  return date.getTime();
}

export function startOfTomorrow(now: number) {
  return startOfDayOffset(now, 1);
}

export function isOverdue(task: WorkTask) {
  if (!task.dueDate || task.status === "Closed") return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(task.dueDate).getTime() < today.getTime();
}

export function isReminderDue(task: WorkTask) {
  return Boolean(isValidDateTime(task.reminderAt) && task.status !== "Closed" && new Date(task.reminderAt).getTime() <= Date.now());
}

export function scheduleValue(task: WorkTask) {
  if (isValidDateTime(task.reminderAt)) return new Date(task.reminderAt).getTime();
  if (task.dueDate) return new Date(`${task.dueDate}T00:00:00`).getTime();
  return Number.MAX_SAFE_INTEGER;
}

export function reminderLabel(task: WorkTask, now: number) {
  const value = scheduleValue(task);
  if (value === Number.MAX_SAFE_INTEGER) return "No reminder or due date";
  const dateText = new Date(value).toLocaleString();
  const source = isValidDateTime(task.reminderAt) ? "Reminder" : "Due date";
  if (value < startOfToday(now)) return `Overdue: ${dateText}`;
  if (value < startOfTomorrow(now)) return `${source} today: ${dateText}`;
  if (value < startOfDayOffset(now, 2)) return `${source} tomorrow: ${dateText}`;
  return `${source} upcoming: ${dateText}`;
}

export function reminderState(task: WorkTask, now: number) {
  const value = scheduleValue(task);
  if (value === Number.MAX_SAFE_INTEGER) return "unscheduled";
  if (value < startOfToday(now)) return "overdue";
  if (value < startOfTomorrow(now)) return "today";
  if (value < startOfDayOffset(now, 2)) return "tomorrow";
  return "upcoming";
}

export function statusSlug(status: TaskStatus) {
  return status.toLowerCase().replace(/\s+/g, "-");
}

export function taskClassName(task: WorkTask, activeWorkTaskId: string) {
  const classes = ["work-task", `status-${statusSlug(task.status)}`, `priority-${task.priority.toLowerCase()}`];
  if (activeWorkTaskId === task.id) classes.push("active");
  if (isOverdue(task)) classes.push("overdue");
  if (isReminderDue(task)) classes.push("reminder-due");
  return classes.join(" ");
}

export function taskDateLabel(task: WorkTask) {
  if (isOverdue(task)) return `Overdue: ${task.dueDate}`;
  if (task.dueDate) return `Due: ${task.dueDate}`;
  if (isReminderDue(task)) return "Reminder due";
  return "No due date";
}

export function missingDetailsForTask(task: WorkTask) {
  const missing: string[] = [];
  if (task.input.trim().length < 20 && task.assets.length === 0) missing.push("Add source input, notes, or a readable file.");
  if (!task.requirements.outputType.trim()) missing.push("Choose the output type you need.");
  if (!task.requirements.audience.trim()) missing.push("Describe who the output is for.");
  if (!task.requirements.sections.trim()) missing.push("List the required sections or structure.");
  return missing;
}

export function buildInputQuality(missingDetails: string[], assets: InputAsset[]) {
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
  } as const;
}

export function taskInputQualityStatus(task: WorkTask) {
  return `AI input: ${buildInputQuality(missingDetailsForTask(task), task.assets).status.toLowerCase()}`;
}
