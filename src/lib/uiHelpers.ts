import type { TaskStatus, WorkTask } from "../types";

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

