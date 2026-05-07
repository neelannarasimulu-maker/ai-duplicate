import { BellRing, CalendarClock, ListTodo, Save } from "lucide-react";
import type { ProjectId, WorkTask } from "../types";
import { reminderLabel, reminderState } from "../lib/uiHelpers";

interface ReminderPlannerData {
  overdue: WorkTask[];
  today: WorkTask[];
  tomorrow: WorkTask[];
  upcoming: WorkTask[];
  planningQueue: WorkTask[];
}

interface RemindersViewProps {
  reminderPlanner: ReminderPlannerData;
  projects: Record<ProjectId, { name: string }>;
  now: number;
  notificationPermission: NotificationPermission | "unsupported";
  enableNotifications: () => void;
  openWorkTask: (task: WorkTask) => void;
  clearReminder: (id: string) => void;
  saveReminder: (id: string) => void;
  updateReminder: (id: string, value: string) => void;
  reminderValue: (task: WorkTask) => string;
  scheduleReminderToday: (id: string) => void;
  setViewMode: (mode: "home" | "work" | "reminders" | "notes" | "mobile") => void;
}

function ReminderColumn({
  items,
  title,
  emptyText,
  now,
  projects,
  onOpen,
  onClear,
  onSave,
  onSchedule,
  reminderValue,
}: {
  items: WorkTask[];
  title: string;
  emptyText: string;
  now: number;
  projects: Record<ProjectId, { name: string }>;
  onOpen: (task: WorkTask) => void;
  onClear: (id: string) => void;
  onSave: (id: string) => void;
  onSchedule: (id: string, value: string) => void;
  reminderValue: (task: WorkTask) => string;
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

export function RemindersView({
  reminderPlanner,
  projects,
  now,
  notificationPermission,
  enableNotifications,
  openWorkTask,
  clearReminder,
  saveReminder,
  updateReminder,
  reminderValue,
  scheduleReminderToday,
  setViewMode,
}: RemindersViewProps) {
  return (
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
            onClick={enableNotifications}
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
            projects={projects}
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
            projects={projects}
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
            projects={projects}
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
            projects={projects}
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
  );
}
