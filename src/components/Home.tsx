import { BarChart3, CalendarClock, ListTodo, Sparkles, Star, StickyNote, WandSparkles } from "lucide-react";
import type { AppNote, WorkTask, ProjectId } from "../types";

interface HomeProps {
  activeWorkTaskId: string;
  favoriteNotes: AppNote[];
  favoriteTasks: WorkTask[];
  noteEntryCount: number;
  now: number;
  openAiStudio: () => void;
  openFavoriteTask: (task: WorkTask) => void;
  openHomeProjectDashboard: () => void;
  openNote: (note: AppNote) => void;
  openWorkTask: (task: WorkTask) => void;
  projects: Record<string, { name: string; tasks: any[] }>;
  recentNotes: AppNote[];
  reminderPlanner: {
    overdue: WorkTask[];
    today: WorkTask[];
    tomorrow: WorkTask[];
  };
  savedOutputs: any[];
  setViewMode: (mode: "home" | "work" | "reminders" | "notes" | "mobile") => void;
  summary: {
    inProgress: number;
    total: number;
    open: number;
    blocked: number;
    toDoLater: number;
    urgent: number;
    reminders: number;
    closed: number;
  };
  toggleTaskFavorite: (task: WorkTask) => void;
  activeFocusTasks: WorkTask[];
  projectDashboard: any[];
  projectId: string;
  updateProject: (id: ProjectId) => void;
}

function reminderState(task: WorkTask, now: number) {
  const due = task.dueDate ? new Date(task.dueDate) : null;
  if (!due) return "";
  const today = new Date(now);
  today.setHours(0, 0, 0, 0);
  const taskDate = new Date(due);
  taskDate.setHours(0, 0, 0, 0);
  if (taskDate < today) return "overdue";
  if (taskDate.getTime() === today.getTime()) return "today";
  return "";
}

function reminderLabel(task: WorkTask, now: number) {
  const due = task.dueDate ? new Date(task.dueDate) : null;
  if (!due) return task.reminderAt ? "reminder" : "no due date";
  const today = new Date(now);
  today.setHours(0, 0, 0, 0);
  const taskDate = new Date(due);
  taskDate.setHours(0, 0, 0, 0);
  const diffDays = Math.floor((taskDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (diffDays < 0) return "overdue";
  if (diffDays === 0) return "due today";
  if (diffDays === 1) return "due tomorrow";
  return `due in ${diffDays} days`;
}

function TaskListRow({ task, projectLabel, onOpen, onToggleFavorite, activeWorkTaskId }: {
  task: WorkTask;
  projectLabel: string;
  onOpen: (task: WorkTask) => void;
  onToggleFavorite: (task: WorkTask) => void;
  activeWorkTaskId: string;
}) {
  return (
    <article className={`task-row ${activeWorkTaskId === task.id ? "active" : ""}`}>
      <button className="task-title" onClick={() => onOpen(task)} type="button">
        <strong>{task.title}</strong>
        <span>{projectLabel} - {task.category} - {task.priority}</span>
      </button>
      <div className="task-meta">
        <span className={`status-pill status-${task.status.toLowerCase().replace(" ", "")}`}>{task.status}</span>
        <button
          className={task.favorite ? "ghost-button favorite-toggle active" : "ghost-button favorite-toggle"}
          onClick={() => onToggleFavorite(task)}
          type="button"
        >
          <Star size={14} />
        </button>
      </div>
    </article>
  );
}

export function Home(props: HomeProps) {
  const {
    activeWorkTaskId,
    favoriteNotes,
    favoriteTasks,
    noteEntryCount,
    now,
    openAiStudio,
    openFavoriteTask,
    openHomeProjectDashboard,
    openNote,
    openWorkTask,
    projects,
    recentNotes,
    reminderPlanner,
    savedOutputs,
    setViewMode,
    summary,
    toggleTaskFavorite,
    activeFocusTasks,
    projectDashboard,
    projectId,
    updateProject,
  } = props;

  return (
    <section className="home-page">
      <section className="hero-panel">
        <div>
          <p className="eyebrow">Live task workspace</p>
          <h2>Plan, capture, and draft with confidence</h2>
          <p className="context-copy">
            A calm operating view for urgent tasks, reminders, notes, project status, and AI-ready drafting.
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

      {(favoriteTasks.length > 0 || favoriteNotes.length > 0) && (
        <section className="panel favorites-panel">
          <div className="result-header">
            <h2>
              <Star size={18} />
              Favorites
            </h2>
            <button className="ghost-button" onClick={() => setViewMode("work")} type="button">
              Open tasks
            </button>
          </div>
          <div className="favorite-grid">
            {favoriteTasks.map((item) => (
              <button className="favorite-card" key={item.id} onClick={() => openFavoriteTask(item)} type="button">
                <Star size={15} />
                <span>
                  <strong>{item.title}</strong>
                  <small>{projects[item.projectId].name} - task details</small>
                </span>
              </button>
            ))}
            {favoriteNotes.map((note) => (
              <button className="favorite-card note-favorite" key={note.id} onClick={() => openNote(note)} type="button">
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
              <TaskListRow
                activeWorkTaskId={activeWorkTaskId}
                key={item.id}
                onOpen={openFavoriteTask}
                onToggleFavorite={toggleTaskFavorite}
                projectLabel={projects[item.projectId].name}
                task={item}
              />
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
            {[...reminderPlanner.overdue, ...reminderPlanner.today, ...reminderPlanner.tomorrow].slice(0, 5).map((item) => (
              <button className={`reminder-title home-reminder ${reminderState(item, now)}`} key={item.id} onClick={() => openWorkTask(item)} type="button">
                <strong>{item.title}</strong>
                <span>{projects[item.projectId].name} - {reminderLabel(item, now)}</span>
              </button>
            ))}
            {reminderPlanner.overdue.length + reminderPlanner.today.length + reminderPlanner.tomorrow.length === 0 && <p className="empty">No reminders or due dates due soon.</p>}
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
              <button className="note-snapshot" key={note.id} onClick={() => openNote(note)} type="button">
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
              AI Engine
            </h2>
          </div>
          <p className="context-copy">Turn source notes, uploaded documents, PowerPoint text, and rough ideas into structured prompts for documents, summaries, reports, emails, and image generation.</p>
          <div className="ai-studio-actions">
            <button className="primary-button ai-action" onClick={openAiStudio} type="button">
              <WandSparkles size={16} />
              Prepare prompt
            </button>
            <button className="ghost-button" onClick={openHomeProjectDashboard} type="button">
              <BarChart3 size={16} />
              Project dashboard
            </button>
          </div>
        </section>
      </section>

      <section className="home-project-dashboard" id="home-project-dashboard">
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
                  setViewMode("work");
                  window.setTimeout(() => {
                    document.getElementById("quick-create-section")?.scrollIntoView({ behavior: "smooth", block: "start" });
                  }, 80);
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
      </section>
    </section>
  );
}
