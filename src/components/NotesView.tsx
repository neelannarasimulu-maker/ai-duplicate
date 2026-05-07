import { Pin, Plus, Save, Star, StickyNote, Trash2 } from "lucide-react";
import type { Dispatch, SetStateAction } from "react";
import type { AppNote, ProjectId } from "../types";

interface NotesViewProps {
  projectId: ProjectId;
  projects: Record<ProjectId, { name: string }>;
  notes: AppNote[];
  noteDraft: { title: string; content: string };
  setNoteDraft: Dispatch<SetStateAction<{ title: string; content: string }>>;
  updateProject: (id: ProjectId) => void;
  createNote: () => void;
  updateNote: (id: string, patch: Partial<AppNote>) => void;
  updateNoteEntry: (noteId: string, entryId: string, content: string) => void;
  addNoteEntry: (id: string) => void;
  removeNoteEntry: (noteId: string, entryId: string) => void;
  removeNote: (id: string) => void;
}

export function NotesView({
  projectId,
  projects,
  notes,
  noteDraft,
  setNoteDraft,
  updateProject,
  createNote,
  updateNote,
  updateNoteEntry,
  addNoteEntry,
  removeNoteEntry,
  removeNote,
}: NotesViewProps) {
  return (
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
          .sort((a, b) => Number(b.pinned) - Number(a.pinned) || new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
          .map((note) => (
            <article className={`${note.pinned ? "note-card pinned" : "note-card"} ${note.favorite ? "favorite" : ""}`} id={`note-${note.id}`} key={note.id}>
              <div className="note-card-header">
                <span className={`project-chip project-${note.projectId}`}>{projects[note.projectId].name}</span>
                <div className="note-card-actions">
                  <button
                    className={note.favorite ? "ghost-button icon-button favorite-toggle active" : "ghost-button icon-button favorite-toggle"}
                    onClick={() => updateNote(note.id, { favorite: !note.favorite })}
                    type="button"
                    title={note.favorite ? "Remove favorite" : "Add favorite"}
                  >
                    <Star size={15} />
                  </button>
                  <button
                    className="ghost-button icon-button"
                    onClick={() => updateNote(note.id, { pinned: !note.pinned })}
                    type="button"
                    title={note.pinned ? "Unpin note" : "Pin note"}
                  >
                    <Pin size={15} />
                  </button>
                </div>
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
                      <button
                        className="ghost-button icon-button"
                        onClick={() => removeNoteEntry(note.id, entry.id)}
                        type="button"
                        title="Delete this note entry"
                      >
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
  );
}
