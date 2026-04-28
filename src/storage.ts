import type { SavedOutput, WorkTask } from "./types";

const dbName = "ai-workbench-db";
const dbVersion = 1;
const taskStore = "tasks";
const outputStore = "outputs";
const metaStore = "meta";
const supabaseUrl = (import.meta.env.VITE_SUPABASE_URL ?? import.meta.env.NEXT_PUBLIC_SUPABASE_URL ?? "").replace(/\/$/, "");
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY ?? import.meta.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ?? "";
const supabaseConfigured = Boolean(supabaseUrl && supabaseAnonKey);

let dbPromise: Promise<IDBDatabase> | undefined;

function openDatabase() {
  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(dbName, dbVersion);

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(taskStore)) db.createObjectStore(taskStore, { keyPath: "id" });
        if (!db.objectStoreNames.contains(outputStore)) db.createObjectStore(outputStore, { keyPath: "id" });
        if (!db.objectStoreNames.contains(metaStore)) db.createObjectStore(metaStore, { keyPath: "key" });
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  return dbPromise;
}

async function getAll<T>(storeName: string) {
  const db = await openDatabase();
  return new Promise<T[]>((resolve, reject) => {
    const transaction = db.transaction(storeName, "readonly");
    const request = transaction.objectStore(storeName).getAll();
    request.onsuccess = () => resolve(request.result as T[]);
    request.onerror = () => reject(request.error);
  });
}

async function put<T>(storeName: string, value: T) {
  const db = await openDatabase();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    transaction.objectStore(storeName).put(value);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

async function remove(storeName: string, id: string) {
  const db = await openDatabase();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(storeName, "readwrite");
    transaction.objectStore(storeName).delete(id);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

export function getTasks() {
  return remoteGetPayloads<WorkTask>("work_tasks", "updated_at.desc", () => getAll<WorkTask>(taskStore));
}

export async function saveTask(task: WorkTask) {
  await put(taskStore, task);
  await remoteUpsertPayload("work_tasks", task.id, task, task.updatedAt);
}

export async function deleteTask(id: string) {
  await remove(taskStore, id);
  await remoteDelete("work_tasks", id);
}

export function getOutputs() {
  return remoteGetPayloads<SavedOutput>("saved_outputs", "created_at.desc", () => getAll<SavedOutput>(outputStore));
}

export async function saveOutput(output: SavedOutput) {
  await put(outputStore, output);
  await remoteUpsertPayload("saved_outputs", output.id, output, output.createdAt);
}

export async function getMeta<T>(key: string, fallback: T) {
  const remoteValue = await remoteGetMeta<T>(key);
  if (remoteValue !== undefined) return remoteValue;

  const db = await openDatabase();
  return new Promise<T>((resolve, reject) => {
    const transaction = db.transaction(metaStore, "readonly");
    const request = transaction.objectStore(metaStore).get(key);
    request.onsuccess = () => resolve(request.result ? (request.result.value as T) : fallback);
    request.onerror = () => reject(request.error);
  });
}

export async function setMeta<T>(key: string, value: T) {
  await put(metaStore, { key, value });
  await remoteSetMeta(key, value);
}

export function getStorageBackendLabel() {
  return supabaseConfigured ? "Supabase" : "IndexedDB";
}

export async function migrateLegacyStorage(taskKey: string, outputKey: string, reminderKey: string) {
  const migrationDone = await getMeta("legacyMigrationDone", false);
  if (migrationDone) return;

  const legacyTasks = readLegacyArray<WorkTask>(taskKey);
  const legacyOutputs = readLegacyArray<SavedOutput>(outputKey);
  const legacyReminderIds = readLegacyArray<string>(reminderKey);

  await Promise.all([
    ...legacyTasks.map((task) => saveTask(task)),
    ...legacyOutputs.map((output) => saveOutput(output)),
    setMeta("triggeredReminderIds", legacyReminderIds),
    setMeta("legacyMigrationDone", true),
  ]);
}

async function remoteGetPayloads<T>(table: "work_tasks" | "saved_outputs", order: string, fallback: () => Promise<T[]>) {
  if (!supabaseConfigured) return fallback();

  try {
    const response = await supabaseFetch(`/rest/v1/${table}?select=payload&order=${order}`);
    const rows = (await response.json()) as Array<{ payload: T }>;
    return rows.map((row) => row.payload);
  } catch {
    return fallback();
  }
}

async function remoteUpsertPayload(table: "work_tasks" | "saved_outputs", id: string, payload: WorkTask | SavedOutput, timestamp: string) {
  if (!supabaseConfigured) return;

  try {
    await supabaseFetch(`/rest/v1/${table}?on_conflict=id`, {
      method: "POST",
      headers: { Prefer: "resolution=merge-duplicates" },
      body: JSON.stringify({
        id,
        payload,
        created_at: "createdAt" in payload ? payload.createdAt : timestamp,
        updated_at: timestamp,
      }),
    });
  } catch {
    // IndexedDB already has a local copy; Supabase sync will resume when available.
  }
}

async function remoteDelete(table: "work_tasks" | "saved_outputs", id: string) {
  if (!supabaseConfigured) return;

  try {
    await supabaseFetch(`/rest/v1/${table}?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
  } catch {
    // Local delete already happened; Supabase sync will resume when available.
  }
}

async function remoteGetMeta<T>(key: string) {
  if (!supabaseConfigured) return undefined;

  try {
    const response = await supabaseFetch(`/rest/v1/app_meta?key=eq.${encodeURIComponent(key)}&select=value&limit=1`);
    const rows = (await response.json()) as Array<{ value: T }>;
    return rows[0]?.value;
  } catch {
    return undefined;
  }
}

async function remoteSetMeta<T>(key: string, value: T) {
  if (!supabaseConfigured) return;

  try {
    await supabaseFetch("/rest/v1/app_meta?on_conflict=key", {
      method: "POST",
      headers: { Prefer: "resolution=merge-duplicates" },
      body: JSON.stringify({ key, value }),
    });
  } catch {
    // IndexedDB already has a local copy; Supabase sync will resume when available.
  }
}

async function supabaseFetch(path: string, init: RequestInit = {}) {
  const response = await fetch(`${supabaseUrl}${path}`, {
    ...init,
    headers: {
      apikey: supabaseAnonKey,
      Authorization: `Bearer ${supabaseAnonKey}`,
      "Content-Type": "application/json",
      ...(init.headers ?? {}),
    },
  });

  if (!response.ok) throw new Error(`Supabase HTTP ${response.status}`);
  return response;
}

function readLegacyArray<T>(key: string) {
  try {
    const value = localStorage.getItem(key);
    return value ? (JSON.parse(value) as T[]) : [];
  } catch {
    return [];
  }
}
