import type { AppNote, SavedOutput, WorkTask } from "./types";

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
  return remoteGetPayloads<WorkTask>("work_tasks", taskStore, "updated_at.desc", taskTimestamp, () => getAll<WorkTask>(taskStore));
}

export function getCachedTasks() {
  return getAll<WorkTask>(taskStore);
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
  return remoteGetPayloads<SavedOutput>("saved_outputs", outputStore, "created_at.desc", outputTimestamp, () => getAll<SavedOutput>(outputStore));
}

export function getCachedOutputs() {
  return getAll<SavedOutput>(outputStore);
}

export async function saveOutput(output: SavedOutput) {
  await put(outputStore, output);
  await remoteUpsertPayload("saved_outputs", output.id, output, output.createdAt);
}

export function getNotes() {
  return getMeta<AppNote[]>("appNotes", []);
}

export function getCachedNotes() {
  return getCachedMeta<AppNote[]>("appNotes", []);
}

export function saveNotes(notes: AppNote[]) {
  return setMeta("appNotes", notes);
}

export async function getMeta<T>(key: string, fallback: T) {
  const remoteValue = await remoteGetMeta<T>(key);
  if (remoteValue !== undefined) {
    await put(metaStore, { key, value: remoteValue });
    return remoteValue;
  }

  return getCachedMeta(key, fallback);
}

export async function getCachedMeta<T>(key: string, fallback: T) {
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

async function remoteGetPayloads<T extends WorkTask | SavedOutput>(
  table: "work_tasks" | "saved_outputs",
  storeName: string,
  order: string,
  getTimestamp: (payload: T) => string,
  fallback: () => Promise<T[]>,
) {
  let localPayloads: T[] = [];
  try {
    localPayloads = await fallback();
  } catch (error) {
    console.warn(`Local cache read failed for ${storeName}; pulling from Supabase instead.`, error);
  }

  if (!supabaseConfigured) return localPayloads;

  try {
    const response = await supabaseFetch(`/rest/v1/${table}?select=id,payload,updated_at&order=${order}`, {
      cache: "no-store",
    });
    const rows = (await response.json()) as Array<{ id: string; payload: T; updated_at: string }>;
    const remotePayloads = rows.map((row) => row.payload).filter((payload): payload is T => Boolean(payload?.id));
    const remoteById = new Map(remotePayloads.map((payload) => [payload.id, payload]));
    const mergedById = new Map(remoteById);
    const localPayloadsToPush: T[] = [];
    const localPayloadsToRemove: T[] = [];

    await Promise.all(
      localPayloads.map(async (localPayload) => {
        const remotePayload = remoteById.get(localPayload.id);
        if (!remotePayload) {
          const deletedAt = await remoteGetMeta<string>(deleteMetaKey(table, localPayload.id));
          if (deletedAt) {
            localPayloadsToRemove.push(localPayload);
            return;
          }
          localPayloadsToPush.push(localPayload);
          mergedById.set(localPayload.id, localPayload);
          return;
        }

        if (timestampValue(getTimestamp(localPayload)) > timestampValue(getTimestamp(remotePayload))) {
          localPayloadsToPush.push(localPayload);
          mergedById.set(localPayload.id, localPayload);
        }
      }),
    );

    const mergedPayloads = Array.from(mergedById.values());

    void Promise.all([
      ...mergedPayloads.map((payload) => put(storeName, payload)),
      ...localPayloadsToPush.map((payload) => remoteUpsertPayload(table, payload.id, payload, getTimestamp(payload))),
      ...localPayloadsToRemove.map((payload) => remove(storeName, payload.id)),
    ]).catch(() => {
      // Merged data is still returned to the app even if cache refresh or retry upload fails.
    });

    return mergedPayloads.sort((a, b) => timestampValue(getTimestamp(b)) - timestampValue(getTimestamp(a)));
  } catch (error) {
    console.error(`Supabase sync failed for ${table}`, error);
    if (localPayloads.length === 0) {
      throw error instanceof Error ? error : new Error(`Supabase sync failed for ${table}`);
    }
    return localPayloads;
  }
}

async function remoteUpsertPayload(table: "work_tasks" | "saved_outputs", id: string, payload: WorkTask | SavedOutput, timestamp: string) {
  if (!supabaseConfigured) return;

  try {
    const deletedAt = await remoteGetMeta<string>(deleteMetaKey(table, id));
    if (deletedAt) {
      await remove(table === "work_tasks" ? taskStore : outputStore, id);
      return;
    }

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

function taskTimestamp(task: WorkTask) {
  return task.updatedAt || task.createdAt || "";
}

function outputTimestamp(output: SavedOutput) {
  return output.createdAt || "";
}

function timestampValue(timestamp: string) {
  const value = new Date(timestamp).getTime();
  return Number.isFinite(value) ? value : 0;
}

async function remoteDelete(table: "work_tasks" | "saved_outputs", id: string) {
  if (!supabaseConfigured) return;

  try {
    await remoteSetMeta(deleteMetaKey(table, id), new Date().toISOString());
    await supabaseFetch(`/rest/v1/${table}?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
  } catch {
    // Local delete already happened; Supabase sync will resume when available.
  }
}

function deleteMetaKey(table: "work_tasks" | "saved_outputs", id: string) {
  return `deleted:${table}:${id}`;
}

async function remoteGetMeta<T>(key: string) {
  if (!supabaseConfigured) return undefined;

  try {
    const response = await supabaseFetch(`/rest/v1/app_meta?key=eq.${encodeURIComponent(key)}&select=value&limit=1`, {
      cache: "no-store",
    });
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
    cache: init.cache ?? "no-store",
    headers: {
      apikey: supabaseAnonKey,
      Authorization: `Bearer ${supabaseAnonKey}`,
      "Content-Type": "application/json",
      ...(init.headers ?? {}),
    },
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`Supabase HTTP ${response.status}${body ? `: ${body}` : ""}`);
  }
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
