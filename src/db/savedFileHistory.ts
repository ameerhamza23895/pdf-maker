import * as SQLite from "expo-sqlite";

export type SavedFileHistoryRow = {
  id: number;
  uri: string;
  file_name: string;
  mime_type: string | null;
  source: string | null;
  directory_uri: string | null;
  created_at: number;
};

export type RecordSavedFileInput = {
  uri: string;
  fileName: string;
  mimeType?: string | null;
  source?: string;
  directoryUri?: string | null;
};

let dbPromise: Promise<SQLite.SQLiteDatabase> | null = null;

async function getDb(): Promise<SQLite.SQLiteDatabase> {
  if (!dbPromise) {
    dbPromise = (async () => {
      const db = await SQLite.openDatabaseAsync("saved_file_history.db");
      await db.execAsync(`
        PRAGMA journal_mode = WAL;
        CREATE TABLE IF NOT EXISTS saved_file_history (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          uri TEXT NOT NULL,
          file_name TEXT NOT NULL,
          mime_type TEXT,
          source TEXT,
          directory_uri TEXT,
          created_at INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_saved_file_created ON saved_file_history(created_at DESC);
      `);
      return db;
    })();
  }
  return dbPromise;
}

export async function recordSavedFile(input: RecordSavedFileInput): Promise<void> {
  try {
    const db = await getDb();
    await db.runAsync(
      `INSERT INTO saved_file_history (uri, file_name, mime_type, source, directory_uri, created_at)
       VALUES (?, ?, ?, ?, ?, ?)`,
      [
        input.uri,
        input.fileName,
        input.mimeType ?? null,
        input.source ?? "unknown",
        input.directoryUri ?? null,
        Date.now(),
      ],
    );
  } catch (e) {
    console.warn("[savedFileHistory] recordSavedFile failed", e);
  }
}

export async function listSavedFiles(
  limit = 200,
): Promise<SavedFileHistoryRow[]> {
  const db = await getDb();
  const rows = await db.getAllAsync<SavedFileHistoryRow>(
    `SELECT id, uri, file_name, mime_type, source, directory_uri, created_at
     FROM saved_file_history
     ORDER BY created_at DESC
     LIMIT ?`,
    [limit],
  );
  return rows ?? [];
}

export async function deleteSavedFile(id: number): Promise<void> {
  const db = await getDb();
  await db.runAsync(`DELETE FROM saved_file_history WHERE id = ?`, [id]);
}
