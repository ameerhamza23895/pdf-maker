import * as LegacyFileSystem from "expo-file-system/legacy";
import * as Linking from "expo-linking";
import { Platform } from "react-native";

const FILE_SCHEME = "file://";
const CONTENT_SCHEME = "content://";

type SafUriParts = {
  authority: string;
  treeIdEncoded: string;
  documentIdEncoded: string | null;
};

function safeDecode(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

export function isFileUri(uri: string | null | undefined): uri is string {
  return Boolean(uri?.startsWith(FILE_SCHEME));
}

export function isSafUri(uri: string | null | undefined): uri is string {
  return Boolean(uri?.startsWith(CONTENT_SCHEME));
}

function parseSafUri(uri: string): SafUriParts | null {
  const match = uri.match(
    /^content:\/\/([^/]+)\/tree\/([^/]+)(?:\/document\/([^/]+))?$/,
  );

  if (!match) {
    return null;
  }

  return {
    authority: match[1],
    treeIdEncoded: match[2],
    documentIdEncoded: match[3] ?? null,
  };
}

export function getSafTreeUri(
  uri: string | null | undefined,
): string | null {
  const value = uri?.trim();
  if (!value || !isSafUri(value)) {
    return null;
  }

  const parsed = parseSafUri(value);
  if (!parsed) {
    return null;
  }

  const targetIdEncoded = parsed.documentIdEncoded ?? parsed.treeIdEncoded;
  return `content://${parsed.authority}/tree/${targetIdEncoded}`;
}

export function getSafParentTreeUri(
  uri: string | null | undefined,
): string | null {
  const value = uri?.trim();
  if (!value || !isSafUri(value)) {
    return null;
  }

  const parsed = parseSafUri(value);
  if (!parsed) {
    return null;
  }

  const documentIdEncoded = parsed.documentIdEncoded ?? parsed.treeIdEncoded;
  const documentId = safeDecode(documentIdEncoded);
  const parentBoundary = documentId.lastIndexOf("/");
  const parentIdEncoded =
    parentBoundary === -1
      ? parsed.treeIdEncoded
      : encodeURIComponent(documentId.slice(0, parentBoundary));

  return `content://${parsed.authority}/tree/${parentIdEncoded}`;
}

export function getSavedFileFolderUri(
  fileUri: string | null | undefined,
  directoryUri: string | null | undefined,
): string | null {
  const storedDirectoryUri = directoryUri?.trim();
  if (storedDirectoryUri) {
    return getSafTreeUri(storedDirectoryUri) ?? storedDirectoryUri;
  }

  const savedFileUri = fileUri?.trim();
  if (!savedFileUri) {
    return null;
  }

  if (isFileUri(savedFileUri)) {
    const normalized = savedFileUri.replace(/\/+$/, "");
    const lastSlashIndex = normalized.lastIndexOf("/");

    return lastSlashIndex >= FILE_SCHEME.length
      ? normalized.slice(0, lastSlashIndex + 1)
      : null;
  }

  return getSafParentTreeUri(savedFileUri);
}

export function getLocationLabel(
  uri: string | null | undefined,
  fallback = "Folder",
): string {
  const value = uri?.trim();
  if (!value) {
    return fallback;
  }

  if (isFileUri(value)) {
    const normalized = value.replace(/\/+$/, "");
    const label = normalized.slice(normalized.lastIndexOf("/") + 1);
    return label || fallback;
  }

  const decoded = safeDecode(value.replace(/\/+$/, ""));
  const lastSegment = decoded.slice(decoded.lastIndexOf("/") + 1);
  const label = lastSegment.includes(":")
    ? lastSegment.slice(lastSegment.lastIndexOf(":") + 1)
    : lastSegment;

  return label || fallback;
}

export function getLocationSubtitle(uri: string | null | undefined): string {
  const value = uri?.trim();
  if (!value) {
    return "";
  }

  return safeDecode(value)
    .replace(FILE_SCHEME, "")
    .replace(CONTENT_SCHEME, "");
}

export function joinFileUri(parentUri: string, name: string): string {
  return `${parentUri.replace(/\/?$/, "/")}${name}`;
}

export async function openFileWithSystemViewer(uri: string): Promise<void> {
  const targetUri =
    Platform.OS === "android" && isFileUri(uri)
      ? await LegacyFileSystem.getContentUriAsync(uri)
      : uri;

  await Linking.openURL(targetUri);
}
