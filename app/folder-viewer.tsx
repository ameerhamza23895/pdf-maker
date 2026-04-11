import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import * as LegacyFileSystem from "expo-file-system/legacy";
import { Stack, useLocalSearchParams, useRouter } from "expo-router";
import { useCallback, useEffect, useMemo, useState } from "react";
import {
  ActivityIndicator,
  Alert,
  FlatList,
  Pressable,
  RefreshControl,
  StyleSheet,
  Text,
  View,
} from "react-native";

import { ScreenSafeArea } from "@/src/components/ScreenSafeArea";
import { setPendingEditDocument } from "@/src/navigation/pendingEditDocument";
import { electricCuratorTheme, withAlpha } from "@/src/theme/electric-curator";
import {
  getLocationLabel,
  getLocationSubtitle,
  isFileUri,
  isSafUri,
  joinFileUri,
  getSafTreeUri,
  openFileWithSystemViewer,
} from "@/src/utils/openSavedFileLocation";

const { colors, spacing, radius, typography } = electricCuratorTheme;

type FolderEntry = {
  uri: string;
  name: string;
  isDirectory: boolean;
  size: number | null;
};

function firstParam(
  value: string | string[] | undefined,
): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

function formatFileSize(size: number | null): string {
  if (size == null || !Number.isFinite(size)) {
    return "File";
  }
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

async function listFolderEntries(directoryUri: string): Promise<FolderEntry[]> {
  let childUris: string[];

  if (isSafUri(directoryUri)) {
    if (!LegacyFileSystem.StorageAccessFramework) {
      throw new Error("Storage Access Framework is not available on this device.");
    }

    const safeDirectoryUri = getSafTreeUri(directoryUri) ?? directoryUri;
    childUris =
      await LegacyFileSystem.StorageAccessFramework.readDirectoryAsync(
        safeDirectoryUri,
      );
  } else if (isFileUri(directoryUri)) {
    const names = await LegacyFileSystem.readDirectoryAsync(directoryUri);
    childUris = names.map((name) => joinFileUri(directoryUri, name));
  } else {
    throw new Error("This folder path is not supported.");
  }

  const items = await Promise.all(
    childUris.map(async (childUri) => {
      if (isSafUri(childUri) && LegacyFileSystem.StorageAccessFramework) {
        const childTreeUri = getSafTreeUri(childUri);

        if (childTreeUri) {
          try {
            await LegacyFileSystem.StorageAccessFramework.readDirectoryAsync(
              childTreeUri,
            );

            return {
              uri: childTreeUri,
              name: getLocationLabel(childTreeUri, "Untitled"),
              isDirectory: true,
              size: null,
            };
          } catch {
            /* continue as file */
          }
        }
      }

      try {
        const info = await LegacyFileSystem.getInfoAsync(childUri);

        return {
          uri:
            info.exists && info.isDirectory && isFileUri(childUri)
              ? childUri.replace(/\/?$/, "/")
              : childUri,
          name: getLocationLabel(childUri, "Untitled"),
          isDirectory: info.exists ? info.isDirectory : false,
          size: info.exists && !info.isDirectory ? info.size : null,
        };
      } catch {
        return {
          uri: childUri,
          name: getLocationLabel(childUri, "Untitled"),
          isDirectory: false,
          size: null,
        };
      }
    }),
  );

  return items.sort((left, right) => {
    if (left.isDirectory !== right.isDirectory) {
      return left.isDirectory ? -1 : 1;
    }
    return left.name.localeCompare(right.name, undefined, {
      sensitivity: "base",
    });
  });
}

export default function FolderViewerPage() {
  const router = useRouter();
  const params = useLocalSearchParams<{ uri?: string; title?: string }>();
  const directoryUri = firstParam(params.uri) ?? "";
  const titleParam = firstParam(params.title);

  const screenTitle = useMemo(
    () => titleParam || getLocationLabel(directoryUri, "Folder"),
    [directoryUri, titleParam],
  );
  const folderPath = useMemo(
    () => getLocationSubtitle(directoryUri),
    [directoryUri],
  );

  const [entries, setEntries] = useState<FolderEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const loadEntries = useCallback(async () => {
    if (!directoryUri) {
      setEntries([]);
      setErrorMessage("No folder location was provided.");
      setLoading(false);
      setRefreshing(false);
      return;
    }

    try {
      const nextEntries = await listFolderEntries(directoryUri);
      setEntries(nextEntries);
      setErrorMessage(null);
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "This folder could not be opened.";
      setEntries([]);
      setErrorMessage(message);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  }, [directoryUri]);

  useEffect(() => {
    setLoading(true);
    void loadEntries();
  }, [loadEntries]);

  const onRefresh = useCallback(() => {
    setRefreshing(true);
    void loadEntries();
  }, [loadEntries]);

  const openEntry = useCallback(
    async (entry: FolderEntry) => {
      if (entry.isDirectory) {
        const nextDirectoryUri = isSafUri(entry.uri)
          ? getSafTreeUri(entry.uri) ?? entry.uri
          : entry.uri;

        router.push(
          `/folder-viewer?uri=${encodeURIComponent(nextDirectoryUri)}&title=${encodeURIComponent(entry.name)}`,
        );
        return;
      }

      if (entry.name.toLowerCase().endsWith(".pdf")) {
        setPendingEditDocument({
          uri: entry.uri,
          name: entry.name,
        });
        router.push("/edit-pdf");
        return;
      }

      try {
        await openFileWithSystemViewer(entry.uri);
      } catch {
        Alert.alert(
          "Open file",
          "This file cannot be opened from here. PDF files can be edited in the app, and other files need a compatible system viewer.",
        );
      }
    },
    [router],
  );

  const emptyTitle = errorMessage
    ? "Folder could not be opened"
    : "This folder is empty";
  const emptyBody = errorMessage
    ? "Check folder access and try saving to that location again."
    : "Pull to refresh if you just saved a file here.";

  return (
    <>
      <Stack.Screen
        options={{
          title: screenTitle,
          headerShown: true,
          headerStyle: { backgroundColor: colors.surfaceContainerLow },
          headerTintColor: colors.onSurface,
          headerTitleStyle: { fontWeight: "700" },
          headerBackTitle: "Back",
        }}
      />
      <ScreenSafeArea edges={["left", "right", "bottom"]} style={styles.root}>
        <FlatList
          data={entries}
          keyExtractor={(item) => item.uri}
          refreshControl={
            <RefreshControl refreshing={refreshing} onRefresh={onRefresh} />
          }
          contentContainerStyle={styles.listContent}
          ListHeaderComponent={
            <View style={styles.headerCard}>
              <Text style={typography.labelMd}>Folder path</Text>
              <Text style={styles.pathText} selectable>
                {folderPath || directoryUri}
              </Text>
              {errorMessage ? (
                <Text style={styles.errorText}>{errorMessage}</Text>
              ) : null}
            </View>
          }
          ListEmptyComponent={
            loading ? (
              <View style={styles.centered}>
                <ActivityIndicator size="large" color={colors.primary} />
              </View>
            ) : (
              <View style={styles.empty}>
                <MaterialIcons
                  name="folder-open"
                  size={40}
                  color={withAlpha(colors.onSurface, 0.55)}
                />
                <Text style={styles.emptyTitle}>{emptyTitle}</Text>
                <Text style={styles.emptyBody}>{emptyBody}</Text>
              </View>
            )
          }
          renderItem={({ item }) => (
            <Pressable
              style={({ pressed }) => [
                styles.entryCard,
                pressed && styles.pressed,
              ]}
              onPress={() => void openEntry(item)}
            >
              <View style={styles.entryIcon}>
                <MaterialIcons
                  name={item.isDirectory ? "folder-open" : "insert-drive-file"}
                  size={22}
                  color={item.isDirectory ? colors.primary : colors.onSurface}
                />
              </View>
              <View style={styles.entryBody}>
                <Text style={styles.entryTitle} numberOfLines={2}>
                  {item.name}
                </Text>
                <Text style={styles.entryMeta}>
                  {item.isDirectory ? "Folder" : formatFileSize(item.size)}
                </Text>
              </View>
              <MaterialIcons
                name="chevron-right"
                size={22}
                color={withAlpha(colors.onSurface, 0.45)}
              />
            </Pressable>
          )}
        />
      </ScreenSafeArea>
    </>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.surface,
  },
  listContent: {
    padding: spacing.md,
    gap: spacing.sm,
    flexGrow: 1,
  },
  headerCard: {
    borderRadius: radius.md,
    padding: spacing.md,
    backgroundColor: colors.surfaceContainerLow,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(colors.outlineVariant, 0.55),
    gap: spacing.xs,
  },
  pathText: {
    ...typography.bodyMd,
    color: withAlpha(colors.onSurface, 0.72),
  },
  errorText: {
    ...typography.bodyMd,
    color: colors.error,
  },
  centered: {
    flex: 1,
    alignItems: "center",
    justifyContent: "center",
    paddingVertical: spacing.xl,
  },
  empty: {
    alignItems: "center",
    justifyContent: "center",
    paddingVertical: spacing.xl,
    paddingHorizontal: spacing.md,
    gap: spacing.xs,
  },
  emptyTitle: {
    ...typography.titleSm,
    color: colors.onSurface,
  },
  emptyBody: {
    ...typography.bodyMd,
    color: withAlpha(colors.onSurface, 0.65),
    textAlign: "center",
  },
  entryCard: {
    flexDirection: "row",
    alignItems: "center",
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: radius.md,
    backgroundColor: colors.surfaceContainerLowest,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(colors.outlineVariant, 0.5),
  },
  entryIcon: {
    width: 40,
    height: 40,
    borderRadius: radius.sm,
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: colors.surfaceContainerLow,
  },
  entryBody: {
    flex: 1,
    gap: 2,
  },
  entryTitle: {
    ...typography.titleSm,
    color: colors.onSurface,
  },
  entryMeta: {
    ...typography.bodyMd,
    color: withAlpha(colors.onSurface, 0.65),
  },
  pressed: {
    opacity: 0.86,
  },
});
