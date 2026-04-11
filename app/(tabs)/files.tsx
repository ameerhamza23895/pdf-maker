import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import { useFocusEffect } from "@react-navigation/native";
import * as Clipboard from "expo-clipboard";
import { useRouter } from "expo-router";
import { useCallback, useState } from "react";
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

import {
  deleteSavedFile,
  listSavedFiles,
  type SavedFileHistoryRow,
} from "@/src/db/savedFileHistory";
import { setPendingEditDocument } from "@/src/navigation/pendingEditDocument";
import { electricCuratorTheme, withAlpha } from "@/src/theme/electric-curator";
import {
  getLocationLabel,
  getSavedFileFolderUri,
  openFileWithSystemViewer,
} from "@/src/utils/openSavedFileLocation";

const { colors, spacing, radius, typography } = electricCuratorTheme;

const ui = {
  textMuted: withAlpha(colors.onSurface, 0.65),
  border: colors.outlineVariant,
};

function formatSavedAt(ts: number): string {
  try {
    return new Date(ts).toLocaleString(undefined, {
      dateStyle: "medium",
      timeStyle: "short",
    });
  } catch {
    return String(ts);
  }
}

export default function FilesPage() {
  const router = useRouter();
  const [rows, setRows] = useState<SavedFileHistoryRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);

  const load = useCallback(async () => {
    try {
      const next = await listSavedFiles();
      setRows(next);
    } catch (e) {
      console.warn("[files] listSavedFiles", e);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  }, []);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load]),
  );

  const onRefresh = useCallback(() => {
    setRefreshing(true);
    void load();
  }, [load]);

  const openInEditor = useCallback(
    (item: SavedFileHistoryRow) => {
      setPendingEditDocument({
        uri: item.uri,
        name: item.file_name,
      });
      router.push("/edit-pdf");
    },
    [router],
  );

  const tryOpenFileExternally = useCallback(async (uri: string) => {
    try {
      await openFileWithSystemViewer(uri);
      return;
    } catch {
      /* continue */
    }

    Alert.alert(
      "Open file",
      "This path cannot be opened with the system viewer from here. Use Edit in app instead.",
    );
  }, []);

  const openFolder = useCallback(
    (item: SavedFileHistoryRow) => {
      const folderUri = getSavedFileFolderUri(item.uri, item.directory_uri);

      if (!folderUri) {
        Alert.alert(
          "Open folder",
          "We could not find the containing folder for this saved file.",
        );
        return;
      }

      const title = getLocationLabel(folderUri, "Folder");
      router.push(
        `/folder-viewer?uri=${encodeURIComponent(folderUri)}&title=${encodeURIComponent(title)}`,
      );
    },
    [router],
  );

  const confirmDelete = useCallback(
    (item: SavedFileHistoryRow) => {
      Alert.alert(
        "Remove from history?",
        item.file_name,
        [
          { text: "Cancel", style: "cancel" },
          {
            text: "Remove",
            style: "destructive",
            onPress: () => {
              void (async () => {
                await deleteSavedFile(item.id);
                await load();
              })();
            },
          },
        ],
        { cancelable: true },
      );
    },
    [load],
  );

  const showRowActions = useCallback(
    (item: SavedFileHistoryRow) => {
      Alert.alert(item.file_name, undefined, [
        {
          text: "Edit in app",
          onPress: () => openInEditor(item),
        },
        {
          text: "Open folder",
          onPress: () => openFolder(item),
        },
        {
          text: "Open with system...",
          onPress: () => void tryOpenFileExternally(item.uri),
        },
        {
          text: "Copy path",
          onPress: () => {
            void Clipboard.setStringAsync(item.uri);
          },
        },
        {
          text: "Remove from history",
          style: "destructive",
          onPress: () => confirmDelete(item),
        },
        { text: "Cancel", style: "cancel" },
      ]);
    },
    [confirmDelete, openFolder, openInEditor, tryOpenFileExternally],
  );

  return (
    <View
      style={[
        styles.root,
        {
          paddingTop: 6,
        },
      ]}
    >
      <View style={styles.headerBlock}>
        <Text style={typography.labelMd}>Save history</Text>
        <Text style={typography.headlineMd}>Recent files</Text>
      </View>

      {loading && rows.length === 0 ? (
        <View style={styles.centered}>
          <ActivityIndicator size="large" color={colors.primary} />
        </View>
      ) : (
        <FlatList
          data={rows}
          keyExtractor={(item) => String(item.id)}
          refreshControl={
            <RefreshControl refreshing={refreshing} onRefresh={onRefresh} />
          }
          contentContainerStyle={styles.listContent}
          ListEmptyComponent={
            <View style={styles.empty}>
              <MaterialIcons
                name="folder-open"
                size={40}
                color={ui.textMuted}
              />
              <Text style={[typography.titleSm, { marginTop: spacing.sm }]}>
                Nothing saved yet
              </Text>
              <Text
                style={[
                  typography.bodyMd,
                  { color: ui.textMuted, textAlign: "center" },
                ]}
              >
                Export or save a document from the editor - it will show up
                here.
              </Text>
            </View>
          }
          renderItem={({ item }) => (
            <View style={styles.card}>
              <View style={styles.cardTop}>
                <Text style={styles.fileName} numberOfLines={2}>
                  {item.file_name}
                </Text>
                {/* <Pressable
                  hitSlop={12}
                  onPress={() => showRowActions(item)}
                  accessibilityLabel="More actions"
                >
                  <MaterialIcons
                    name="more-horiz"
                    size={22}
                    color={colors.onSurface}
                  />
                </Pressable> */}
              </View>
              <Text style={styles.meta}>
                {formatSavedAt(item.created_at)}
                {item.source ? ` - ${item.source}` : ""}
              </Text>
              <Text style={styles.path} numberOfLines={2} selectable>
                {item.uri}
              </Text>
              <View style={styles.actions}>
                <Pressable
                  style={({ pressed }) => [
                    styles.btnPrimary,
                    pressed && styles.pressed,
                  ]}
                  onPress={() => openInEditor(item)}
                >
                  <MaterialIcons
                    name="edit"
                    size={18}
                    color={colors.onPrimary}
                  />
                  <Text style={styles.btnPrimaryLabel}>Edit</Text>
                </Pressable>
                {/* <Pressable
                  style={({ pressed }) => [
                    styles.btnGhost,
                    pressed && styles.pressed,
                  ]}
                  onPress={() => openFolder(item)}
                >
                  <MaterialIcons
                    name="folder-open"
                    size={18}
                    color={colors.primary}
                  />
                  <Text style={styles.btnGhostLabel}>Open folder</Text>
                </Pressable> */}
              </View>
            </View>
          )}
        />
      )}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.surface,
    paddingHorizontal: spacing.md,
  },
  headerBlock: {
    marginBottom: spacing.md,
    gap: spacing.xs,
  },
  centered: {
    flex: 1,
    alignItems: "center",
    justifyContent: "center",
  },
  listContent: {
    paddingBottom: spacing.xl,
    gap: spacing.sm,
  },
  empty: {
    alignItems: "center",
    paddingVertical: spacing.xl,
    paddingHorizontal: spacing.md,
    gap: spacing.xs,
  },
  card: {
    borderRadius: radius.md,
    padding: spacing.md,
    backgroundColor: colors.surfaceContainerLowest,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(ui.border, 0.5),
    gap: spacing.xs,
  },
  cardTop: {
    flexDirection: "row",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: spacing.sm,
  },
  fileName: {
    ...typography.titleSm,
    flex: 1,
  },
  meta: {
    ...typography.bodyMd,
    fontSize: 12,
    color: ui.textMuted,
  },
  path: {
    fontSize: 11,
    lineHeight: 16,
    color: ui.textMuted,
  },
  actions: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: spacing.sm,
    marginTop: spacing.sm,
  },
  btnPrimary: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    paddingVertical: 10,
    paddingHorizontal: 16,
    borderRadius: radius.pill,
    backgroundColor: colors.primary,
  },
  btnPrimaryLabel: {
    color: colors.onPrimary,
    fontWeight: "700",
    fontSize: 14,
  },
  btnGhost: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    paddingVertical: 10,
    paddingHorizontal: 16,
    borderRadius: radius.pill,
    backgroundColor: colors.surfaceContainerLow,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(colors.primary, 0.35),
  },
  btnGhostLabel: {
    color: colors.primary,
    fontWeight: "700",
    fontSize: 14,
  },
  pressed: {
    opacity: 0.85,
  },
});
