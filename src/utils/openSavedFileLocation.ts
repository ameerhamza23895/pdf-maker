import * as Linking from "expo-linking";
import { Alert, Platform } from "react-native";

/**
 * Best-effort: open a SAF tree URI (Android) or show guidance when unavailable.
 */
export async function openSavedFileFolder(
  directoryUri: string | null | undefined,
): Promise<void> {
  if (directoryUri) {
    try {
      const can = await Linking.canOpenURL(directoryUri);
      if (can) {
        await Linking.openURL(directoryUri);
        return;
      }
    } catch {
      /* fall through */
    }
  }

  Alert.alert(
    "Folder",
    Platform.OS === "android"
      ? "No folder link is stored for this save. When you save to device storage, allow the folder picker so we can open that location later."
      : "Opening the containing folder is only available for saves that include a folder link (Android folder picker).",
  );
}
