import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import { ScrollView, StyleSheet, Text, View } from "react-native";
import { useSafeAreaInsets } from "react-native-safe-area-context";

import { staticUserProfile } from "@/src/data/staticUserProfile";
import { electricCuratorTheme, withAlpha } from "@/src/theme/electric-curator";

const { colors, spacing, radius, typography } = electricCuratorTheme;

const muted = withAlpha(colors.onSurface, 0.68);

export default function SettingsPage() {
  const insets = useSafeAreaInsets();

  return (
    <ScrollView
      style={[styles.root, { paddingTop: insets.top + spacing.sm }]}
      contentContainerStyle={{
        paddingBottom: insets.bottom + spacing.xl,
        paddingHorizontal: spacing.md,
        gap: spacing.lg,
      }}
      contentInsetAdjustmentBehavior="automatic"
    >
      <View style={styles.intro}>
        <Text style={typography.labelMd}>Settings</Text>
        <Text style={typography.headlineMd}>Profile & account</Text>
        <Text style={[typography.bodyMd, { color: muted }]}>
          Account details below are static placeholders. Wire them to Firebase Auth / Firestore
          when you are ready.
        </Text>
      </View>

      <View style={styles.profileCard}>
        <View style={styles.avatar}>
          <MaterialIcons name="person" size={36} color={colors.primary} />
        </View>
        <View style={styles.profileText}>
          <Text style={typography.titleSm}>{staticUserProfile.displayName}</Text>
          <Text style={[typography.bodyMd, { color: muted }]}>{staticUserProfile.email}</Text>
        </View>
      </View>

      <View style={styles.section}>
        <Text style={styles.sectionLabel}>Account</Text>
        <View style={styles.row}>
          <MaterialIcons name="login" size={22} color={colors.primary} />
          <View style={styles.rowBody}>
            <Text style={typography.titleSm}>Sign in</Text>
            <Text style={[typography.bodyMd, { color: muted, fontSize: 13 }]}>
              {staticUserProfile.accountHint}
            </Text>
          </View>
          <MaterialIcons name="chevron-right" size={22} color={muted} />
        </View>
      </View>

      <View style={styles.section}>
        <Text style={styles.sectionLabel}>App</Text>
        <View style={styles.row}>
          <MaterialIcons name="history" size={22} color={colors.primary} />
          <View style={styles.rowBody}>
            <Text style={typography.titleSm}>Save history</Text>
            <Text style={[typography.bodyMd, { color: muted, fontSize: 13 }]}>
              Stored on-device in SQLite (Files tab).
            </Text>
          </View>
        </View>
      </View>
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.surface,
  },
  intro: {
    gap: spacing.xs,
  },
  profileCard: {
    flexDirection: "row",
    alignItems: "center",
    gap: spacing.md,
    padding: spacing.md,
    borderRadius: radius.md,
    backgroundColor: colors.surfaceContainerLowest,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(colors.outlineVariant, 0.45),
  },
  avatar: {
    width: 64,
    height: 64,
    borderRadius: 32,
    backgroundColor: colors.secondaryContainer,
    alignItems: "center",
    justifyContent: "center",
  },
  profileText: {
    flex: 1,
    gap: 4,
  },
  section: {
    gap: spacing.sm,
  },
  sectionLabel: {
    ...typography.labelMd,
    fontSize: 11,
    opacity: 0.85,
  },
  row: {
    flexDirection: "row",
    alignItems: "center",
    gap: spacing.md,
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.md,
    borderRadius: radius.md,
    backgroundColor: colors.surfaceContainerLow,
  },
  rowBody: {
    flex: 1,
    gap: 4,
  },
});
