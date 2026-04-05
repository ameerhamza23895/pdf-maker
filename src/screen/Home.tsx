import type { Href } from "expo-router";
import { ScrollView, Text, View } from "react-native";

import { HomeActionCard } from "@/src/components/home-action-card";
import type { IconSymbolName } from "@/src/components/icon-symbol";
import { electricCuratorTheme } from "@/src/theme/electric-curator";

const { colors, spacing, radius, typography } = electricCuratorTheme;

type HomeAction = {
  title: string;
  description: string;
  icon: IconSymbolName;
  href?: Href;
  tone?: "primary" | "sky" | "lilac" | "paper";
};

const actions: HomeAction[] = [
  {
    title: "Scan",
    description: "Camera Capture",
    icon: "camera.fill",
    href: "/scan",
    tone: "primary",
  },
  {
    title: "Edit PDF",
    description: "Upload PDF file",
    icon: "doc.fill",
    href: "/edit-pdf",
    tone: "sky",
  },
  {
    title: "Files",
    description: "Browse saved PDFs",
    icon: "folder.fill",
    href: "/files",
    tone: "paper",
  },
  {
    title: "PDF Converter",
    description: "Convert documents — PDF, Office, and more",
    icon: "doc.text.fill",
    href: "/convert-pdf",
    tone: "lilac",
  },
];

export default function HomeScreen() {
  return (
    <ScrollView
      contentInsetAdjustmentBehavior="automatic"
      style={{ flex: 1, backgroundColor: colors.surface }}
      contentContainerStyle={{
        paddingHorizontal: spacing.md,
        paddingTop: spacing.md,
        paddingBottom: 140,
        gap: spacing.lg,
      }}
    >
      <View style={{ gap: spacing.xs }}>
        <Text selectable style={typography.labelMd}>
          Quick Actions
        </Text>
        <Text selectable style={typography.headlineMd}>
          Choose your next PDF flow.
        </Text>
        {/* <Text selectable style={typography.bodyMd}>
          Small cards, strong contrast, and quick entry points for the tools
          you use most.
        </Text> */}
      </View>

      <View
        style={{
          padding: spacing.sm,
          borderRadius: radius.md,
          backgroundColor: colors.surfaceContainerLow,
        }}
      >
        <View
          style={{
            flexDirection: "row",
            flexWrap: "wrap",
            justifyContent: "space-between",
            rowGap: spacing.sm,
          }}
        >
          {actions.map((action) => (
            <HomeActionCard
              key={action.title}
              title={action.title}
              description={action.description}
              icon={action.icon}
              href={action.href}
              tone={action.tone}
            />
          ))}
        </View>
      </View>
    </ScrollView>
  );
}
