import type { Href } from "expo-router";
import { useEffect, useState } from "react";
import { ScrollView, Text, View } from "react-native";

import { useAdsOptional } from "@/src/ads/AdsProvider";
import { AdBanner } from "@/src/ads/components/AdBanner";
import { HomeActionCard } from "@/src/components/home-action-card";
import type { IconSymbolName } from "@/src/components/icon-symbol";
import { ScreenSafeArea } from "@/src/components/ScreenSafeArea";
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
    description: "Convert documents - PDF, Office, and more",
    icon: "doc.text.fill",
    href: "/convert-pdf",
    tone: "lilac",
  },
];

export default function HomeScreen() {
  const ads = useAdsOptional();
  const bannerEnabled = ads?.preferences.bannerEnabled ?? false;
  const [showBanner, setShowBanner] = useState(false);

  useEffect(() => {
    if (!bannerEnabled) {
      setShowBanner(false);
    }
  }, [bannerEnabled]);

  return (
    <ScreenSafeArea edges={["left", "right"]} style={{ flex: 1 }}>
      {bannerEnabled ? (
        <View
          style={{
            gap: spacing.xs,
            height: showBanner ? undefined : 0,
            overflow: "hidden",
          }}
        >
          <AdBanner
            visible={bannerEnabled}
            onVisibilityChange={setShowBanner}
          />
        </View>
      ) : null}
      <ScrollView
        contentInsetAdjustmentBehavior="automatic"
        style={{ flex: 1, backgroundColor: colors.surface }}
        contentContainerStyle={{
          paddingHorizontal: spacing.md,
          paddingTop: spacing.sm,
          paddingBottom: 4,
          gap: spacing.lg,
          // borderWidth: 2,
        }}
      >
        <View style={{ gap: spacing.xs }}>
          <Text selectable style={typography.labelMd}>
            Quick Actions
          </Text>
          <Text selectable style={typography.headlineMd}>
            Choose your next PDF flow
          </Text>
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
    </ScreenSafeArea>
  );
}
