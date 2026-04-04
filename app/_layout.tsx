import { Stack } from "expo-router";
import { StatusBar } from "expo-status-bar";
import * as SystemUI from "expo-system-ui";
import { useEffect } from "react";
import { GestureHandlerRootView } from "react-native-gesture-handler";
import { SafeAreaProvider } from "react-native-safe-area-context";

import { EditImagesProvider } from "@/src/context/edit-images-context";
import { electricCuratorTheme } from "@/src/theme/electric-curator";

export default function RootLayout() {
  useEffect(() => {
    void SystemUI.setBackgroundColorAsync(electricCuratorTheme.colors.surface);
  }, []);

  return (
    <GestureHandlerRootView style={{ flex: 1 }}>
      <SafeAreaProvider>
        <EditImagesProvider>
          <StatusBar
            style="dark"
            backgroundColor={electricCuratorTheme.colors.surface}
          />
          <Stack
            screenOptions={{
              headerShown: false,
              contentStyle: {
                backgroundColor: electricCuratorTheme.colors.surface,
              },
            }}
          />
        </EditImagesProvider>
      </SafeAreaProvider>
    </GestureHandlerRootView>
  );
}
