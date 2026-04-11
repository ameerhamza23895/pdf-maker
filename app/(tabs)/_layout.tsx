import { Tabs } from "expo-router";
import { useEffect, useRef } from "react";
import { Animated, Pressable, View } from "react-native";

import { IconSymbol, type IconSymbolName } from "@/src/components/icon-symbol";
import { electricCuratorTheme } from "@/src/theme/electric-curator";

const { colors, radius, spacing } = electricCuratorTheme;

function TabIconWithLabel(
  focused: boolean,
  label: string,
  activeIcon: IconSymbolName,
  inactiveIcon: IconSymbolName,
) {
  const scale = useRef(new Animated.Value(1)).current;

  useEffect(() => {
    if (focused) {
      Animated.sequence([
        Animated.spring(scale, {
          toValue: 1.25,
          friction: 4,
          tension: 180,
          useNativeDriver: true,
        }),
        Animated.spring(scale, {
          toValue: 1,
          friction: 5,
          tension: 120,
          useNativeDriver: true,
        }),
      ]).start();
    }
  }, [focused, scale]);

  return (
    <View
      style={{
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        minHeight: 44,
        paddingHorizontal: focused ? 24 : spacing.xs,
        paddingVertical: spacing.xs,
        borderRadius: radius.pill,
        backgroundColor: focused ? colors.primary : "transparent",
      }}
    >
      <Animated.View
        style={{
          transform: [{ scale }],
          width: 28,
          height: 28,
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        <IconSymbol
          name={focused ? activeIcon : inactiveIcon}
          size={22}
          color={focused ? colors.onPrimary : colors.onSurface}
        />
      </Animated.View>

      {/* {focused && (
        <View style={{ marginTop: 6, alignItems: "center" }}>
          <Text
            style={[
              typography.labelMd,
              {
                color: colors.onPrimary,
                fontSize: 12,
              },
            ]}
          >
            {label}
          </Text>
        </View>
      )} */}
    </View>
  );
}

function AnimatedTabButton(props: any) {
  const scale = useRef(new Animated.Value(1)).current;

  const animateIn = () => {
    Animated.spring(scale, {
      toValue: 0.92,
      useNativeDriver: true,
    }).start();
  };

  const animateOut = () => {
    Animated.spring(scale, {
      toValue: 1,
      useNativeDriver: true,
    }).start();
  };

  const AnimatedPressable = Animated.createAnimatedComponent(Pressable);

  return (
    <AnimatedPressable
      {...props}
      onPressIn={animateIn}
      onPressOut={animateOut}
      style={[
        {
          flex: 1,
          alignItems: "center",
          justifyContent: "center",
        },
        { transform: [{ scale }] },
      ]}
    >
      {props.children}
    </AnimatedPressable>
  );
}

export default function TabLayout() {
  return (
    <Tabs
      screenOptions={{
        sceneStyle: { backgroundColor: colors.surface },

        headerStyle: { backgroundColor: colors.surfaceContainerLow },
        headerShadowVisible: false,
        headerTintColor: colors.onSurface,
        headerTitleStyle: {
          color: colors.onSurface,
          fontSize: 18,
          fontWeight: "700",
        },

        tabBarStyle: {
          backgroundColor: colors.primaryContainer,
          borderTopWidth: 0,
          height: 60,
          paddingHorizontal: spacing.md,
          paddingVertical: spacing.sm,
          borderRadius: radius.pill,
          marginHorizontal: spacing.sm,
          marginBottom: spacing.lg,
        },

        tabBarItemStyle: {
          alignItems: "center",
          justifyContent: "center",
          paddingTop: spacing.sm,
        },

        tabBarLabel: () => null,

        tabBarButton: AnimatedTabButton,
      }}
    >
      <Tabs.Screen
        name="home"
        options={{
          title: "Home",
          headerTitle: "Scan PDF",
          tabBarIcon: ({ focused }) =>
            TabIconWithLabel(focused, "Home", "house.fill", "house"),
        }}
      />

      <Tabs.Screen
        name="files"
        options={{
          title: "Files",
          tabBarIcon: ({ focused }) =>
            TabIconWithLabel(focused, "Files", "folder.fill", "folder"),
        }}
      />

      <Tabs.Screen
        name="scan"
        options={{
          title: "Scan",
          tabBarIcon: ({ focused }) =>
            TabIconWithLabel(focused, "Scan", "camera.fill", "camera"),
        }}
      />

      <Tabs.Screen
        name="settings"
        options={{
          title: "Settings",
          tabBarIcon: ({ focused }) =>
            TabIconWithLabel(
              focused,
              "Settings",
              "gearshape.fill",
              "gearshape",
            ),
        }}
      />
    </Tabs>
  );
}
