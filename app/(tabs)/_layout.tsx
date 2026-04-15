import { Tabs } from "expo-router";
import type { ReactNode } from "react";
import { useEffect, useRef } from "react";
import {
  Animated,
  type GestureResponderEvent,
  Pressable,
  type PressableProps,
  type PressableStateCallbackType,
  View,
} from "react-native";

import { IS_SETTING_SHOW } from "@/ENV";
import { IconSymbol, type IconSymbolName } from "@/src/components/icon-symbol";
import { electricCuratorTheme } from "@/src/theme/electric-curator";

const { colors, radius, spacing } = electricCuratorTheme;

function TabIconWithLabel({
  focused,
  label,
  activeIcon,
  inactiveIcon,
}: {
  focused: boolean;
  label: string;
  activeIcon: IconSymbolName;
  inactiveIcon: IconSymbolName;
}) {
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

/**
 * IMPORTANT: keep the `style` prop coming from React Navigation.
 * Overriding it causes "ghost" tab spacing (e.g. looks like 4 slots when only 3 tabs exist).
 */
function AnimatedTabButton({
  style,
  children,
  onPressIn,
  onPressOut,
  ...rest
}: PressableProps) {
  const scale = useRef(new Animated.Value(1)).current;

  const handlePressIn = (e: GestureResponderEvent) => {
    Animated.spring(scale, { toValue: 0.92, useNativeDriver: true }).start();
    onPressIn?.(e);
  };

  const handlePressOut = (e: GestureResponderEvent) => {
    Animated.spring(scale, { toValue: 1, useNativeDriver: true }).start();
    onPressOut?.(e);
  };

  const innerStyle = [
    {
      flex: 1,
      alignItems: "center" as const,
      justifyContent: "center" as const,
      width: "100%" as const,
    },
    { transform: [{ scale }] },
  ];

  const wrap = (node: ReactNode) => (
    <Animated.View style={innerStyle}>{node}</Animated.View>
  );

  if (typeof children === "function") {
    const renderChild = children as (
      state: PressableStateCallbackType,
    ) => ReactNode;
    return (
      <Pressable
        {...rest}
        onPressIn={handlePressIn}
        onPressOut={handlePressOut}
        style={style}
      >
        {(state) => wrap(renderChild(state))}
      </Pressable>
    );
  }

  return (
    <Pressable
      {...rest}
      onPressIn={handlePressIn}
      onPressOut={handlePressOut}
      style={style}
    >
      {wrap(children)}
    </Pressable>
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

        // IMPORTANT: pass a component element so hooks inside are valid.
        tabBarButton: (props) => <AnimatedTabButton {...props} />,
      }}
    >
      <Tabs.Screen
        name="home"
        options={{
          title: "Home",
          headerTitle: "Scan PDF",
          tabBarIcon: ({ focused }) =>
            (
              <TabIconWithLabel
                focused={focused}
                label="Home"
                activeIcon="house.fill"
                inactiveIcon="house"
              />
            ),
        }}
      />

      <Tabs.Screen
        name="files"
        options={{
          title: "Files",
          tabBarIcon: ({ focused }) =>
            (
              <TabIconWithLabel
                focused={focused}
                label="Files"
                activeIcon="folder.fill"
                inactiveIcon="folder"
              />
            ),
        }}
      />

      <Tabs.Screen
        name="scan"
        options={{
          title: "Scan",
          tabBarIcon: ({ focused }) =>
            (
              <TabIconWithLabel
                focused={focused}
                label="Scan"
                activeIcon="camera.fill"
                inactiveIcon="camera"
              />
            ),
        }}
      />

      <Tabs.Screen
        name="settings"
        options={
          IS_SETTING_SHOW
            ? {
                title: "Settings",
                tabBarIcon: ({ focused }) =>
                  (
                    <TabIconWithLabel
                      focused={focused}
                      label="Settings"
                      activeIcon="gearshape.fill"
                      inactiveIcon="gearshape"
                    />
                  ),
              }
            : {
                // Expo Router will still register the route because the file exists.
                // This removes it from the tab bar (no slot, not clickable) while
                // keeping it navigable via `router.push("/settings")` if needed.
                href: null,
              }
        }
      />
    </Tabs>
  );
}
