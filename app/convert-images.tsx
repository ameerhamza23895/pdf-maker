import { useEditImages } from "@/src/context/edit-images-context";
import { electricCuratorTheme } from "@/src/theme/electric-curator";
import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import { useRouter } from "expo-router";
import React, { useEffect, useRef, useState } from "react";
import {
  Dimensions,
  Image,
  SafeAreaView,
  ScrollView,
  StyleSheet,
  Text,
  TouchableOpacity,
  View,
} from "react-native";
import { DraggableGrid } from "react-native-draggable-grid";

const { colors, spacing, radius, typography } = electricCuratorTheme;
const screenWidth = Dimensions.get("window").width;
const screenHeight = Dimensions.get("window").height;
const imageSize = (screenWidth - spacing.md * 2) / 3 - 8;

// Thresholds for auto-scroll
const SCROLL_THRESHOLD = 150;
const MAX_SCROLL_SPEED = 20;

export default function ConvertImagesPage() {
  const router = useRouter();
  const { images, setImages } = useEditImages();

  const scrollViewRef = useRef<ScrollView | null>(null);
  const scrollOffset = useRef(0);
  const scrollTimer = useRef<NodeJS.Timeout | null>(null);
  const lastMoveY = useRef<number | null>(null);

  const [gridData, setGridData] = useState<any[]>([]);
  const [isScrollEnabled, setIsScrollEnabled] = useState(true);

  useEffect(() => {
    setGridData(images.map((img) => ({ ...img, key: img.id })));
  }, [images]);

  // The "Smooth Engine": This runs independently of finger movement
  const startAutoScroll = () => {
    if (scrollTimer.current) return;

    scrollTimer.current = setInterval(() => {
      if (lastMoveY.current === null) return;

      let speed = 0;
      if (lastMoveY.current < SCROLL_THRESHOLD) {
        // Calculate variable speed: faster when closer to the top edge
        speed = -Math.min(
          MAX_SCROLL_SPEED,
          (SCROLL_THRESHOLD - lastMoveY.current) / 5,
        );
      } else if (lastMoveY.current > screenHeight - SCROLL_THRESHOLD) {
        // Faster when closer to the bottom edge
        speed = Math.min(
          MAX_SCROLL_SPEED,
          (lastMoveY.current - (screenHeight - SCROLL_THRESHOLD)) / 5,
        );
      }

      if (speed !== 0) {
        const nextScroll = scrollOffset.current + speed;
        scrollViewRef.current?.scrollTo({
          y: Math.max(0, nextScroll),
          animated: false,
        });
      }
    }, 16); // ~60fps smooth loop
  };

  const stopAutoScroll = () => {
    if (scrollTimer.current) {
      clearInterval(scrollTimer.current);
      scrollTimer.current = null;
    }
    lastMoveY.current = null;
  };

  const onDragStart = () => {
    setIsScrollEnabled(false);
    startAutoScroll();
  };

  const handleDragging = (gestureState: { moveY: number }) => {
    lastMoveY.current = gestureState.moveY;
  };

  const handleDragRelease = (newData: any[]) => {
    stopAutoScroll();
    setIsScrollEnabled(true);
    setGridData(newData);
    setImages(newData);
  };

  const renderItem = (item: any) => (
    <View
      key={item.key}
      style={[styles.imageCard, { width: imageSize, height: imageSize }]}
    >
      <Image
        source={{ uri: item.processedUri || item.uri }}
        style={styles.image}
      />
    </View>
  );

  return (
    <SafeAreaView style={styles.page}>
      <View style={styles.container}>
        <View style={styles.header}>
          <Text style={styles.title}>Convert Images</Text>
          <Text style={styles.subtitle}>
            {gridData.length} images ready for PDF
          </Text>
        </View>

        <ScrollView
          ref={scrollViewRef}
          scrollEnabled={isScrollEnabled}
          onScroll={(e) => {
            scrollOffset.current = e.nativeEvent.contentOffset.y;
          }}
          scrollEventThrottle={16}
          contentContainerStyle={styles.gridContainer}
        >
          <DraggableGrid
            numColumns={3}
            data={gridData}
            renderItem={renderItem}
            onDragStart={onDragStart}
            onDragRelease={handleDragRelease}
            onDragging={handleDragging}
            itemHeight={imageSize + 8}
            style={styles.draggableGrid}
          />
        </ScrollView>

        <View style={styles.footer}>
          <TouchableOpacity
            style={styles.addCard}
            onPress={() => router.push("/scan")}
          >
            <MaterialIcons
              name="add-a-photo"
              size={22}
              color={colors.primary}
            />
            <Text style={styles.addText}>Add More</Text>
          </TouchableOpacity>

          <TouchableOpacity
            style={styles.button}
            onPress={() => console.log("Finalizing...")}
          >
            <Text style={styles.buttonText}>Generate Document</Text>
          </TouchableOpacity>
        </View>
      </View>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  page: { flex: 1, backgroundColor: colors.surface },
  container: { flex: 1, padding: spacing.md },
  header: { marginBottom: spacing.md },
  title: { ...typography.headlineMd, color: colors.onSurface },
  subtitle: { ...typography.bodyMedium, color: colors.onSurfaceVariant },
  gridContainer: { paddingBottom: 120 },
  draggableGrid: { backgroundColor: colors.surface },
  imageCard: {
    borderRadius: radius.md,
    overflow: "hidden",
    backgroundColor: colors.surfaceContainerLow,
    elevation: 4,
    shadowColor: "#000",
    shadowOpacity: 0.1,
    shadowRadius: 4,
  },
  image: { width: "100%", height: "100%", resizeMode: "cover" },
  footer: { gap: spacing.sm, paddingTop: spacing.md },
  addCard: {
    flexDirection: "row",
    height: 52,
    borderRadius: radius.md,
    borderWidth: 1,
    borderStyle: "dashed",
    borderColor: colors.outlineVariant,
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: colors.surfaceContainerLow,
    gap: spacing.sm,
  },
  addText: { color: colors.primary, fontWeight: "600" },
  button: {
    paddingVertical: spacing.md,
    borderRadius: radius.pill,
    backgroundColor: colors.primary,
    alignItems: "center",
  },
  buttonText: { color: colors.onPrimary, fontWeight: "700", fontSize: 16 },
});
