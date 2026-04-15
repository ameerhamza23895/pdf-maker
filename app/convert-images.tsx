import { ScreenSafeArea } from "@/src/components/ScreenSafeArea";
import { useEditImages } from "@/src/context/edit-images-context";
import { recordSavedFile } from "@/src/db/savedFileHistory";
import { setPendingEditDocument } from "@/src/navigation/pendingEditDocument";
import { electricCuratorTheme } from "@/src/theme/electric-curator";
import { buildPdfFromImageUris } from "@/src/utils/imagesToPdf";
import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import { useRouter } from "expo-router";
import React, { useEffect, useRef, useState } from "react";
import {
  ActivityIndicator,
  Alert,
  Dimensions,
  Image,
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

const SCROLL_THRESHOLD = 150;
const MAX_SCROLL_SPEED = 20;

export default function ConvertImagesPage() {
  const router = useRouter();
  const { images, setImages } = useEditImages();

  const scrollViewRef = useRef<ScrollView | null>(null);
  const scrollOffset = useRef(0);
  const scrollTimer = useRef<ReturnType<typeof setInterval> | null>(null);
  const lastMoveY = useRef<number | null>(null);

  const [gridData, setGridData] = useState<any[]>([]);
  const [isScrollEnabled, setIsScrollEnabled] = useState(true);
  const [generating, setGenerating] = useState(false);

  useEffect(() => {
    setGridData(images.map((img) => ({ ...img, key: img.id })));
  }, [images]);

  const startAutoScroll = () => {
    if (scrollTimer.current) return;

    scrollTimer.current = setInterval(() => {
      if (lastMoveY.current === null) return;

      let speed = 0;
      if (lastMoveY.current < SCROLL_THRESHOLD) {
        speed = -Math.min(
          MAX_SCROLL_SPEED,
          (SCROLL_THRESHOLD - lastMoveY.current) / 5,
        );
      } else if (lastMoveY.current > screenHeight - SCROLL_THRESHOLD) {
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
    }, 16);
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

  const handleGeneratePdf = async () => {
    if (gridData.length === 0) {
      Alert.alert(
        "No images",
        "Add images from Scan or Edit images before generating a PDF.",
      );
      return;
    }
    try {
      setGenerating(true);
      const uris = gridData.map(
        (item) => (item.processedUri || item.uri) as string,
      );
      const { outputUri, fileName } = await buildPdfFromImageUris(uris);
      void recordSavedFile({
        uri: outputUri,
        fileName,
        mimeType: "application/pdf",
        source: "images_to_pdf",
      });
      setPendingEditDocument({ uri: outputUri, name: fileName });
      router.push({
        pathname: "/edit-pdf",
        params: { backToHome: "1" },
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      Alert.alert(
        "Could not create PDF",
        `${msg}${msg.includes("embed") || msg.includes("Expected") ? " Try JPEG or PNG images." : ""}`,
      );
    } finally {
      setGenerating(false);
    }
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
    <ScreenSafeArea edges={["top", "left", "right", "bottom"]}>
      <View style={styles.container}>
        <View style={styles.header}>
          <Text style={styles.title}>Convert Images</Text>
          <Text style={styles.subtitle}>
            {gridData.length} images ready for PDF
          </Text>
        </View>

        <ScrollView
          ref={scrollViewRef}
          style={styles.gridScroll}
          scrollEnabled={isScrollEnabled}
          onScroll={(e) => {
            scrollOffset.current = e.nativeEvent.contentOffset.y;
          }}
          scrollEventThrottle={16}
          showsVerticalScrollIndicator
          contentContainerStyle={styles.gridContainer}
          nestedScrollEnabled
        >
          <View style={styles.gridWidth}>
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
          </View>
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
            style={[
              styles.button,
              (generating || gridData.length === 0) && styles.buttonDisabled,
            ]}
            onPress={() => void handleGeneratePdf()}
            disabled={generating || gridData.length === 0}
          >
            {generating ? (
              <ActivityIndicator color={colors.onPrimary} />
            ) : (
              <Text style={styles.buttonText}>Generate PDF</Text>
            )}
          </TouchableOpacity>
        </View>
      </View>
    </ScreenSafeArea>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, padding: spacing.md },
  header: { marginBottom: spacing.md },
  title: { ...typography.headlineMd, color: colors.onSurface },
  subtitle: { ...typography.bodyMd, color: colors.onSurface },
  gridScroll: {
    flex: 1,
    width: "100%",
  },
  gridContainer: {
    flexGrow: 1,
    paddingBottom: 120,
    width: "100%",
  },
  gridWidth: {
    width: "100%",
    alignSelf: "stretch",
  },
  draggableGrid: {
    width: "100%",
    backgroundColor: colors.surface,
  },
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
    minHeight: 52,
    justifyContent: "center",
  },
  buttonDisabled: {
    opacity: 0.55,
  },
  buttonText: { color: colors.onPrimary, fontWeight: "700", fontSize: 16 },
});
