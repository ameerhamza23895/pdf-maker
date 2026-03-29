import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import * as ImageManipulator from "expo-image-manipulator";
import { useRouter } from "expo-router";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
    ActivityIndicator,
    Image,
    LayoutChangeEvent,
    PanResponder,
    ScrollView,
    StyleSheet,
    Text,
    TouchableOpacity,
    View,
} from "react-native";

import { useEditImages } from "@/src/context/edit-images-context";
import { electricCuratorTheme } from "@/src/theme/electric-curator";

const { colors, spacing, radius, typography } = electricCuratorTheme;

const filterOptions = ["none", "mono", "warm", "cool"] as const;
type Mode = "crop" | "filter";
type HandleType =
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "topLeft"
  | "topRight"
  | "bottomLeft"
  | "bottomRight"
  | "move";

export default function EditImagesPage() {
  const router = useRouter();
  const { images, currentIndex, setCurrentIndex, updateImage, removeImage } =
    useEditImages();

  const [mode, setMode] = useState<Mode>("crop");
  const [previewLayout, setPreviewLayout] = useState({ width: 0, height: 0 });
  const [imgSize, setImgSize] = useState({ width: 1, height: 1 });
  const [isProcessing, setIsProcessing] = useState(false);
  const dragStartCrop = useRef({ top: 0, left: 0, right: 0, bottom: 0 });

  const currentImage = images[currentIndex];

  useEffect(() => {
    if (!currentImage) {
      router.replace("/scan");
    }
  }, [currentImage, router]);

  useEffect(() => {
    if (currentImage?.uri) {
      Image.getSize(currentImage.uri, (width, height) => {
        setImgSize({ width, height });
      });
    }
  }, [currentImage?.uri]);

  // Calculate actual visual size of the image within the container
  const containerW = previewLayout.width;
  const containerH = previewLayout.height;
  let visualW = containerW;
  let visualH = containerH;

  if (containerW > 0 && containerH > 0 && imgSize.width > 0) {
    const containerAspect = containerW / containerH;
    const imgAspect = imgSize.width / imgSize.height;
    if (imgAspect > containerAspect) {
      visualW = containerW;
      visualH = containerW / imgAspect;
    } else {
      visualH = containerH;
      visualW = containerH * imgAspect;
    }
  }

  const offsetX = (containerW - visualW) / 2;
  const offsetY = (containerH - visualH) / 2;

  const crop = useMemo(() => {
    const c = currentImage?.crop ?? { top: 0, left: 0, right: 0, bottom: 0 };
    return {
      top: Math.max(c.top, offsetY),
      bottom: Math.max(c.bottom, offsetY),
      left: Math.max(c.left, offsetX),
      right: Math.max(c.right, offsetX),
    };
  }, [currentImage, offsetX, offsetY]);

  useEffect(() => {
    if (
      mode !== "crop" ||
      !currentImage ||
      previewLayout.width === 0 ||
      previewLayout.height === 0
    ) {
      return;
    }

    const currentCrop = currentImage.crop;
    const isDefaultCrop =
      currentCrop.top === 0 &&
      currentCrop.left === 0 &&
      currentCrop.right === 0 &&
      currentCrop.bottom === 0;

    if (!isDefaultCrop) {
      return;
    }

    const side = Math.min(previewLayout.width, previewLayout.height) * 0.65;
    const left = (previewLayout.width - side) / 2;
    const top = (previewLayout.height - side) / 2;

    updateImage(currentImage.id, {
      crop: {
        top,
        left,
        right: previewLayout.width - left - side,
        bottom: previewLayout.height - top - side,
      },
    });
  }, [
    mode,
    currentImage,
    previewLayout.width,
    previewLayout.height,
    updateImage,
  ]);

  const clamp = (value: number, min: number, max: number) =>
    Math.min(Math.max(value, min), max);

  const createCropResponder = useCallback(
    (type: HandleType) =>
      PanResponder.create({
        onStartShouldSetPanResponder: () => mode === "crop",
        onMoveShouldSetPanResponder: () => mode === "crop",
        onStartShouldSetPanResponderCapture: () =>
          mode === "crop" && type !== "move",
        onMoveShouldSetPanResponderCapture: () =>
          mode === "crop" && type !== "move",
        onPanResponderTerminationRequest: () => false,
        onPanResponderGrant: () => {
          dragStartCrop.current = { ...crop };
        },
        onPanResponderMove: (_, gestureState) => {
          if (!currentImage || previewLayout.width === 0) return;

          let nextCrop = { ...dragStartCrop.current };
          const MIN_SIZE = 50; // Prevents the crop box from becoming too small

          if (type === "move") {
            const width = previewLayout.width - nextCrop.left - nextCrop.right;
            const height =
              previewLayout.height - nextCrop.top - nextCrop.bottom;

            nextCrop.left = clamp(
              dragStartCrop.current.left + gestureState.dx,
              offsetX,
              previewLayout.width - width - offsetX,
            );
            nextCrop.top = clamp(
              dragStartCrop.current.top + gestureState.dy,
              offsetY,
              previewLayout.height - height - offsetY,
            );
            nextCrop.right = previewLayout.width - nextCrop.left - width;
            nextCrop.bottom = previewLayout.height - nextCrop.top - height;
          } else {
            // Handle Edges & Corners
            if (type.includes("top")) {
              nextCrop.top = clamp(
                dragStartCrop.current.top + gestureState.dy,
                offsetY,
                previewLayout.height - nextCrop.bottom - MIN_SIZE,
              );
            }
            if (type.includes("bottom")) {
              nextCrop.bottom = clamp(
                dragStartCrop.current.bottom - gestureState.dy,
                offsetY,
                previewLayout.height - nextCrop.top - MIN_SIZE,
              );
            }
            if (type.includes("left")) {
              nextCrop.left = clamp(
                dragStartCrop.current.left + gestureState.dx,
                offsetX,
                previewLayout.width - nextCrop.right - MIN_SIZE,
              );
            }
            if (type.includes("right")) {
              nextCrop.right = clamp(
                dragStartCrop.current.right - gestureState.dx,
                offsetX,
                previewLayout.width - nextCrop.left - MIN_SIZE,
              );
            }
          }

          updateImage(currentImage.id, { crop: nextCrop });
        },
      }),
    [mode, currentImage, previewLayout, crop, updateImage, offsetX, offsetY],
  );

  // Responders for all handles
  const handlers = {
    move: useMemo(() => createCropResponder("move"), [createCropResponder]),
    top: useMemo(() => createCropResponder("top"), [createCropResponder]),
    bottom: useMemo(() => createCropResponder("bottom"), [createCropResponder]),
    left: useMemo(() => createCropResponder("left"), [createCropResponder]),
    right: useMemo(() => createCropResponder("right"), [createCropResponder]),
    topLeft: useMemo(
      () => createCropResponder("topLeft"),
      [createCropResponder],
    ),
    topRight: useMemo(
      () => createCropResponder("topRight"),
      [createCropResponder],
    ),
    bottomLeft: useMemo(
      () => createCropResponder("bottomLeft"),
      [createCropResponder],
    ),
    bottomRight: useMemo(
      () => createCropResponder("bottomRight"),
      [createCropResponder],
    ),
  };

  const onPreviewLayout = (event: LayoutChangeEvent) => {
    const { width, height } = event.nativeEvent.layout;
    if (width !== previewLayout.width || height !== previewLayout.height) {
      setPreviewLayout({ width, height });
    }
  };

  const removeCurrentImage = () => {
    if (!currentImage) return;
    removeImage(currentImage.id);
  };

  const rotateReal = async () => {
    if (!currentImage) return;
    try {
      const result = await ImageManipulator.manipulateAsync(currentImage.uri, [
        { rotate: 90 },
      ]);
      // Reset crop to boundaries when rotating!
      updateImage(currentImage.id, {
        uri: result.uri,
        crop: { top: 0, left: 0, right: 0, bottom: 0 },
      });
    } catch (e) {
      console.error(e);
    }
  };

  const processAndContinue = async () => {
    setIsProcessing(true);
    for (const img of images) {
      // Find the visual rect params that were used for THIS image.
      // Since calculating layout for all here is tough, we approximate back to actual image resolution
      // by fetching dimensions if not cached, but let's just use a simple ratio.
      const dims = await new Promise<{ width: number; height: number }>(
        (resolve, reject) => {
          Image.getSize(
            img.uri,
            (width, height) => resolve({ width, height }),
            reject,
          );
        },
      );

      // Determine visual size using previewLayout assuming layout didn't change wildly
      const imgAspect = dims.width / dims.height;
      const containerAspect = previewLayout.width / previewLayout.height;

      let vW = previewLayout.width;
      let vH = previewLayout.height;
      if (imgAspect > containerAspect) {
        vH = previewLayout.width / imgAspect;
      } else {
        vW = previewLayout.height * imgAspect;
      }

      const offX = (previewLayout.width - vW) / 2;
      const offY = (previewLayout.height - vH) / 2;

      // Ensure bounding fits
      const t = Math.max(img.crop.top, offY);
      const b = Math.max(img.crop.bottom, offY);
      const l = Math.max(img.crop.left, offX);
      const r = Math.max(img.crop.right, offX);

      // Map to original pixels
      const originX = ((l - offX) / vW) * dims.width;
      const originY = ((t - offY) / vH) * dims.height;
      const w = ((previewLayout.width - l - r) / vW) * dims.width;
      const h = ((previewLayout.height - t - b) / vH) * dims.height;

      // Avoid cropping if it's the full image
      if (Math.abs(w - dims.width) > 5 || Math.abs(h - dims.height) > 5) {
        try {
          const result = await ImageManipulator.manipulateAsync(img.uri, [
            {
              crop: {
                originX: Math.max(0, Math.floor(originX)),
                originY: Math.max(0, Math.floor(originY)),
                width: Math.max(1, Math.floor(w)),
                height: Math.max(1, Math.floor(h)),
              },
            },
          ]);
          updateImage(img.id, { processedUri: result.uri });
        } catch (e) {
          console.error("Crop error", e);
          updateImage(img.id, { processedUri: img.uri }); // Fallback
        }
      } else {
        updateImage(img.id, { processedUri: img.uri });
      }
    }
    setIsProcessing(false);
    router.push("/convert-images");
  };

  if (!currentImage) {
    return null;
  }

  const getFilterStyle = (filter: string) => {
    if (filter === "mono") return { tintColor: "#999999", opacity: 0.9 };
    if (filter === "warm") return { tintColor: "rgba(255, 180, 100, 0.35)" };
    if (filter === "cool") return { tintColor: "rgba(100, 180, 255, 0.25)" };
    return {};
  };

  const imageStyle = [styles.image, getFilterStyle(currentImage.filter)];

  return (
    <View style={styles.page}>
      <ScrollView
        scrollEnabled={mode !== "crop"}
        contentContainerStyle={styles.pageContent}
      >
        <Text style={styles.header}>Edit Images</Text>

        <View style={styles.thumbnailRow}>
          {images.map((img, idx) => (
            <TouchableOpacity
              key={img.id}
              style={[
                styles.thumbnailButton,
                idx === currentIndex && styles.thumbnailButtonActive,
              ]}
              onPress={() => setCurrentIndex(idx)}
            >
              <Image source={{ uri: img.uri }} style={styles.thumbnail} />
            </TouchableOpacity>
          ))}
        </View>

        <View style={styles.previewCard} onLayout={onPreviewLayout}>
          <Image source={{ uri: currentImage.uri }} style={imageStyle} />
          <TouchableOpacity
            style={styles.removeButton}
            onPress={removeCurrentImage}
            accessibilityLabel="Remove current image"
          >
            <MaterialIcons name="close" size={18} color="white" />
          </TouchableOpacity>

          {/* Masking out the cropped area for better UX */}
          {mode === "crop" && (
            <View style={StyleSheet.absoluteFill} pointerEvents="none">
              <View
                style={{
                  position: "absolute",
                  top: 0,
                  left: 0,
                  right: 0,
                  height: crop.top,
                  backgroundColor: "rgba(0,0,0,0.6)",
                }}
              />
              <View
                style={{
                  position: "absolute",
                  bottom: 0,
                  left: 0,
                  right: 0,
                  height: crop.bottom,
                  backgroundColor: "rgba(0,0,0,0.6)",
                }}
              />
              <View
                style={{
                  position: "absolute",
                  top: crop.top,
                  bottom: crop.bottom,
                  left: 0,
                  width: crop.left,
                  backgroundColor: "rgba(0,0,0,0.6)",
                }}
              />
              <View
                style={{
                  position: "absolute",
                  top: crop.top,
                  bottom: crop.bottom,
                  right: 0,
                  width: crop.right,
                  backgroundColor: "rgba(0,0,0,0.6)",
                }}
              />
            </View>
          )}

          {mode === "crop" && (
            <View
              style={[
                styles.cropFrame,
                {
                  top: crop.top,
                  left: crop.left,
                  right: crop.right,
                  bottom: crop.bottom,
                },
              ]}
              {...handlers.move.panHandlers}
            >
              {/* Grid Lines */}
              <View style={styles.gridLineHorizontal} />
              <View style={[styles.gridLineHorizontal, { top: "66%" }]} />
              <View style={styles.gridLineVertical} />
              <View style={[styles.gridLineVertical, { left: "66%" }]} />

              {/* Edge Handles */}
              <View
                style={[styles.edgeHandle, styles.edgeTop]}
                pointerEvents="box-only"
                {...handlers.top.panHandlers}
              />
              <View
                style={[styles.edgeHandle, styles.edgeBottom]}
                pointerEvents="box-only"
                {...handlers.bottom.panHandlers}
              />
              <View
                style={[styles.edgeHandle, styles.edgeLeft]}
                pointerEvents="box-only"
                {...handlers.left.panHandlers}
              />
              <View
                style={[styles.edgeHandle, styles.edgeRight]}
                pointerEvents="box-only"
                {...handlers.right.panHandlers}
              />

              {/* Corner Handles */}
              <View
                style={[styles.cornerHandle, styles.topLeft]}
                pointerEvents="box-only"
                {...handlers.topLeft.panHandlers}
              />
              <View
                style={[styles.cornerHandle, styles.topRight]}
                pointerEvents="box-only"
                {...handlers.topRight.panHandlers}
              />
              <View
                style={[styles.cornerHandle, styles.bottomLeft]}
                pointerEvents="box-only"
                {...handlers.bottomLeft.panHandlers}
              />
              <View
                style={[styles.cornerHandle, styles.bottomRight]}
                pointerEvents="box-only"
                {...handlers.bottomRight.panHandlers}
              />
            </View>
          )}
        </View>

        <View style={styles.actionRow}>
          <TouchableOpacity style={styles.actionButton} onPress={rotateReal}>
            <MaterialIcons
              name="rotate-right"
              size={24}
              color={colors.onSurface}
            />
          </TouchableOpacity>

          <TouchableOpacity
            style={[
              styles.actionButton,
              mode === "crop" && styles.actionButtonActive,
            ]}
            onPress={() => setMode("crop")}
          >
            <MaterialIcons
              name="crop-free"
              size={24}
              color={mode === "crop" ? colors.onPrimary : colors.onSurface}
            />
          </TouchableOpacity>

          <TouchableOpacity
            style={[
              styles.actionButton,
              mode === "filter" && styles.actionButtonActive,
            ]}
            onPress={() => setMode("filter")}
          >
            <MaterialIcons
              name="filter-alt"
              size={24}
              color={mode === "filter" ? colors.onPrimary : colors.onSurface}
            />
          </TouchableOpacity>
        </View>

        {mode === "filter" && (
          <ScrollView
            horizontal
            showsHorizontalScrollIndicator={false}
            contentContainerStyle={styles.filterRow}
          >
            {filterOptions.map((opt) => (
              <TouchableOpacity
                key={opt}
                style={[
                  styles.filterCard,
                  currentImage.filter === opt && styles.filterCardActive,
                ]}
                onPress={() => updateImage(currentImage.id, { filter: opt })}
              >
                <View style={styles.filterThumbnailContainer}>
                  <Image
                    source={{ uri: currentImage.uri }}
                    style={[styles.filterThumbnail, getFilterStyle(opt)]}
                  />
                </View>
                <Text
                  style={[
                    styles.filterText,
                    currentImage.filter === opt && styles.filterTextActive,
                  ]}
                >
                  {opt}
                </Text>
              </TouchableOpacity>
            ))}
          </ScrollView>
        )}

        <TouchableOpacity
          style={[styles.nextButton, isProcessing && { opacity: 0.7 }]}
          onPress={processAndContinue}
          disabled={isProcessing}
        >
          {isProcessing ? (
            <ActivityIndicator color={colors.onPrimary} />
          ) : (
            <Text style={styles.nextButtonText}>Continue to review</Text>
          )}
        </TouchableOpacity>
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  page: { flex: 1, backgroundColor: colors.surface },
  pageContent: { padding: spacing.md, gap: spacing.md, paddingBottom: 40 },
  header: { ...typography.headlineMd, marginBottom: spacing.sm },
  thumbnailRow: { flexDirection: "row", gap: spacing.sm },
  thumbnailButton: {
    borderRadius: radius.md,
    overflow: "hidden",
    borderWidth: 2,
    borderColor: "transparent",
  },
  thumbnailButtonActive: { borderColor: colors.primary },
  thumbnail: { width: 64, height: 64 },
  previewCard: {
    backgroundColor: "#111111",
    borderRadius: radius.md,
    height: 400,
    overflow: "hidden",
    position: "relative",
    borderWidth: 1,
    borderColor: "rgba(255,255,255,0.08)",
  },
  image: { width: "100%", height: "100%", resizeMode: "contain" },

  // Crop Frame UI
  cropFrame: {
    position: "absolute",
    borderWidth: 2,
    borderColor: "rgba(255,255,255,0.95)",
    borderStyle: "dashed",
    backgroundColor: "rgba(255,255,255,0.12)",
    zIndex: 10,
    overflow: "visible",
  },
  gridLineHorizontal: {
    position: "absolute",
    left: 0,
    right: 0,
    top: "33%",
    height: 1,
    backgroundColor: "rgba(255,255,255,0.4)",
  },
  gridLineVertical: {
    position: "absolute",
    top: 0,
    bottom: 0,
    left: "33%",
    width: 1,
    backgroundColor: "rgba(255,255,255,0.4)",
  },
  cornerHandle: {
    position: "absolute",
    width: 26,
    height: 26,
    backgroundColor: "white",
    borderRadius: 13,
    borderWidth: 2,
    borderColor: colors.primary,
    elevation: 4,
    shadowColor: "#000",
    shadowOpacity: 0.3,
    shadowRadius: 4,
    shadowOffset: { width: 0, height: 2 },
    zIndex: 11,
  },
  edgeHandle: {
    position: "absolute",
    backgroundColor: "rgba(255,255,255,0.45)",
    borderRadius: 15,
    zIndex: 11,
  },
  topLeft: { top: -12, left: -12 },
  topRight: { top: -12, right: -12 },
  bottomLeft: { bottom: -12, left: -12 },
  bottomRight: { bottom: -12, right: -12 },
  edgeTop: { top: -15, left: 20, right: 20, height: 30 },
  edgeBottom: { bottom: -15, left: 20, right: 20, height: 30 },
  edgeLeft: { left: -15, top: 20, bottom: 20, width: 30 },
  edgeRight: { right: -15, top: 20, bottom: 20, width: 30 },

  actionRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    gap: spacing.sm,
  },
  actionButton: {
    flex: 1,
    paddingVertical: spacing.md,
    borderRadius: radius.md,
    backgroundColor: colors.surfaceContainerLow,
    alignItems: "center",
  },
  actionButtonActive: { backgroundColor: colors.primary },
  filterRow: {
    flexDirection: "row",
    gap: spacing.md,
    paddingVertical: spacing.sm,
  },
  filterCard: {
    alignItems: "center",
    gap: spacing.xs,
    width: 72,
  },
  filterCardActive: {},
  filterThumbnailContainer: {
    width: 72,
    height: 72,
    borderRadius: radius.md,
    overflow: "hidden",
    borderWidth: 2,
    borderColor: "transparent",
  },
  filterThumbnail: {
    width: "100%",
    height: "100%",
    resizeMode: "cover",
  },
  filterText: {
    color: colors.onSurface,
    fontSize: 13,
    textTransform: "capitalize",
    fontWeight: "500",
  },
  filterTextActive: { color: colors.primary, fontWeight: "700" },
  nextButton: {
    marginTop: spacing.lg,
    paddingVertical: spacing.md,
    borderRadius: radius.pill,
    backgroundColor: colors.primary,
    alignItems: "center",
  },
  nextButtonText: { color: colors.onPrimary, fontWeight: "700" },
  emptyContainer: { flex: 1, alignItems: "center", justifyContent: "center" },
  emptyText: { color: colors.onSurface, marginBottom: spacing.md },
  removeButton: {
    position: "absolute",
    top: 12,
    right: 12,
    width: 32,
    height: 32,
    borderRadius: 16,
    backgroundColor: "rgba(0,0,0,0.6)",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 12,
  },
  linkButton: {
    padding: spacing.md,
    backgroundColor: colors.primary,
    borderRadius: radius.pill,
  },
  linkText: { color: colors.onPrimary, fontWeight: "700" },
});
