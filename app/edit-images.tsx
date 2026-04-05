import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import * as ImageManipulator from "expo-image-manipulator";
import { Stack, useRouter } from "expo-router";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type Dispatch,
  type MutableRefObject,
  type SetStateAction,
} from "react";
import {
    ActivityIndicator,
    Image,
    LayoutChangeEvent,
    PanResponder,
    ScrollView,
    StyleSheet,
    Text,
    TouchableOpacity,
    useWindowDimensions,
    View,
} from "react-native";
import { useSafeAreaInsets } from "react-native-safe-area-context";

import {
    IMAGE_FILTER_PRESETS,
    type ImageFilterId,
} from "@/src/constants/imageFilterPresets";
import {
    EditableImage,
    useEditImages,
} from "@/src/context/edit-images-context";
import { electricCuratorTheme, withAlpha } from "@/src/theme/electric-curator";
import {
  clampCropToLetterbox,
  clampNumber,
} from "@/src/utils/cropRect";

const { colors, spacing, radius } = electricCuratorTheme;

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

const MIN_CROP_PX = 48;
const MOVE_ACTIVATION_PX = 12;

type CropInsetState = {
  top: number;
  left: number;
  right: number;
  bottom: number;
};

type CropGestureRefs = {
  modeRef: MutableRefObject<Mode>;
  cropRef: MutableRefObject<CropInsetState>;
  layoutRef: MutableRefObject<{
    cw: number;
    ch: number;
    ox: number;
    oy: number;
    vw: number;
    vh: number;
  }>;
  dragStartCrop: MutableRefObject<CropInsetState>;
  imageIdRef: MutableRefObject<string>;
  updateImageRef: MutableRefObject<
    (id: string, patch: Partial<Omit<EditableImage, "id">>) => void
  >;
  setCropGestureActive: Dispatch<SetStateAction<boolean>>;
};

function buildCropResponder(type: HandleType, r: CropGestureRefs) {
  const endGesture = () => {
    r.setCropGestureActive(false);
  };

  return PanResponder.create({
    onStartShouldSetPanResponder: () =>
      r.modeRef.current === "crop" && type !== "move",
    onMoveShouldSetPanResponder: (_, g) => {
      if (r.modeRef.current !== "crop") return false;
      if (type === "move") {
        return (
          Math.abs(g.dx) > MOVE_ACTIVATION_PX ||
          Math.abs(g.dy) > MOVE_ACTIVATION_PX
        );
      }
      return false;
    },
    onStartShouldSetPanResponderCapture: () =>
      r.modeRef.current === "crop" && type !== "move",
    onMoveShouldSetPanResponderCapture: () =>
      r.modeRef.current === "crop" && type !== "move",
    onPanResponderTerminationRequest: () => false,
    onPanResponderGrant: () => {
      r.dragStartCrop.current = { ...r.cropRef.current };
      r.setCropGestureActive(true);
    },
    onPanResponderMove: (_, gestureState) => {
      const id = r.imageIdRef.current;
      if (!id) return;

      const { cw, ch, ox, oy, vw, vh } = r.layoutRef.current;
      if (cw === 0 || vw <= 0 || vh <= 0) return;

      let nextCrop: CropInsetState = { ...r.dragStartCrop.current };

      if (type === "move") {
        const w = cw - r.dragStartCrop.current.left - r.dragStartCrop.current.right;
        const h = ch - r.dragStartCrop.current.top - r.dragStartCrop.current.bottom;
        const nextLeft = clampNumber(
          r.dragStartCrop.current.left + gestureState.dx,
          ox,
          ox + vw - w,
        );
        const nextTop = clampNumber(
          r.dragStartCrop.current.top + gestureState.dy,
          oy,
          oy + vh - h,
        );
        nextCrop = {
          left: nextLeft,
          top: nextTop,
          right: cw - nextLeft - w,
          bottom: ch - nextTop - h,
        };
      } else {
        if (type.includes("top")) {
          nextCrop.top = r.dragStartCrop.current.top + gestureState.dy;
        }
        if (type.includes("bottom")) {
          nextCrop.bottom = r.dragStartCrop.current.bottom - gestureState.dy;
        }
        if (type.includes("left")) {
          nextCrop.left = r.dragStartCrop.current.left + gestureState.dx;
        }
        if (type.includes("right")) {
          nextCrop.right = r.dragStartCrop.current.right - gestureState.dx;
        }
        nextCrop = clampCropToLetterbox(
          nextCrop,
          cw,
          ch,
          ox,
          oy,
          vw,
          vh,
          MIN_CROP_PX,
        );
      }

      r.updateImageRef.current(id, { crop: nextCrop });
    },
    onPanResponderRelease: endGesture,
    onPanResponderTerminate: endGesture,
  });
}

const THUMB_SIZE = 68;
const THUMB_GAP = 10;

export default function EditImagesPage() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const { width: windowWidth } = useWindowDimensions();
  const thumbScrollRef = useRef<ScrollView>(null);
  const {
    batchImages,
    currentIndex,
    setCurrentIndex,
    updateImage,
    removeImage,
    commitBatchImages,
  } = useEditImages();

  const [mode, setMode] = useState<Mode>("crop");
  const [previewLayout, setPreviewLayout] = useState({ width: 0, height: 0 });
  const [imgSize, setImgSize] = useState({ width: 1, height: 1 });
  const [isProcessing, setIsProcessing] = useState(false);
  const skipScanRedirect = useRef(false);
  const dragStartCrop = useRef({ top: 0, left: 0, right: 0, bottom: 0 });
  const cropInitializedIds = useRef(new Set<string>());
  const [cropGestureActive, setCropGestureActive] = useState(false);

  const modeRef = useRef(mode);
  const cropRef = useRef<CropInsetState>({
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
  });
  const layoutRef = useRef({
    cw: 0,
    ch: 0,
    ox: 0,
    oy: 0,
    vw: 0,
    vh: 0,
  });
  const imageIdRef = useRef("");
  const updateImageRef = useRef(updateImage);
  const cropPanHandlersRef = useRef<{
    move: ReturnType<typeof PanResponder.create>;
    top: ReturnType<typeof PanResponder.create>;
    bottom: ReturnType<typeof PanResponder.create>;
    left: ReturnType<typeof PanResponder.create>;
    right: ReturnType<typeof PanResponder.create>;
    topLeft: ReturnType<typeof PanResponder.create>;
    topRight: ReturnType<typeof PanResponder.create>;
    bottomLeft: ReturnType<typeof PanResponder.create>;
    bottomRight: ReturnType<typeof PanResponder.create>;
  } | null>(null);

  const currentImage = batchImages[currentIndex];

  useEffect(() => {
    if (!currentImage && !skipScanRedirect.current) {
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
    return { ...c };
  }, [currentImage?.crop]);

  modeRef.current = mode;
  cropRef.current = crop;
  layoutRef.current = {
    cw: containerW,
    ch: containerH,
    ox: offsetX,
    oy: offsetY,
    vw: visualW,
    vh: visualH,
  };
  imageIdRef.current = currentImage?.id ?? "";
  updateImageRef.current = updateImage;

  if (!cropPanHandlersRef.current) {
    const r: CropGestureRefs = {
      modeRef,
      cropRef,
      layoutRef,
      dragStartCrop,
      imageIdRef,
      updateImageRef,
      setCropGestureActive,
    };
    cropPanHandlersRef.current = {
      move: buildCropResponder("move", r),
      top: buildCropResponder("top", r),
      bottom: buildCropResponder("bottom", r),
      left: buildCropResponder("left", r),
      right: buildCropResponder("right", r),
      topLeft: buildCropResponder("topLeft", r),
      topRight: buildCropResponder("topRight", r),
      bottomLeft: buildCropResponder("bottomLeft", r),
      bottomRight: buildCropResponder("bottomRight", r),
    };
  }
  const handlers = cropPanHandlersRef.current;

  /** One-time default crop inside the letterboxed image area so each image has its own box. */
  useEffect(() => {
    if (
      mode !== "crop" ||
      !currentImage ||
      previewLayout.width === 0 ||
      previewLayout.height === 0 ||
      imgSize.width <= 0 ||
      visualW <= 0
    ) {
      return;
    }

    if (cropInitializedIds.current.has(currentImage.id)) {
      return;
    }

    const c = currentImage.crop;
    const isDefaultCrop =
      c.top === 0 && c.left === 0 && c.right === 0 && c.bottom === 0;
    if (!isDefaultCrop) {
      cropInitializedIds.current.add(currentImage.id);
      return;
    }

    const containerW = previewLayout.width;
    const containerH = previewLayout.height;
    const side = Math.min(visualW, visualH) * 0.72;
    const left = offsetX + (visualW - side) / 2;
    const top = offsetY + (visualH - side) / 2;

    updateImage(currentImage.id, {
      crop: {
        top,
        left,
        right: containerW - left - side,
        bottom: containerH - top - side,
      },
    });
    cropInitializedIds.current.add(currentImage.id);
  }, [
    mode,
    currentImage,
    previewLayout.width,
    previewLayout.height,
    imgSize.width,
    imgSize.height,
    visualW,
    visualH,
    offsetX,
    offsetY,
    updateImage,
  ]);

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
      cropInitializedIds.current.delete(currentImage.id);
      updateImage(currentImage.id, {
        uri: result.uri,
        crop: { top: 0, left: 0, right: 0, bottom: 0 },
      });
    } catch (e) {
      console.error(e);
    }
  };

  const flipImage = async () => {
    if (!currentImage) return;
    try {
      const result = await ImageManipulator.manipulateAsync(currentImage.uri, [
        { flip: ImageManipulator.FlipType.Horizontal },
      ]);
      cropInitializedIds.current.delete(currentImage.id);
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
    skipScanRedirect.current = true;
    const processedBatch: EditableImage[] = [];

    for (const img of batchImages) {
      const dims = await new Promise<{ width: number; height: number }>(
        (resolve, reject) => {
          Image.getSize(
            img.uri,
            (width, height) => resolve({ width, height }),
            reject,
          );
        },
      );

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

      const t = Math.max(img.crop.top, offY);
      const b = Math.max(img.crop.bottom, offY);
      const l = Math.max(img.crop.left, offX);
      const r = Math.max(img.crop.right, offX);

      const originX = ((l - offX) / vW) * dims.width;
      const originY = ((t - offY) / vH) * dims.height;
      const w = ((previewLayout.width - l - r) / vW) * dims.width;
      const h = ((previewLayout.height - t - b) / vH) * dims.height;

      let processedUri = img.uri;

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
          processedUri = result.uri;
        } catch (e) {
          console.error("Crop error", e);
          processedUri = img.uri;
        }
      }

      processedBatch.push({
        ...img,
        processedUri,
      });
    }

    commitBatchImages(processedBatch);
    setIsProcessing(false);
    router.push("/convert-images");
  };

  const activePreset = useMemo(
    () =>
      IMAGE_FILTER_PRESETS.find(
        (p) => p.id === (currentImage?.filter ?? "none"),
      ) ?? IMAGE_FILTER_PRESETS[0],
    [currentImage?.filter],
  );

  useEffect(() => {
    if (batchImages.length === 0) {
      return;
    }
    const itemStride = THUMB_SIZE + THUMB_GAP;
    const viewport = windowWidth - spacing.md * 2;
    const centerOffset = Math.max(
      0,
      currentIndex * itemStride - viewport / 2 + THUMB_SIZE / 2,
    );
    thumbScrollRef.current?.scrollTo({
      x: centerOffset,
      animated: true,
    });
  }, [currentIndex, windowWidth, batchImages.length]);

  if (!currentImage) {
    return null;
  }

  return (
    <View style={[styles.page, { paddingBottom: insets.bottom }]}>
      <Stack.Screen
        options={{
          title: "Edit images",
          headerShown: true,
          headerStyle: { backgroundColor: colors.surfaceContainerLow },
          headerTintColor: colors.onSurface,
          headerTitleStyle: { fontWeight: "700" },
          headerBackTitle: "Back",
        }}
      />

      <View style={styles.thumbStripOuter}>
        <ScrollView
          ref={thumbScrollRef}
          horizontal
          showsHorizontalScrollIndicator={false}
          decelerationRate="fast"
          keyboardShouldPersistTaps="handled"
          contentContainerStyle={styles.thumbStripContent}
        >
          {batchImages.map((img, idx) => {
            const thumbPreset =
              IMAGE_FILTER_PRESETS.find((p) => p.id === img.filter) ??
              IMAGE_FILTER_PRESETS[0];
            return (
              <TouchableOpacity
                key={img.id}
                style={[
                  styles.thumbnailButton,
                  {
                    width: THUMB_SIZE,
                    marginRight: idx === batchImages.length - 1 ? 0 : THUMB_GAP,
                  },
                  idx === currentIndex && styles.thumbnailButtonActive,
                ]}
                onPress={() => setCurrentIndex(idx)}
                activeOpacity={0.85}
              >
                <View
                  style={[
                    styles.thumbnailWrap,
                    { width: THUMB_SIZE, height: THUMB_SIZE },
                  ]}
                >
                  <Image
                    source={{ uri: img.uri }}
                    style={[
                      styles.thumbnail,
                      { width: THUMB_SIZE, height: THUMB_SIZE },
                    ]}
                  />
                  {thumbPreset.preview.overlayColor ? (
                    <View
                      pointerEvents="none"
                      style={[
                        StyleSheet.absoluteFill,
                        {
                          backgroundColor: thumbPreset.preview.overlayColor,
                          opacity: thumbPreset.preview.overlayOpacity ?? 0.2,
                        },
                      ]}
                    />
                  ) : null}
                </View>
              </TouchableOpacity>
            );
          })}
        </ScrollView>
      </View>

      <View style={styles.previewShell}>
        <View style={styles.previewCard} onLayout={onPreviewLayout}>
          <Image
            key={currentImage.id + currentImage.uri}
            source={{ uri: currentImage.uri }}
            style={[styles.image, activePreset.preview.imageStyle as object]}
          />
          {activePreset.preview.overlayColor ? (
            <View
              pointerEvents="none"
              style={[
                StyleSheet.absoluteFill,
                {
                  backgroundColor: activePreset.preview.overlayColor,
                  opacity: activePreset.preview.overlayOpacity ?? 0.2,
                },
              ]}
            />
          ) : null}
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
                  backgroundColor: "rgba(0,0,0,0.52)",
                }}
              />
              <View
                style={{
                  position: "absolute",
                  bottom: 0,
                  left: 0,
                  right: 0,
                  height: crop.bottom,
                  backgroundColor: "rgba(0,0,0,0.52)",
                }}
              />
              <View
                style={{
                  position: "absolute",
                  top: crop.top,
                  bottom: crop.bottom,
                  left: 0,
                  width: crop.left,
                  backgroundColor: "rgba(0,0,0,0.52)",
                }}
              />
              <View
                style={{
                  position: "absolute",
                  top: crop.top,
                  bottom: crop.bottom,
                  right: 0,
                  width: crop.right,
                  backgroundColor: "rgba(0,0,0,0.52)",
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
              pointerEvents="box-none"
            >
              <View
                style={styles.cropDragPlane}
                {...handlers.move.panHandlers}
              />
              <View style={styles.cropGridOverlay} pointerEvents="none">
                <View style={styles.gridLineHorizontal} />
                <View style={[styles.gridLineHorizontal, { top: "66%" }]} />
                <View style={styles.gridLineVertical} />
                <View style={[styles.gridLineVertical, { left: "66%" }]} />
              </View>

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
      </View>

      <ScrollView
        horizontal
        nestedScrollEnabled
        showsHorizontalScrollIndicator={false}
        keyboardShouldPersistTaps="handled"
        scrollEnabled={!cropGestureActive}
        contentContainerStyle={styles.toolStripContent}
      >
        <TouchableOpacity
          style={styles.toolChip}
          onPress={rotateReal}
          accessibilityLabel="Rotate 90 degrees"
        >
          <MaterialIcons
            name="rotate-right"
            size={22}
            color={colors.onSurface}
          />
        </TouchableOpacity>
        <TouchableOpacity
          style={styles.toolChip}
          onPress={flipImage}
          accessibilityLabel="Flip horizontally"
        >
          <MaterialIcons name="flip" size={22} color={colors.onSurface} />
        </TouchableOpacity>
        <TouchableOpacity
          style={[
            styles.toolChip,
            mode === "crop" && styles.toolChipActive,
          ]}
          onPress={() => setMode("crop")}
          accessibilityLabel="Crop"
        >
          <MaterialIcons
            name="crop-free"
            size={22}
            color={mode === "crop" ? colors.onPrimary : colors.onSurface}
          />
        </TouchableOpacity>
        <View style={styles.toolStripDivider} />
        {IMAGE_FILTER_PRESETS.map((preset) => {
          const active = currentImage.filter === preset.id;
          return (
            <TouchableOpacity
              key={preset.id}
              style={[
                styles.filterIconBtn,
                active && styles.filterIconBtnActive,
              ]}
              onPress={() => {
                updateImage(currentImage.id, {
                  filter: preset.id as ImageFilterId,
                });
                setMode("filter");
              }}
              accessibilityLabel={preset.label}
              accessibilityState={{ selected: active }}
            >
              <MaterialIcons
                name={preset.materialIcon as never}
                size={22}
                color={active ? colors.onPrimary : colors.onSurface}
              />
            </TouchableOpacity>
          );
        })}
      </ScrollView>

      <TouchableOpacity
        style={[
          styles.nextButton,
          isProcessing && { opacity: 0.7 },
          { marginHorizontal: spacing.md, marginBottom: spacing.sm },
        ]}
        onPress={processAndContinue}
        disabled={isProcessing}
      >
        {isProcessing ? (
          <ActivityIndicator color={colors.onPrimary} />
        ) : (
          <Text style={styles.nextButtonText}>Continue to review</Text>
        )}
      </TouchableOpacity>
    </View>
  );
}

const styles = StyleSheet.create({
  page: { flex: 1, backgroundColor: colors.surface },
  thumbStripOuter: {
    paddingTop: spacing.xs,
    paddingBottom: spacing.sm,
    paddingHorizontal: spacing.md,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: withAlpha(colors.outlineVariant, 0.85),
  },
  thumbStripContent: {
    flexDirection: "row",
    alignItems: "center",
    paddingVertical: spacing.xs,
  },
  previewShell: {
    flex: 1,
    minHeight: 0,
    marginHorizontal: spacing.md,
    marginTop: spacing.sm,
  },
  thumbnailWrap: {
    borderRadius: radius.sm,
    overflow: "hidden",
    position: "relative",
  },
  thumbnailButton: {
    borderRadius: radius.md,
    overflow: "hidden",
    borderWidth: 2,
    borderColor: "transparent",
  },
  thumbnailButtonActive: {
    borderColor: colors.primary,
    backgroundColor: withAlpha(colors.primary, 0.06),
  },
  thumbnail: { resizeMode: "cover" },
  previewCard: {
    flex: 1,
    minHeight: 220,
    backgroundColor: "#111111",
    borderRadius: radius.md,
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
    borderColor: "rgba(255,255,255,0.92)",
    borderStyle: "solid",
    backgroundColor: "transparent",
    zIndex: 10,
    overflow: "visible",
    shadowColor: "#000",
    shadowOpacity: 0.45,
    shadowRadius: 6,
    shadowOffset: { width: 0, height: 0 },
    elevation: 6,
  },
  cropDragPlane: {
    ...StyleSheet.absoluteFillObject,
    zIndex: 5,
  },
  cropGridOverlay: {
    ...StyleSheet.absoluteFillObject,
    zIndex: 6,
  },
  gridLineHorizontal: {
    position: "absolute",
    left: 0,
    right: 0,
    top: "33%",
    height: StyleSheet.hairlineWidth,
    backgroundColor: "rgba(255,255,255,0.38)",
  },
  gridLineVertical: {
    position: "absolute",
    top: 0,
    bottom: 0,
    left: "33%",
    width: StyleSheet.hairlineWidth,
    backgroundColor: "rgba(255,255,255,0.38)",
  },
  cornerHandle: {
    position: "absolute",
    width: 32,
    height: 32,
    backgroundColor: "white",
    borderRadius: 16,
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
    backgroundColor: "rgba(255,255,255,0.5)",
    borderRadius: 18,
    zIndex: 11,
  },
  topLeft: { top: -15, left: -15 },
  topRight: { top: -15, right: -15 },
  bottomLeft: { bottom: -15, left: -15 },
  bottomRight: { bottom: -15, right: -15 },
  edgeTop: { top: -18, left: 24, right: 24, height: 36 },
  edgeBottom: { bottom: -18, left: 24, right: 24, height: 36 },
  edgeLeft: { left: -18, top: 24, bottom: 24, width: 36 },
  edgeRight: { right: -18, top: 24, bottom: 24, width: 36 },

  toolStripContent: {
    flexDirection: "row",
    alignItems: "center",
    gap: 8,
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.md,
    minHeight: 52,
  },
  toolChip: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: colors.surfaceContainerLow,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.9),
    alignItems: "center",
    justifyContent: "center",
  },
  toolChipActive: {
    backgroundColor: colors.primary,
    borderColor: colors.primaryDim,
  },
  toolStripDivider: {
    width: 1,
    height: 28,
    marginHorizontal: spacing.xs,
    backgroundColor: withAlpha(colors.outlineVariant, 0.95),
  },
  filterIconBtn: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: colors.surfaceContainerLowest,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.9),
    alignItems: "center",
    justifyContent: "center",
  },
  filterIconBtnActive: {
    backgroundColor: colors.primary,
    borderColor: colors.primaryDim,
  },
  nextButton: {
    marginTop: spacing.sm,
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
