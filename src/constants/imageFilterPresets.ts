/**
 * Per-image filter presets. Preview uses Image + overlays; export applies crop/flip then
 * optional tint via manipulate pipeline where possible.
 */
export type ImageFilterId =
  | "none"
  | "mono"
  | "warm"
  | "cool"
  | "vivid"
  | "fade"
  | "noir"
  | "chrome"
  | "dramatic"
  | "soft"
  | "golden"
  | "moody"
  | "sunset"
  | "mint"
  | "rose"
  | "vintage"
  | "sepia"
  | "crisp"
  | "ocean"
  | "ember";

export type FilterPreviewLayers = {
  /** Applied to <Image> via style */
  imageStyle?: Record<string, unknown>;
  /** Absolute overlay on top of image (multiply / darken look) */
  overlayColor?: string;
  overlayOpacity?: number;
};

/** MaterialIcons glyph name for compact tool row (see @expo/vector-icons/MaterialIcons). */
export type ImageFilterPreset = {
  id: ImageFilterId;
  label: string;
  materialIcon: string;
  preview: FilterPreviewLayers;
};

export const IMAGE_FILTER_PRESETS: ImageFilterPreset[] = [
  { id: "none", label: "Original", materialIcon: "image", preview: {} },
  {
    id: "vivid",
    label: "Vivid",
    materialIcon: "auto-awesome",
    preview: { overlayColor: "#ff6b4a", overlayOpacity: 0.12 },
  },
  {
    id: "warm",
    label: "Warm",
    materialIcon: "wb-sunny",
    preview: { overlayColor: "#ffb347", overlayOpacity: 0.22 },
  },
  {
    id: "cool",
    label: "Cool",
    materialIcon: "ac-unit",
    preview: { overlayColor: "#6ec8ff", overlayOpacity: 0.18 },
  },
  {
    id: "golden",
    label: "Golden",
    materialIcon: "flare",
    preview: { overlayColor: "#f4d03f", overlayOpacity: 0.2 },
  },
  {
    id: "fade",
    label: "Fade",
    materialIcon: "brightness-low",
    preview: {
      imageStyle: { opacity: 0.88 },
      overlayColor: "#ffffff",
      overlayOpacity: 0.15,
    },
  },
  {
    id: "soft",
    label: "Soft",
    materialIcon: "blur-on",
    preview: {
      imageStyle: { opacity: 0.92 },
      overlayColor: "#f5e6ff",
      overlayOpacity: 0.12,
    },
  },
  {
    id: "mono",
    label: "Mono",
    materialIcon: "invert-colors",
    preview: {
      imageStyle: { opacity: 0.95 },
      overlayColor: "#888888",
      overlayOpacity: 0.35,
    },
  },
  {
    id: "noir",
    label: "Noir",
    materialIcon: "nights-stay",
    preview: {
      imageStyle: { opacity: 0.9 },
      overlayColor: "#1a1a2e",
      overlayOpacity: 0.45,
    },
  },
  {
    id: "chrome",
    label: "Chrome",
    materialIcon: "hdr-strong",
    preview: { overlayColor: "#c0d8e8", overlayOpacity: 0.2 },
  },
  {
    id: "dramatic",
    label: "Drama",
    materialIcon: "contrast",
    preview: {
      imageStyle: { opacity: 0.88 },
      overlayColor: "#2c1810",
      overlayOpacity: 0.35,
    },
  },
  {
    id: "moody",
    label: "Moody",
    materialIcon: "dark-mode",
    preview: {
      imageStyle: { opacity: 0.9 },
      overlayColor: "#4a2c5a",
      overlayOpacity: 0.28,
    },
  },
  {
    id: "sunset",
    label: "Sunset",
    materialIcon: "wb-twilight",
    preview: { overlayColor: "#ff7e5f", overlayOpacity: 0.22 },
  },
  {
    id: "mint",
    label: "Mint",
    materialIcon: "spa",
    preview: { overlayColor: "#7fd8be", overlayOpacity: 0.18 },
  },
  {
    id: "rose",
    label: "Rose",
    materialIcon: "palette",
    preview: { overlayColor: "#f8b4c4", overlayOpacity: 0.2 },
  },
  {
    id: "vintage",
    label: "Vintage",
    materialIcon: "photo-filter",
    preview: {
      imageStyle: { opacity: 0.9 },
      overlayColor: "#c4a574",
      overlayOpacity: 0.25,
    },
  },
  {
    id: "sepia",
    label: "Sepia",
    materialIcon: "filter-vintage",
    preview: {
      imageStyle: { opacity: 0.92 },
      overlayColor: "#8b6914",
      overlayOpacity: 0.32,
    },
  },
  {
    id: "crisp",
    label: "Crisp",
    materialIcon: "tonality",
    preview: { overlayColor: "#e8f4fc", overlayOpacity: 0.1 },
  },
  {
    id: "ocean",
    label: "Ocean",
    materialIcon: "water",
    preview: { overlayColor: "#006994", overlayOpacity: 0.2 },
  },
  {
    id: "ember",
    label: "Ember",
    materialIcon: "whatshot",
    preview: { overlayColor: "#c0392b", overlayOpacity: 0.18 },
  },
];
