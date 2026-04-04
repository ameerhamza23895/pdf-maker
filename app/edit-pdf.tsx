// @ts-nocheck
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import MaterialIcons from '@expo/vector-icons/MaterialIcons';
import { electricCuratorTheme, withAlpha } from '@/src/theme/electric-curator';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import {
  ActivityIndicator,
  Alert,
  Modal,
  PanResponder,
  Platform,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  TouchableOpacity,
  View,
} from 'react-native';
import { StatusBar } from 'expo-status-bar';
import { LinearGradient } from 'expo-linear-gradient';
import * as DocumentPicker from 'expo-document-picker';
import { File, Paths } from 'expo-file-system';
import * as LegacyFileSystem from 'expo-file-system/legacy';
import * as Print from 'expo-print';
import * as Sharing from 'expo-sharing';
import {
  BlendMode,
  LineCapStyle,
  PDFDocument,
  StandardFonts,
  rgb,
} from 'pdf-lib';
import { WebView } from 'react-native-webview';
import { getPdfViewerHtml } from '@/src/utils/pdfViewerHtml';
import { getPdfConverterHtml } from '@/src/utils/pdfConverterHtml';
import { getOfficePreviewHtml } from '@/src/utils/officePreviewHtml';

const { colors, spacing, radius } = electricCuratorTheme;
const ui = {
  shell: colors.surface,
  shellElevated: colors.surfaceContainerLowest,
  shellSoft: colors.surfaceContainerLow,
  border: colors.outlineVariant,
  text: colors.onSurface,
  textMuted: withAlpha(colors.onSurface, 0.72),
  textSoft: withAlpha(colors.onSurface, 0.55),
  accent: colors.primary,
  accentStrong: colors.primaryDim,
  accentSoft: colors.primaryContainer,
  accentText: colors.onPrimary,
  chip: colors.secondaryContainer,
  chipText: colors.onSecondaryContainer,
  cardShadow: withAlpha(colors.onSurface, 0.08),
  danger: '#c94b6d',
  warning: '#d6852d',
  success: '#0c7564',
};

const TOOLS = [
  {
    id: 'view',
    label: 'Select',
    icon: 'touch-app',
    hint: 'Select and move annotations',
  },
  {
    id: 'highlight',
    label: 'Area highlight',
    icon: 'highlight-alt',
    hint: 'Drag a rectangle over text',
  },
  {
    id: 'marker',
    label: 'Highlighter',
    icon: 'format-color-fill',
    hint: 'Freehand translucent highlighter strokes',
  },
  {
    id: 'draw',
    label: 'Pen',
    icon: 'brush',
    hint: 'Solid pen strokes — adjust brush size below',
  },
  {
    id: 'text',
    label: 'Note',
    icon: 'sticky-note-2',
    hint: 'Add sticky note',
  },
];

const COLORS = [
  { name: 'Red', value: '#FF0000' },
  { name: 'Blue', value: '#2196F3' },
  { name: 'Green', value: '#4CAF50' },
  { name: 'Orange', value: '#FF9800' },
  { name: 'Purple', value: '#9C27B0' },
];

/** Mini gradient strips — tap uses the left (primary) stop as the solid color. */
const GRADIENT_COLOR_PRESETS = [
  { id: 'dawn', colors: ['#FF9A9E', '#FECFEF'], label: 'Dawn' },
  { id: 'sunset', colors: ['#FF512F', '#F09819'], label: 'Sunset' },
  { id: 'ocean', colors: ['#2193B0', '#6DD5ED'], label: 'Ocean' },
  { id: 'lavender', colors: ['#834D9B', '#D04ED6'], label: 'Lilac' },
  { id: 'mint', colors: ['#00B09B', '#96C93D'], label: 'Mint' },
  { id: 'royal', colors: ['#141E30', '#243B55'], label: 'Royal' },
  { id: 'peach', colors: ['#FFECD2', '#FCB69F'], label: 'Peach' },
  { id: 'ember', colors: ['#F12711', '#F5AF19'], label: 'Ember' },
];

const DEFAULT_COLOR = COLORS[0].value;
const HIGHLIGHT_ALPHA_HEX = '66';
const ANNOTATION_STORAGE_DIR = `${Paths.document.uri}annotation-state/`;
const WORKING_PDF_SUFFIX = '.working.pdf';
const OFFICE_PREVIEW_TYPES = ['doc', 'docx', 'pptx', 'xls', 'xlsx'];
const PDF_CONVERSION_FORMATS = [
  {
    id: 'docx',
    label: 'Word (.docx)',
    extension: 'docx',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  },
  {
    id: 'pptx',
    label: 'PowerPoint (.pptx)',
    extension: 'pptx',
    mimeType:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  },
  {
    id: 'txt',
    label: 'Text (.txt)',
    extension: 'txt',
    mimeType: 'text/plain',
  },
  {
    id: 'md',
    label: 'Markdown (.md)',
    extension: 'md',
    mimeType: 'text/markdown',
  },
  {
    id: 'html',
    label: 'HTML (.html)',
    extension: 'html',
    mimeType: 'text/html',
  },
];
const OFFICE_FONT_FAMILY_OPTIONS = [
  { label: 'Arial', value: 'Arial' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Times', value: 'Times New Roman' },
  { label: 'Verdana', value: 'Verdana' },
  { label: 'Courier', value: 'Courier New' },
  { label: 'Tahoma', value: 'Tahoma' },
];
const OFFICE_HIGHLIGHT_COLOR = '#FFF59D';
const OFFICE_FONT_SIZE_OPTIONS = [
  { label: 'A-', action: 'adjustFontSize', value: -2 },
  { label: 'A', action: 'resetFontSize', value: 16 },
  { label: 'A+', action: 'adjustFontSize', value: 2 },
];

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

/** @returns {{ r: number, g: number, b: number } | null} */
function parseHexStringToRgb(input) {
  if (!input || typeof input !== 'string') {
    return null;
  }
  let s = input.trim();
  if (s.startsWith('#')) {
    s = s.slice(1);
  }
  if (s.length === 3) {
    s = s
      .split('')
      .map((c) => c + c)
      .join('');
  }
  if (!/^[0-9a-fA-F]{6}$/.test(s)) {
    return null;
  }
  return {
    r: parseInt(s.slice(0, 2), 16),
    g: parseInt(s.slice(2, 4), 16),
    b: parseInt(s.slice(4, 6), 16),
  };
}

function rgbToHex(r, g, b) {
  const c = (n) => clamp(Math.round(Number(n)), 0, 255);
  return (
    '#' +
    [c(r), c(g), c(b)]
      .map((x) => x.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  );
}

function rgbToHsv(r, g, b) {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const d = max - min;
  let h = 0;
  if (d !== 0) {
    if (max === rn) {
      h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6;
    } else if (max === gn) {
      h = ((bn - rn) / d + 2) / 6;
    } else {
      h = ((rn - gn) / d + 4) / 6;
    }
  }
  h *= 360;
  const s = max === 0 ? 0 : d / max;
  const v = max;
  return { h, s, v };
}

function hsvToRgb(h, s, v) {
  const hh = ((h % 360) + 360) % 360;
  const c = v * s;
  const x = c * (1 - Math.abs(((hh / 60) % 2) - 1));
  const m = v - c;
  let rp = 0;
  let gp = 0;
  let bp = 0;
  if (hh < 60) {
    rp = c;
    gp = x;
  } else if (hh < 120) {
    rp = x;
    gp = c;
  } else if (hh < 180) {
    gp = c;
    bp = x;
  } else if (hh < 240) {
    gp = x;
    bp = c;
  } else if (hh < 300) {
    rp = x;
    bp = c;
  } else {
    rp = c;
    bp = x;
  }
  return {
    r: Math.round((rp + m) * 255),
    g: Math.round((gp + m) * 255),
    b: Math.round((bp + m) * 255),
  };
}

function colorMatchesPreset(hex, presetList = COLORS) {
  const p = parseHexStringToRgb(hex);
  if (!p) {
    return false;
  }
  const h = rgbToHex(p.r, p.g, p.b);
  return presetList.some((c) => c.value.toUpperCase() === h.toUpperCase());
}

function createEmptyPageAnnotations() {
  return {
    highlights: [],
    drawings: [],
    notes: [],
  };
}

function normalizeAnnotationState(rawAnnotations) {
  if (!rawAnnotations || typeof rawAnnotations !== 'object') {
    return {};
  }

  return Object.entries(rawAnnotations).reduce((pages, [pageKey, pageAnnotations]) => {
    if (!pageAnnotations || typeof pageAnnotations !== 'object') {
      return pages;
    }

    pages[pageKey] = createEmptyPageAnnotations();
    pages[pageKey].highlights = Array.isArray(pageAnnotations.highlights)
      ? pageAnnotations.highlights
      : [];
    pages[pageKey].drawings = Array.isArray(pageAnnotations.drawings)
      ? pageAnnotations.drawings
      : [];
    pages[pageKey].notes = Array.isArray(pageAnnotations.notes)
      ? pageAnnotations.notes
      : [];

    return pages;
  }, {});
}

function hasAnnotations(annotationState = {}) {
  return Object.values(annotationState).some((pageAnnotations) => {
    if (!pageAnnotations || typeof pageAnnotations !== 'object') {
      return false;
    }

    return (
      (pageAnnotations.highlights || []).length > 0 ||
      (pageAnnotations.drawings || []).length > 0 ||
      (pageAnnotations.notes || []).length > 0
    );
  });
}

function getAnnotationStorageUri(documentKey) {
  return `${ANNOTATION_STORAGE_DIR}${documentKey}.json`;
}

function getWorkingPdfStorageUri(documentKey) {
  return `${ANNOTATION_STORAGE_DIR}${documentKey}${WORKING_PDF_SUFFIX}`;
}

async function deleteFileIfExists(uri) {
  const info = await LegacyFileSystem.getInfoAsync(uri);
  if (info.exists) {
    await LegacyFileSystem.deleteAsync(uri);
  }
}

function insertPageIntoAnnotationState(rawAnnotations, afterPageNumber) {
  const normalizedAnnotations = normalizeAnnotationState(rawAnnotations);
  const nextAnnotations = {
    [afterPageNumber + 1]: createEmptyPageAnnotations(),
  };

  Object.entries(normalizedAnnotations)
    .sort((a, b) => Number(a[0]) - Number(b[0]))
    .forEach(([pageKey, pageAnnotations]) => {
      const pageNumber = Number(pageKey);
      const nextPageNumber = pageNumber <= afterPageNumber ? pageNumber : pageNumber + 1;
      nextAnnotations[nextPageNumber] = pageAnnotations;
    });

  return nextAnnotations;
}

function removePageFromAnnotationState(rawAnnotations, removedPageNumber) {
  const normalizedAnnotations = normalizeAnnotationState(rawAnnotations);
  const nextAnnotations = {};

  Object.entries(normalizedAnnotations)
    .sort((a, b) => Number(a[0]) - Number(b[0]))
    .forEach(([pageKey, pageAnnotations]) => {
      const pageNumber = Number(pageKey);
      if (pageNumber === removedPageNumber) {
        return;
      }

      const nextPageNumber = pageNumber < removedPageNumber ? pageNumber : pageNumber - 1;
      nextAnnotations[nextPageNumber] = pageAnnotations;
    });

  return nextAnnotations;
}

function sanitizeFileSegment(value = 'document') {
  return (
    value
      .replace(/\.pdf$/i, '')
      .replace(/[^a-z0-9-_]+/gi, '_')
      .replace(/^_+|_+$/g, '') || 'document'
  );
}

function buildTimestampToken() {
  return new Date()
    .toISOString()
    .replace(/[^0-9]/g, '')
    .slice(0, 14);
}

function buildAnnotatedFileName(documentName) {
  return `${sanitizeFileSegment(documentName)}_annotated_${buildTimestampToken()}.pdf`;
}

function buildConvertedFileName(documentName, extension) {
  return `${sanitizeFileSegment(documentName)}_converted_${buildTimestampToken()}.${extension}`;
}

function parseHexColor(hexColor, fallbackOpacity = 1) {
  if (typeof hexColor !== 'string' || !hexColor.startsWith('#')) {
    return { color: rgb(1, 0, 0), opacity: fallbackOpacity };
  }

  let normalized = hexColor.slice(1);
  if (normalized.length === 3 || normalized.length === 4) {
    normalized = normalized
      .split('')
      .map((character) => character + character)
      .join('');
  }

  let opacity = fallbackOpacity;
  if (normalized.length === 8) {
    opacity = parseInt(normalized.slice(6, 8), 16) / 255;
    normalized = normalized.slice(0, 6);
  }

  if (normalized.length !== 6) {
    return { color: rgb(1, 0, 0), opacity: fallbackOpacity };
  }

  const red = parseInt(normalized.slice(0, 2), 16);
  const green = parseInt(normalized.slice(2, 4), 16);
  const blue = parseInt(normalized.slice(4, 6), 16);

  if ([red, green, blue].some((channel) => Number.isNaN(channel))) {
    return { color: rgb(1, 0, 0), opacity: fallbackOpacity };
  }

  return {
    color: rgb(red / 255, green / 255, blue / 255),
    opacity,
  };
}

function normalizePdfPoint(point, pageWidth, pageHeight) {
  return {
    x: clamp(Number(point?.x) || 0, 0, 1) * pageWidth,
    y: pageHeight - clamp(Number(point?.y) || 0, 0, 1) * pageHeight,
  };
}

function wrapNoteText(text, font, fontSize, maxWidth) {
  const paragraphs = String(text || '').split(/\r?\n/);
  const lines = [];

  const pushWrappedWord = (word) => {
    let fragment = '';
    for (const character of word) {
      const candidate = fragment + character;
      if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth) {
        fragment = candidate;
      } else {
        if (fragment) {
          lines.push(fragment);
        }
        fragment = character;
      }
    }

    if (fragment) {
      lines.push(fragment);
    }
  };

  paragraphs.forEach((paragraph) => {
    const words = paragraph.trim().split(/\s+/).filter(Boolean);

    if (!words.length) {
      lines.push('');
      return;
    }

    let currentLine = '';

    for (const word of words) {
      if (!currentLine) {
        if (font.widthOfTextAtSize(word, fontSize) > maxWidth) {
          pushWrappedWord(word);
          continue;
        }

        currentLine = word;
        continue;
      }

      const nextLine = `${currentLine} ${word}`;

      if (font.widthOfTextAtSize(nextLine, fontSize) <= maxWidth) {
        currentLine = nextLine;
        continue;
      }

      if (font.widthOfTextAtSize(word, fontSize) > maxWidth) {
        lines.push(currentLine);
        currentLine = '';
        pushWrappedWord(word);
      } else {
        lines.push(currentLine);
        currentLine = word;
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }
  });

  return lines.length ? lines : [''];
}

async function bakePdfWithAnnotations(base64Contents, annotationState) {
  const pdfDocument = await PDFDocument.load(base64Contents);
  const noteFont = await pdfDocument.embedFont(StandardFonts.Helvetica);
  const pages = pdfDocument.getPages();

  Object.entries(annotationState).forEach(([pageKey, pageAnnotations]) => {
    const page = pages[Number(pageKey) - 1];
    if (!page) {
      return;
    }

    const { width, height } = page.getSize();

    (pageAnnotations.highlights || []).forEach((highlight) => {
      const highlightWidth = clamp(Number(highlight.width) || 0, 0, 1) * width;
      const highlightHeight = clamp(Number(highlight.height) || 0, 0, 1) * height;

      if (highlightWidth < 1 || highlightHeight < 1) {
        return;
      }

      const { color, opacity } = parseHexColor(highlight.color, 0.4);
      const x = clamp(Number(highlight.x) || 0, 0, 1) * width;
      const y =
        height -
        (clamp(Number(highlight.y) || 0, 0, 1) + clamp(Number(highlight.height) || 0, 0, 1)) *
          height;

      page.drawRectangle({
        x,
        y,
        width: highlightWidth,
        height: highlightHeight,
        color,
        opacity,
        blendMode: BlendMode.Multiply,
      });
    });

    (pageAnnotations.drawings || []).forEach((drawing) => {
      const pdfPoints = (drawing.points || []).map((point) =>
        normalizePdfPoint(point, width, height)
      );

      if (!pdfPoints.length) {
        return;
      }

      const isMarker = drawing.kind === 'marker';
      const { color, opacity } = parseHexColor(drawing.color, isMarker ? 0.45 : 1);
      const thickness = Math.max(
        clamp(Number(drawing.width) || 0.006, 0.001, 0.09) * width,
        1.2
      );

      if (pdfPoints.length === 1) {
        page.drawCircle({
          x: pdfPoints[0].x,
          y: pdfPoints[0].y,
          size: Math.max(thickness / 2, 0.8),
          color,
          opacity,
          blendMode: isMarker ? BlendMode.Multiply : undefined,
        });
        return;
      }

      for (let index = 0; index < pdfPoints.length - 1; index += 1) {
        page.drawLine({
          start: pdfPoints[index],
          end: pdfPoints[index + 1],
          thickness,
          color,
          opacity,
          lineCap: LineCapStyle.Round,
          blendMode: isMarker ? BlendMode.Multiply : undefined,
        });
      }
    });

    (pageAnnotations.notes || []).forEach((note) => {
      const fontSize = 12;
      const lineHeight = fontSize * 1.25;
      const maxTextWidth = Math.min(width * 0.35, 180);
      const lines = wrapNoteText(note.text, noteFont, fontSize, maxTextWidth);
      const textWidth = lines.reduce((currentWidth, line) => {
        const measuredWidth = noteFont.widthOfTextAtSize(line || ' ', fontSize);
        return Math.max(currentWidth, measuredWidth);
      }, 40);

      const noteWidth = Math.min(textWidth + 14, width * 0.42);
      const noteHeight = Math.max(lines.length * lineHeight + 14, 28);
      const x = clamp(
        (Number(note.x) || 0) * width,
        8,
        Math.max(width - noteWidth - 8, 8)
      );
      const y = clamp(
        height - (clamp(Number(note.y) || 0, 0, 1) * height) - noteHeight,
        8,
        Math.max(height - noteHeight - 8, 8)
      );

      page.drawRectangle({
        x,
        y,
        width: noteWidth,
        height: noteHeight,
        color: rgb(1, 0.976, 0.769),
        borderWidth: 1,
        borderColor: rgb(0.976, 0.659, 0.145),
        opacity: 0.98,
      });

      lines.forEach((line, index) => {
        page.drawText(line, {
          x: x + 7,
          y: y + noteHeight - 9 - fontSize - index * lineHeight,
          size: fontSize,
          font: noteFont,
          color: rgb(0.2, 0.2, 0.2),
        });
      });
    });
  });

  return pdfDocument.saveAsBase64();
}

function getSelectedAnnotationDescription(selectedAnnotation) {
  if (!selectedAnnotation) {
    return 'Tap any highlight, drawing, or note to select it.';
  }

  const label =
    selectedAnnotation.type === 'highlight'
      ? 'highlight'
      : selectedAnnotation.type === 'drawing'
        ? 'drawing'
        : 'note';

  return `Selected ${label} on page ${selectedAnnotation.page}.`;
}

async function saveBase64ToAndroidDeviceFolder(base64Contents, fileName, mimeType) {
  if (
    Platform.OS !== 'android' ||
    !LegacyFileSystem.StorageAccessFramework
  ) {
    return { savedToDevice: false, reason: 'unsupported_platform' };
  }

  try {
    const downloadsRootUri =
      LegacyFileSystem.StorageAccessFramework.getUriForDirectoryInRoot('Download');
    const permission =
      await LegacyFileSystem.StorageAccessFramework.requestDirectoryPermissionsAsync(
        downloadsRootUri
      );

    if (!permission.granted) {
      return { savedToDevice: false, reason: 'permission_denied' };
    }

    const targetUri = await LegacyFileSystem.StorageAccessFramework.createFileAsync(
      permission.directoryUri,
      fileName,
      mimeType
    );

    await LegacyFileSystem.StorageAccessFramework.writeAsStringAsync(
      targetUri,
      base64Contents,
      {
        encoding: LegacyFileSystem.EncodingType.Base64,
      }
    );

    return {
      savedToDevice: true,
      directoryUri: permission.directoryUri,
      targetUri,
    };
  } catch (error) {
    return {
      savedToDevice: false,
      reason: 'write_failed',
      error,
    };
  }
}

async function savePdfToAndroidDeviceFolder(base64Contents, fileName) {
  return saveBase64ToAndroidDeviceFolder(
    base64Contents,
    fileName,
    'application/pdf'
  );
}

function BrushWidthTrack({
  min,
  max,
  value,
  onChange,
  accessibilityLabel,
  trackStyle,
  fillStyle,
  integerStep = false,
}) {
  const widthRef = useRef(1);
  const applyLocationX = useCallback(
    (x) => {
      const w = widthRef.current;
      const ratio = Math.max(0, Math.min(1, x / w));
      const raw = min + ratio * (max - min);
      const stepped = integerStep
        ? Math.round(raw)
        : Math.round(raw * 2) / 2;
      onChange(Math.max(min, Math.min(max, stepped)));
    },
    [min, max, onChange, integerStep]
  );

  const panResponder = useMemo(
    () =>
      PanResponder.create({
        onStartShouldSetPanResponder: () => true,
        onMoveShouldSetPanResponder: () => true,
        onPanResponderTerminationRequest: () => false,
        onPanResponderGrant: (e) => {
          applyLocationX(e.nativeEvent.locationX);
        },
        onPanResponderMove: (e) => {
          applyLocationX(e.nativeEvent.locationX);
        },
      }),
    [applyLocationX]
  );

  return (
    <View
      style={trackStyle}
      onLayout={(e) => {
        widthRef.current = e.nativeEvent.layout.width || 1;
      }}
      accessibilityLabel={accessibilityLabel}
      accessibilityRole="adjustable"
      {...panResponder.panHandlers}
    >
      <View
        style={[
          fillStyle,
          {
            width: `${((value - min) / (max - min)) * 100}%`,
          },
        ]}
      />
    </View>
  );
}

const HUE_STRIP_GRADIENT = [
  '#FF0000',
  '#FFFF00',
  '#00FF00',
  '#00FFFF',
  '#0000FF',
  '#FF00FF',
  '#FF0000',
];

function CustomColorModal({ visible, seedHex, title, onApply, onClose }) {
  const [hsv, setHsv] = useState({ h: 0, s: 1, v: 1 });
  const [hexInput, setHexInput] = useState('#FF0000');
  const [showRgbAdvanced, setShowRgbAdvanced] = useState(false);
  const [svLayout, setSvLayout] = useState({ w: 280, h: 200 });
  const [hueLayout, setHueLayout] = useState({ w: 280 });
  const svBoxRef = useRef({ w: 280, h: 200 });
  const hueBarRef = useRef({ w: 280 });

  useEffect(() => {
    if (!visible) {
      return;
    }
    const p = parseHexStringToRgb(seedHex) || { r: 255, g: 0, b: 0 };
    const nextHsv = rgbToHsv(p.r, p.g, p.b);
    setHsv(nextHsv);
    setHexInput(rgbToHex(p.r, p.g, p.b));
  }, [visible, seedHex]);

  const rgb = hsvToRgb(hsv.h, hsv.s, hsv.v);
  const previewHex = rgbToHex(rgb.r, rgb.g, rgb.b);
  const pureHueRgb = hsvToRgb(hsv.h, 1, 1);
  const pureHueHex = rgbToHex(pureHueRgb.r, pureHueRgb.g, pureHueRgb.b);

  const applySvFromLocal = useCallback((x, y) => {
    const { w, h } = svBoxRef.current;
    const sx = clamp(x, 0, w);
    const sy = clamp(y, 0, h);
    const s = sx / w;
    const v = 1 - sy / h;
    setHsv((prev) => {
      const next = { ...prev, s, v };
      const r = hsvToRgb(next.h, next.s, next.v);
      setHexInput(rgbToHex(r.r, r.g, r.b));
      return next;
    });
  }, []);

  const applyHueFromLocal = useCallback((x) => {
    const w = hueBarRef.current.w;
    const hx = clamp(x, 0, w);
    const hDeg = (hx / w) * 360;
    setHsv((prev) => {
      const next = { ...prev, h: hDeg };
      const r = hsvToRgb(next.h, next.s, next.v);
      setHexInput(rgbToHex(r.r, r.g, r.b));
      return next;
    });
  }, []);

  const svPan = useMemo(
    () =>
      PanResponder.create({
        onStartShouldSetPanResponder: () => true,
        onMoveShouldSetPanResponder: () => true,
        onPanResponderTerminationRequest: () => false,
        onPanResponderGrant: (e) => {
          applySvFromLocal(e.nativeEvent.locationX, e.nativeEvent.locationY);
        },
        onPanResponderMove: (e) => {
          applySvFromLocal(e.nativeEvent.locationX, e.nativeEvent.locationY);
        },
      }),
    [applySvFromLocal]
  );

  const huePan = useMemo(
    () =>
      PanResponder.create({
        onStartShouldSetPanResponder: () => true,
        onMoveShouldSetPanResponder: () => true,
        onPanResponderTerminationRequest: () => false,
        onPanResponderGrant: (e) => {
          applyHueFromLocal(e.nativeEvent.locationX);
        },
        onPanResponderMove: (e) => {
          applyHueFromLocal(e.nativeEvent.locationX);
        },
      }),
    [applyHueFromLocal]
  );

  const onHexInputChange = useCallback((text) => {
    setHexInput(text);
    const p = parseHexStringToRgb(text);
    if (p) {
      setHsv(rgbToHsv(p.r, p.g, p.b));
    }
  }, []);

  const setChannel = useCallback((channel, v) => {
    setHsv((prev) => {
      const r0 = hsvToRgb(prev.h, prev.s, prev.v);
      const nextRgb = { ...r0, [channel]: v };
      const nextHsv = rgbToHsv(nextRgb.r, nextRgb.g, nextRgb.b);
      setHexInput(rgbToHex(nextRgb.r, nextRgb.g, nextRgb.b));
      return nextHsv;
    });
  }, []);

  const applyGradientPreset = useCallback((hex) => {
    const p = parseHexStringToRgb(hex);
    if (!p) {
      return;
    }
    const nextHsv = rgbToHsv(p.r, p.g, p.b);
    setHsv(nextHsv);
    setHexInput(rgbToHex(p.r, p.g, p.b));
  }, []);

  const handleApply = useCallback(() => {
    const p = parseHexStringToRgb(hexInput);
    const final = p ? rgbToHex(p.r, p.g, p.b) : previewHex;
    onApply(final);
  }, [hexInput, previewHex, onApply]);

  const svW = svLayout.w || 1;
  const svH = svLayout.h || 1;
  const hueW = hueLayout.w || 1;
  const svThumbLeft = hsv.s * svW - 11;
  const svThumbTop = (1 - hsv.v) * svH - 11;
  const hueThumbLeft = (hsv.h / 360) * hueW;
  const hueThumbW = 5;

  return (
    <Modal
      visible={visible}
      transparent={true}
      animationType="fade"
      onRequestClose={onClose}
    >
      <View style={styles.modalOverlay}>
        <View style={[styles.modalContent, styles.customColorModalContent]}>
          <View style={styles.customColorHeader}>
            <Text style={styles.customColorHeaderTitle} numberOfLines={1}>
              {title}
            </Text>
            <TouchableOpacity
              onPress={onClose}
              hitSlop={{ top: 12, bottom: 12, left: 12, right: 12 }}
              accessibilityLabel="Close"
            >
              <MaterialIcons name="close" size={22} color={colors.onSurface} />
            </TouchableOpacity>
          </View>

          <ScrollView
            style={styles.customColorScroll}
            keyboardShouldPersistTaps="handled"
            showsVerticalScrollIndicator={false}
          >
            <Text style={styles.customColorHint}>
              Drag the square and hue bar (like Google&apos;s picker), choose a gradient
              swatch, or type hex below.
            </Text>

            <View style={styles.customColorPreviewRow}>
              <View
                style={[styles.customColorPreviewCircle, { backgroundColor: previewHex }]}
              />
              <View style={styles.customColorPreviewMeta}>
                <Text style={styles.customColorPreviewLabel}>Selected</Text>
                <Text style={styles.customColorPreviewHex}>{previewHex}</Text>
              </View>
            </View>

            <Text style={styles.customColorSectionLabel}>Spectrum</Text>
            <View style={styles.spectrumPadOuter}>
              <View
                style={styles.spectrumPad}
                onLayout={(e) => {
                  const { width, height } = e.nativeEvent.layout;
                  const w = width || 1;
                  const h = height || 1;
                  svBoxRef.current = { w, h };
                  setSvLayout({ w, h });
                }}
                {...svPan.panHandlers}
              >
                <LinearGradient
                  style={StyleSheet.absoluteFill}
                  colors={['#FFFFFF', pureHueHex]}
                  start={{ x: 0, y: 0.5 }}
                  end={{ x: 1, y: 0.5 }}
                />
                <LinearGradient
                  style={StyleSheet.absoluteFill}
                  colors={['rgba(255,255,255,0)', '#000000']}
                  start={{ x: 0.5, y: 0 }}
                  end={{ x: 0.5, y: 1 }}
                />
                <View
                  pointerEvents="none"
                  style={[
                    styles.svThumb,
                    {
                      left: clamp(svThumbLeft, -2, svW - 20),
                      top: clamp(svThumbTop, -2, svH - 20),
                    },
                  ]}
                />
              </View>
            </View>

            <Text style={styles.customColorSectionLabel}>Hue</Text>
            <View
              style={styles.hueStripOuter}
              onLayout={(e) => {
                const w = e.nativeEvent.layout.width || 1;
                hueBarRef.current = { w };
                setHueLayout({ w });
              }}
            >
              <LinearGradient
                style={styles.hueStripGradient}
                colors={HUE_STRIP_GRADIENT}
                start={{ x: 0, y: 0.5 }}
                end={{ x: 1, y: 0.5 }}
              />
              <View style={styles.hueStripOverlay} {...huePan.panHandlers} />
              <View
                pointerEvents="none"
                style={[
                  styles.hueThumb,
                  {
                    left: clamp(
                      hueThumbLeft - hueThumbW / 2,
                      0,
                      Math.max(0, hueW - hueThumbW)
                    ),
                  },
                ]}
              />
            </View>

            <Text style={styles.customColorSectionLabel}>Gradient looks</Text>
            <Text style={styles.customColorGradientHint}>
              Tap a strip — uses the left color (solid for pen, text, and highlights).
            </Text>
            <ScrollView
              horizontal
              showsHorizontalScrollIndicator={false}
              contentContainerStyle={styles.gradientPresetsScroll}
            >
              {GRADIENT_COLOR_PRESETS.map((preset) => (
                <TouchableOpacity
                  key={preset.id}
                  style={styles.gradientPresetChip}
                  onPress={() => applyGradientPreset(preset.colors[0])}
                  accessibilityLabel={`Gradient ${preset.label}, primary ${preset.colors[0]}`}
                >
                  <LinearGradient
                    colors={preset.colors}
                    start={{ x: 0, y: 0.5 }}
                    end={{ x: 1, y: 0.5 }}
                    style={styles.gradientPresetFill}
                  />
                  <Text style={styles.gradientPresetLabel}>{preset.label}</Text>
                </TouchableOpacity>
              ))}
            </ScrollView>

            <Text style={styles.customColorSectionLabel}>Hex</Text>
            <TextInput
              style={styles.customColorHexInput}
              value={hexInput}
              onChangeText={onHexInputChange}
              autoCapitalize="characters"
              autoCorrect={false}
              placeholder="#RRGGBB"
              placeholderTextColor={withAlpha(colors.onSurface, 0.45)}
            />

            <TouchableOpacity
              style={styles.customColorAdvancedToggle}
              onPress={() => setShowRgbAdvanced((v) => !v)}
              accessibilityLabel={showRgbAdvanced ? 'Hide RGB sliders' : 'Show RGB sliders'}
            >
              <MaterialIcons
                name="tune"
                size={20}
                color={colors.primaryDim}
              />
              <Text style={styles.customColorAdvancedToggleText}>
                {showRgbAdvanced ? 'Hide RGB sliders' : 'RGB sliders (advanced)'}
              </Text>
              <MaterialIcons
                name={showRgbAdvanced ? 'expand-less' : 'expand-more'}
                size={22}
                color={withAlpha(colors.onSurface, 0.6)}
              />
            </TouchableOpacity>

            {showRgbAdvanced && (
              <>
                {[
                  { key: 'r', label: 'Red' },
                  { key: 'g', label: 'Green' },
                  { key: 'b', label: 'Blue' },
                ].map(({ key, label }) => (
                  <View key={key} style={styles.customColorRgbRow}>
                    <Text style={styles.customColorChannelLabel}>{label}</Text>
                    <BrushWidthTrack
                      min={0}
                      max={255}
                      value={rgb[key]}
                      integerStep={true}
                      onChange={(v) => setChannel(key, v)}
                      accessibilityLabel={`${label}`}
                      trackStyle={styles.customColorTrack}
                      fillStyle={styles.brushTrackFill}
                    />
                    <Text style={styles.customColorChannelValue}>{rgb[key]}</Text>
                  </View>
                ))}
              </>
            )}
          </ScrollView>

          <View style={styles.modalButtons}>
            <TouchableOpacity style={styles.modalBtnCancel} onPress={onClose}>
              <Text style={styles.modalBtnCancelText}>Cancel</Text>
            </TouchableOpacity>
            <TouchableOpacity style={styles.modalBtnSubmit} onPress={handleApply}>
              <Text style={styles.modalBtnSubmitText}>Apply</Text>
            </TouchableOpacity>
          </View>
        </View>
      </View>
    </Modal>
  );
}

export default function EditPdfPage() {
  const [documentUri, setDocumentUri] = useState(null);
  const [documentName, setDocumentName] = useState('');
  const [documentType, setDocumentType] = useState('');
  const [documentKey, setDocumentKey] = useState(null);
  const [base64Data, setBase64Data] = useState(null);
  const [viewerHtml, setViewerHtml] = useState(null);
  const [annotationState, setAnnotationState] = useState({});
  const [selectedAnnotation, setSelectedAnnotation] = useState(null);
  const [loading, setLoading] = useState(false);
  const [savingAnnotatedPdf, setSavingAnnotatedPdf] = useState(false);
  const [activeTool, setActiveTool] = useState('view');
  const [activeColor, setActiveColor] = useState(DEFAULT_COLOR);
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [showToolbar, setShowToolbar] = useState(true);
  const [totalPages, setTotalPages] = useState(0);
  const [textNoteModal, setTextNoteModal] = useState(null);
  const [noteText, setNoteText] = useState('');
  const [showPageEditorModal, setShowPageEditorModal] = useState(false);
  const [showConvertModal, setShowConvertModal] = useState(false);
  const [conversionTask, setConversionTask] = useState(null);
  const [conversionStatus, setConversionStatus] = useState('');
  const [officePreviewMeta, setOfficePreviewMeta] = useState({
    previewKind: '',
    canConvertToPdf: false,
    isEditable: false,
  });
  const [officePdfBusy, setOfficePdfBusy] = useState(false);
  const [pageNumberInput, setPageNumberInput] = useState('1');
  const [pageEditorBusy, setPageEditorBusy] = useState(false);
  const [viewerInstanceId, setViewerInstanceId] = useState(0);
  const [showRenameModal, setShowRenameModal] = useState(false);
  const [renameDraft, setRenameDraft] = useState('');
  const [customColorModalVisible, setCustomColorModalVisible] = useState(false);
  const [customColorModalTitle, setCustomColorModalTitle] =
    useState('Annotation color');
  const customColorTargetRef = useRef('pdf');
  const [showPdfOverflowMenu, setShowPdfOverflowMenu] = useState(false);
  const [drawBrushPx, setDrawBrushPx] = useState(6);
  const [markerBrushPx, setMarkerBrushPx] = useState(16);

  const insets = useSafeAreaInsets();

  const webViewRef = useRef(null);
  const conversionChunksRef = useRef({
    fileName: '',
    mimeType: '',
    chunks: [],
  });
  const annotationWriteQueueRef = useRef(Promise.resolve());
  const annotationDirectoryReadyRef = useRef(false);
  const officePreviewExportRef = useRef({
    requestId: null,
    chunks: [],
  });

  const ensureAnnotationStorageDirectory = useCallback(async () => {
    if (annotationDirectoryReadyRef.current) {
      return;
    }

    try {
      await LegacyFileSystem.makeDirectoryAsync(ANNOTATION_STORAGE_DIR, {
        intermediates: true,
      });
    } catch (error) {
      if (!String(error?.message || '').toLowerCase().includes('exists')) {
        throw error;
      }
    }

    annotationDirectoryReadyRef.current = true;
  }, []);

  const buildViewerHtml = useCallback(
    (nextBase64, nextAnnotations, nextMode = 'view', nextSelectedPage = 1) =>
      getPdfViewerHtml(nextBase64, {
        initialAnnotations: nextAnnotations,
        initialMode: nextMode,
        initialDrawColor: activeColor,
        initialHighlightColor: activeColor + HIGHLIGHT_ALPHA_HEX,
        initialSelectedPage: nextSelectedPage,
      }),
    [activeColor]
  );

  const writeAnnotationSnapshot = useCallback(
    async (nextDocumentKey, nextAnnotations) => {
      if (!nextDocumentKey) {
        return;
      }

      const normalizedAnnotations = normalizeAnnotationState(nextAnnotations);
      await ensureAnnotationStorageDirectory();
      const annotationUri = getAnnotationStorageUri(nextDocumentKey);

      if (!hasAnnotations(normalizedAnnotations)) {
        const info = await LegacyFileSystem.getInfoAsync(annotationUri);
        if (info.exists) {
          await LegacyFileSystem.deleteAsync(annotationUri);
        }
        return;
      }

      await LegacyFileSystem.writeAsStringAsync(
        annotationUri,
        JSON.stringify({
          version: 1,
          annotations: normalizedAnnotations,
        })
      );
    },
    [ensureAnnotationStorageDirectory]
  );

  const writeWorkingPdfCopy = useCallback(
    async (nextDocumentKey, nextBase64) => {
      if (!nextDocumentKey || !nextBase64) {
        return;
      }

      await ensureAnnotationStorageDirectory();
      const workingPdfUri = getWorkingPdfStorageUri(nextDocumentKey);
      await LegacyFileSystem.writeAsStringAsync(workingPdfUri, nextBase64, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });
    },
    [ensureAnnotationStorageDirectory]
  );

  const loadWorkingPdfCopy = useCallback(
    async (nextDocumentKey) => {
      if (!nextDocumentKey) {
        return null;
      }

      await ensureAnnotationStorageDirectory();
      const workingPdfUri = getWorkingPdfStorageUri(nextDocumentKey);
      const info = await LegacyFileSystem.getInfoAsync(workingPdfUri);

      if (!info.exists) {
        return null;
      }

      return LegacyFileSystem.readAsStringAsync(workingPdfUri, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });
    },
    [ensureAnnotationStorageDirectory]
  );

  const clearPersistedPdfEditState = useCallback(
    async (nextDocumentKey) => {
      if (!nextDocumentKey) {
        return;
      }

      await ensureAnnotationStorageDirectory();
      await Promise.all([
        deleteFileIfExists(getAnnotationStorageUri(nextDocumentKey)),
        deleteFileIfExists(getWorkingPdfStorageUri(nextDocumentKey)),
      ]);
    },
    [ensureAnnotationStorageDirectory]
  );

  const promptForPdfResumeChoice = useCallback(
    (nextDocumentName) =>
      new Promise((resolve) => {
        let didResolve = false;
        const finish = (choice) => {
          if (!didResolve) {
            didResolve = true;
            resolve(choice);
          }
        };

        Alert.alert(
          'Resume Previous Edit?',
          `We found a saved local working copy for ${nextDocumentName}. Continue editing that copy, or reset back to the original PDF?`,
          [
            {
              text: 'Cancel',
              style: 'cancel',
              onPress: () => finish('cancel'),
            },
            {
              text: 'Open Original',
              style: 'destructive',
              onPress: () => finish('reset'),
            },
            {
              text: 'Continue Edit',
              onPress: () => finish('continue'),
            },
          ],
          {
            cancelable: true,
            onDismiss: () => finish('cancel'),
          }
        );
      }),
    []
  );

  const loadEditablePdfBase64 = useCallback(async () => {
    if (documentKey) {
      const workingPdfBase64 = await loadWorkingPdfCopy(documentKey);
      if (workingPdfBase64) {
        return workingPdfBase64;
      }
    }

    return base64Data;
  }, [base64Data, documentKey, loadWorkingPdfCopy]);

  const buildEditedPdfBase64 = useCallback(async () => {
    const editablePdfBase64 = await loadEditablePdfBase64();

    if (!editablePdfBase64) {
      throw new Error('No editable PDF data found.');
    }

    return bakePdfWithAnnotations(editablePdfBase64, annotationState);
  }, [annotationState, loadEditablePdfBase64]);

  const buildEditedPdfFile = useCallback(
    async (outputDirectoryUri = Paths.cache.uri) => {
      const bakedPdfBase64 = await buildEditedPdfBase64();
      const fileName = buildAnnotatedFileName(documentName);
      const outputUri = `${outputDirectoryUri}${fileName}`;

      await LegacyFileSystem.writeAsStringAsync(outputUri, bakedPdfBase64, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });

      return {
        bakedPdfBase64,
        fileName,
        outputUri,
      };
    },
    [buildEditedPdfBase64, documentName]
  );

  const promptForPdfShareChoice = useCallback(
    () =>
      new Promise((resolve) => {
        let didResolve = false;
        const finish = (choice) => {
          if (!didResolve) {
            didResolve = true;
            resolve(choice);
          }
        };

        Alert.alert(
          'Share PDF',
          'Choose whether to share the original imported PDF or the current edited PDF copy.',
          [
            {
              text: 'Cancel',
              style: 'cancel',
              onPress: () => finish('cancel'),
            },
            {
              text: 'Share Original',
              onPress: () => finish('original'),
            },
            {
              text: 'Share Edited PDF',
              onPress: () => finish('edited'),
            },
          ],
          {
            cancelable: true,
            onDismiss: () => finish('cancel'),
          }
        );
      }),
    []
  );

  const resetConversionState = useCallback(() => {
    conversionChunksRef.current = {
      fileName: '',
      mimeType: '',
      chunks: [],
    };
    setConversionTask(null);
    setConversionStatus('');
  }, []);

  const resetOfficePreviewState = useCallback(() => {
    officePreviewExportRef.current = {
      requestId: null,
      chunks: [],
    };
    setOfficePreviewMeta({
      previewKind: '',
      canConvertToPdf: false,
      isEditable: false,
    });
    setOfficePdfBusy(false);
  }, []);

  const openRenameModal = useCallback(() => {
    if (!documentUri) {
      return;
    }
    setRenameDraft(documentName);
    setShowRenameModal(true);
  }, [documentName, documentUri]);

  const applyDocumentRename = useCallback(() => {
    const trimmed = renameDraft.trim();
    if (!trimmed) {
      Alert.alert('Name required', 'Enter a file name.');
      return;
    }
    let next = trimmed;
    if (documentType && !trimmed.includes('.')) {
      next = `${trimmed}.${documentType}`;
    }
    setDocumentName(next);
    setShowRenameModal(false);
  }, [renameDraft, documentType]);

  const persistAnnotationState = useCallback(
    (nextDocumentKey, nextAnnotations) => {
      if (!nextDocumentKey) {
        return;
      }

      annotationWriteQueueRef.current = annotationWriteQueueRef.current
        .then(() => writeAnnotationSnapshot(nextDocumentKey, nextAnnotations))
        .catch((error) => {
          Alert.alert(
            'Annotation Save Error',
            'Failed to save annotations: ' + error.message
          );
        });
    },
    [writeAnnotationSnapshot]
  );

  const loadPersistedAnnotations = useCallback(
    async (nextDocumentKey) => {
      if (!nextDocumentKey) {
        return {};
      }

      await ensureAnnotationStorageDirectory();
      const annotationUri = getAnnotationStorageUri(nextDocumentKey);
      const info = await LegacyFileSystem.getInfoAsync(annotationUri);

      if (!info.exists) {
        return {};
      }

      const rawContents = await LegacyFileSystem.readAsStringAsync(annotationUri);
      const parsed = JSON.parse(rawContents);
      return normalizeAnnotationState(parsed.annotations || parsed);
    },
    [ensureAnnotationStorageDirectory]
  );

  const applyPdfEditorState = useCallback(
    (
      nextBase64,
      nextAnnotations,
      nextPageCount,
      nextMode = 'view',
      nextSelectedPage = 1
    ) => {
      const safeSelectedPage = Math.max(
        1,
        Math.min(nextSelectedPage, Math.max(nextPageCount, 1))
      );
      setBase64Data(nextBase64);
      setAnnotationState(nextAnnotations);
      setSelectedAnnotation(null);
      setActiveTool(nextMode);
      setShowColorPicker(
        nextMode === 'draw' || nextMode === 'highlight' || nextMode === 'marker'
      );
      setTotalPages(nextPageCount);
      setPageNumberInput(String(safeSelectedPage));
      setViewerHtml(
        buildViewerHtml(nextBase64, nextAnnotations, nextMode, safeSelectedPage)
      );
      setViewerInstanceId((currentValue) => currentValue + 1);
    },
    [buildViewerHtml]
  );

  const sendCommand = useCallback((command) => {
    if (!webViewRef.current) {
      return;
    }

    const js = `
      (function() {
        try {
          var cmd = ${JSON.stringify(command)};
          if (typeof window.dispatchViewerCommand === 'function') {
            window.dispatchViewerCommand(cmd);
          }
        } catch (error) {
          console.error('Command error:', error);
        }
      })();
      true;
    `;

    webViewRef.current.injectJavaScript(js);
  }, []);

  const pickDocument = useCallback(async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: [
          'application/pdf',
          'application/msword',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'application/vnd.ms-powerpoint',
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'application/vnd.ms-excel',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'text/plain',
          'image/*',
        ],
        copyToCacheDirectory: true,
      });

      if (result.canceled) {
        return;
      }

      const file = result.assets[0];
      const extension = file.name.split('.').pop().toLowerCase();

      setLoading(true);

      if (extension === 'pdf') {
        const fileInfo = await LegacyFileSystem.getInfoAsync(file.uri, { md5: true });
        const nextDocumentKey =
          fileInfo.md5 || `${file.name}-${fileInfo.size || Date.now()}`;
        const [originalBase64, persistedAnnotations, workingPdfBase64] = await Promise.all([
          LegacyFileSystem.readAsStringAsync(file.uri, {
            encoding: LegacyFileSystem.EncodingType.Base64,
          }),
          loadPersistedAnnotations(nextDocumentKey),
          loadWorkingPdfCopy(nextDocumentKey),
        ]);
        const hasSavedEdits =
          !!workingPdfBase64 || hasAnnotations(persistedAnnotations);
        let nextAnnotations = persistedAnnotations;
        let activeBase64 = workingPdfBase64 || originalBase64;

        if (hasSavedEdits) {
          const resumeChoice = await promptForPdfResumeChoice(file.name);

          if (resumeChoice === 'cancel') {
            return;
          }

          if (resumeChoice === 'reset') {
            await annotationWriteQueueRef.current;
            await clearPersistedPdfEditState(nextDocumentKey);
            nextAnnotations = {};
            activeBase64 = originalBase64;
          }
        }

        setDocumentUri(file.uri);
        setDocumentName(file.name);
        setDocumentType(extension);
        setDocumentKey(nextDocumentKey);
        setBase64Data(activeBase64);
        setTotalPages(0);
        setTextNoteModal(null);
        setNoteText('');
        setActiveTool('view');
        setActiveColor(DEFAULT_COLOR);
        setShowColorPicker(false);
        setShowToolbar(true);
        setSelectedAnnotation(null);
        setAnnotationState(nextAnnotations);
        setViewerHtml(buildViewerHtml(activeBase64, nextAnnotations, 'view', 1));
        setViewerInstanceId((currentValue) => currentValue + 1);
        setShowPageEditorModal(false);
        setShowConvertModal(false);
        resetConversionState();
        resetOfficePreviewState();
        setPageNumberInput('1');
      } else {
        const shouldLoadPreviewBase64 = OFFICE_PREVIEW_TYPES.includes(extension);
        const nextBase64 = shouldLoadPreviewBase64
          ? await LegacyFileSystem.readAsStringAsync(file.uri, {
              encoding: LegacyFileSystem.EncodingType.Base64,
            })
          : null;

        setDocumentUri(file.uri);
        setDocumentName(file.name);
        setDocumentType(extension);
        setDocumentKey(null);
        setBase64Data(nextBase64);
        setTotalPages(0);
        setTextNoteModal(null);
        setNoteText('');
        setActiveTool('view');
        setActiveColor(DEFAULT_COLOR);
        setShowColorPicker(false);
        setShowToolbar(true);
        setSelectedAnnotation(null);
        setAnnotationState({});
        setViewerHtml(null);
        setViewerInstanceId((currentValue) => currentValue + 1);
        setShowPageEditorModal(false);
        setShowConvertModal(false);
        resetConversionState();
        resetOfficePreviewState();
        setPageNumberInput('1');
      }
    } catch (error) {
      Alert.alert('Error', 'Failed to pick document: ' + error.message);
    } finally {
      setLoading(false);
    }
  }, [
    buildViewerHtml,
    clearPersistedPdfEditState,
    loadPersistedAnnotations,
    loadWorkingPdfCopy,
    promptForPdfResumeChoice,
    resetConversionState,
    resetOfficePreviewState,
  ]);

  const copyDocument = useCallback(async () => {
    if (!documentUri) {
      return;
    }

    try {
      const hasDot = documentName.includes('.');
      const extension = hasDot
        ? documentName.split('.').pop()
        : documentType;
      const baseName = hasDot
        ? documentName.slice(0, documentName.length - extension.length - 1)
        : documentName;
      const newName = `${baseName}_copy.${extension}`;
      const nextUri = Paths.document.uri + newName;

      await LegacyFileSystem.copyAsync({
        from: documentUri,
        to: nextUri,
      });

      Alert.alert('Success', 'Document copied as: ' + newName);
    } catch (error) {
      Alert.alert('Error', 'Failed to copy: ' + error.message);
    }
  }, [documentName, documentUri]);

  const shareDocument = useCallback(async () => {
    if (!documentUri) {
      return;
    }

    try {
      const isAvailable = await Sharing.isAvailableAsync();
      if (!isAvailable) {
        Alert.alert('Error', 'Sharing is not available on this device');
        return;
      }

      if (documentType !== 'pdf') {
        await Sharing.shareAsync(documentUri);
        return;
      }

      const shareChoice = await promptForPdfShareChoice();

      if (shareChoice === 'cancel') {
        return;
      }

      if (shareChoice === 'original') {
        await Sharing.shareAsync(documentUri);
        return;
      }

      const { outputUri } = await buildEditedPdfFile(Paths.cache.uri);
      await Sharing.shareAsync(outputUri);
    } catch (error) {
      Alert.alert('Error', 'Failed to share: ' + error.message);
    }
  }, [buildEditedPdfFile, documentType, documentUri, promptForPdfShareChoice]);

  const convertOfficePreviewToPdf = useCallback(() => {
    if (!officePreviewMeta.canConvertToPdf || officePdfBusy) {
      return;
    }

    const requestId = Date.now().toString();
    officePreviewExportRef.current = {
      requestId,
      chunks: [],
    };
    setOfficePdfBusy(true);

    sendCommand({
      command: 'exportPreview',
      requestId,
      title: documentName,
    });
  }, [documentName, officePdfBusy, officePreviewMeta.canConvertToPdf, sendCommand]);

  const sendOfficePreviewCommand = useCallback(
    (payload) => {
      sendCommand({
        command: 'officeFormat',
        ...payload,
      });
    },
    [sendCommand]
  );

  const applyOfficeStyleCommand = useCallback(
    (action, value) => {
      sendOfficePreviewCommand({ action, value });
    },
    [sendOfficePreviewCommand]
  );

  const applyOfficeTextColor = useCallback(
    (color) => {
      const parsed = parseHexStringToRgb(color);
      const normalized = parsed ? rgbToHex(parsed.r, parsed.g, parsed.b) : String(color);
      setActiveColor(normalized);
      sendOfficePreviewCommand({
        action: 'setColor',
        value: normalized,
      });
    },
    [sendOfficePreviewCommand]
  );

  const handleCustomColorApply = useCallback(
    (hex) => {
      if (customColorTargetRef.current === 'office') {
        applyOfficeTextColor(hex);
      } else {
        selectColor(hex);
      }
      setCustomColorModalVisible(false);
    },
    [applyOfficeTextColor, selectColor]
  );

  const insertImageIntoOfficePreview = useCallback(async () => {
    try {
      const imageResult = await DocumentPicker.getDocumentAsync({
        type: 'image/*',
        copyToCacheDirectory: true,
      });

      if (imageResult.canceled) {
        return;
      }

      const imageAsset = imageResult.assets[0];
      const imageBase64 = await LegacyFileSystem.readAsStringAsync(imageAsset.uri, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });
      const dataUrl = `data:${imageAsset.mimeType || 'image/png'};base64,${imageBase64}`;

      sendOfficePreviewCommand({
        action: 'insertImage',
        dataUrl,
      });
    } catch (error) {
      Alert.alert('Image Insert Failed', error.message);
    }
  }, [sendOfficePreviewCommand]);

  const startPdfConversion = useCallback(
    async (formatId) => {
      if (documentType !== 'pdf') {
        return;
      }

      const formatConfig = PDF_CONVERSION_FORMATS.find(
        (conversionFormat) => conversionFormat.id === formatId
      );

      if (!formatConfig) {
        return;
      }

      try {
        setShowConvertModal(false);
        setConversionStatus('Preparing edited PDF...');

        const bakedPdfBase64 = await buildEditedPdfBase64();
        const outputFileName = buildConvertedFileName(
          documentName,
          formatConfig.extension
        );

        conversionChunksRef.current = {
          fileName: outputFileName,
          mimeType: formatConfig.mimeType,
          chunks: [],
        };

        setConversionTask({
          id: Date.now().toString(),
          formatId,
          formatLabel: formatConfig.label,
          fileName: outputFileName,
          mimeType: formatConfig.mimeType,
          html: getPdfConverterHtml(bakedPdfBase64, {
            format: formatId,
            fileName: outputFileName,
            mimeType: formatConfig.mimeType,
          }),
        });
      } catch (error) {
        resetConversionState();
        Alert.alert('Conversion Failed', error.message);
      }
    },
    [buildEditedPdfBase64, documentName, documentType, resetConversionState]
  );

  const finalizeOfficePreviewPdfExport = useCallback(
    async (html) => {
      const convertedFileName = buildConvertedFileName(documentName, 'pdf');
      const printResult = await Print.printToFileAsync({
        html,
      });
      const convertedPdfBase64 = await LegacyFileSystem.readAsStringAsync(
        printResult.uri,
        {
          encoding: LegacyFileSystem.EncodingType.Base64,
        }
      );
      const outputUri = `${Paths.document.uri}${convertedFileName}`;

      await LegacyFileSystem.writeAsStringAsync(outputUri, convertedPdfBase64, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });

      const deviceSaveResult =
        Platform.OS === 'android'
          ? await savePdfToAndroidDeviceFolder(convertedPdfBase64, convertedFileName)
          : { savedToDevice: false, reason: 'unsupported_platform' };

      const canShare = await Sharing.isAvailableAsync();
      const saveSummary = [`Saved in app storage as ${convertedFileName}.`];

      if (deviceSaveResult.savedToDevice) {
        saveSummary.push(
          'A second copy was also saved to the folder you picked on your phone.'
        );
      } else if (Platform.OS === 'android') {
        saveSummary.push(
          'The visible device copy was not created. Next time, allow the folder picker to save directly into phone storage.'
        );
      }

      Alert.alert('PDF Created', saveSummary.join(' '), [
        ...(canShare
          ? [
              {
                text: 'Share',
                onPress: () => Sharing.shareAsync(outputUri),
              },
            ]
          : []),
        { text: 'OK' },
      ]);
    },
    [documentName]
  );

  const handleConverterMessage = useCallback(
    async (event) => {
      try {
        const data = JSON.parse(event.nativeEvent.data);

        switch (data.type) {
          case 'conversionStatus':
            setConversionStatus(data.message || 'Converting document...');
            break;
          case 'conversionMeta':
            conversionChunksRef.current = {
              fileName: data.fileName || '',
              mimeType: data.mimeType || 'application/octet-stream',
              chunks: new Array(Math.max(Number(data.totalChunks) || 0, 0)).fill(''),
            };
            break;
          case 'conversionChunk':
            if (
              Number.isInteger(data.index) &&
              conversionChunksRef.current.chunks[data.index] !== undefined
            ) {
              conversionChunksRef.current.chunks[data.index] = data.chunk || '';
            }
            break;
          case 'conversionComplete': {
            const { fileName, mimeType, chunks } = conversionChunksRef.current;
            const convertedBase64 = chunks.join('');

            if (!fileName || !convertedBase64) {
              throw new Error('Converted document data was incomplete.');
            }

            const outputUri = `${Paths.document.uri}${fileName}`;
            await LegacyFileSystem.writeAsStringAsync(outputUri, convertedBase64, {
              encoding: LegacyFileSystem.EncodingType.Base64,
            });

            const deviceSaveResult =
              Platform.OS === 'android'
                ? await saveBase64ToAndroidDeviceFolder(
                    convertedBase64,
                    fileName,
                    mimeType
                  )
                : { savedToDevice: false, reason: 'unsupported_platform' };

            const canShare = await Sharing.isAvailableAsync();
            const saveSummary = [`Saved in app storage as ${fileName}.`];

            if (deviceSaveResult.savedToDevice) {
              saveSummary.push(
                'A second copy was also saved to the folder you picked on your phone.'
              );
            } else if (Platform.OS === 'android') {
              saveSummary.push(
                'The visible device copy was not created. Next time, allow the folder picker to save directly into phone storage.'
              );
            }

            resetConversionState();

            Alert.alert('Document Converted', saveSummary.join(' '), [
              ...(canShare
                ? [
                    {
                      text: 'Share',
                      onPress: () =>
                        Sharing.shareAsync(outputUri, {
                          mimeType,
                        }),
                    },
                  ]
                : []),
              { text: 'OK' },
            ]);
            break;
          }
          case 'conversionError':
            resetConversionState();
            Alert.alert(
              'Conversion Failed',
              data.message ||
                'The document could not be converted. Check your connection and try again.'
            );
            break;
        }
      } catch (error) {
        resetConversionState();
        Alert.alert('Conversion Failed', error.message);
      }
    },
    [resetConversionState]
  );

  const selectTool = useCallback(
    (toolId) => {
      setActiveTool(toolId);
      setSelectedAnnotation(null);
      setShowColorPicker(
        toolId === 'draw' || toolId === 'highlight' || toolId === 'marker'
      );
      sendCommand({ command: 'setMode', mode: toolId });
    },
    [sendCommand]
  );

  const selectColor = useCallback(
    (color) => {
      const parsed = parseHexStringToRgb(color);
      const normalized = parsed ? rgbToHex(parsed.r, parsed.g, parsed.b) : String(color);
      setActiveColor(normalized);
      // Always sync both to the WebView so pen, area highlight, and marker stay in
      // sync (injected scripts call window.dispatchViewerCommand, not page globals).
      sendCommand({ command: 'setDrawColor', color: normalized });
      sendCommand({
        command: 'setHighlightColor',
        color: normalized + HIGHLIGHT_ALPHA_HEX,
      });
    },
    [sendCommand]
  );

  const openCustomColorModal = useCallback((target) => {
    customColorTargetRef.current = target;
    setCustomColorModalTitle(
      target === 'office' ? 'Text & highlight color' : 'Annotation color'
    );
    setCustomColorModalVisible(true);
  }, []);

  const deleteSelectedAnnotation = useCallback(() => {
    if (!selectedAnnotation) {
      return;
    }

    sendCommand({ command: 'deleteSelected' });
  }, [selectedAnnotation, sendCommand]);

  const clearAll = useCallback(() => {
    Alert.alert('Clear Annotations', 'Remove all highlights, drawings, and notes?', [
      { text: 'Cancel', style: 'cancel' },
      {
        text: 'Clear All',
        style: 'destructive',
        onPress: () => sendCommand({ command: 'clearAll' }),
      },
    ]);
  }, [sendCommand]);

  const getValidatedPageNumber = useCallback(
    (allowInsertAfterLast = false) => {
      const parsedPageNumber = Number.parseInt(pageNumberInput, 10);
      const upperBound = allowInsertAfterLast ? totalPages : totalPages;

      if (!Number.isInteger(parsedPageNumber)) {
        Alert.alert(
          'Invalid Page',
          'Tap a page in the PDF first so I know which page to edit.'
        );
        return null;
      }

      if (parsedPageNumber < 1 || parsedPageNumber > upperBound) {
        Alert.alert(
          'Invalid Page',
          `Choose a page number between 1 and ${upperBound}.`
        );
        return null;
      }

      return parsedPageNumber;
    },
    [pageNumberInput, totalPages]
  );

  const removePdfPage = useCallback(async () => {
    const pageNumber = getValidatedPageNumber(false);
    if (!pageNumber || !base64Data || !documentKey) {
      return;
    }

    if (totalPages <= 1) {
      Alert.alert('Cannot Remove Page', 'A PDF must keep at least one page.');
      return;
    }

    try {
      setPageEditorBusy(true);
      const pdfDocument = await PDFDocument.load(base64Data);
      pdfDocument.removePage(pageNumber - 1);

      const nextAnnotations = removePageFromAnnotationState(annotationState, pageNumber);
      const nextBase64 = await pdfDocument.saveAsBase64();
      const nextPageCount = pdfDocument.getPageCount();
      const nextSelectedPage = Math.max(1, Math.min(pageNumber, nextPageCount));

      await annotationWriteQueueRef.current;
      await writeWorkingPdfCopy(documentKey, nextBase64);
      await writeAnnotationSnapshot(documentKey, nextAnnotations);
      applyPdfEditorState(
        nextBase64,
        nextAnnotations,
        nextPageCount,
        'view',
        nextSelectedPage
      );
      setShowPageEditorModal(false);
    } catch (error) {
      Alert.alert('Page Update Failed', error.message);
    } finally {
      setPageEditorBusy(false);
    }
  }, [
    annotationState,
    applyPdfEditorState,
    base64Data,
    documentKey,
    getValidatedPageNumber,
    totalPages,
    writeAnnotationSnapshot,
    writeWorkingPdfCopy,
  ]);

  const insertBlankPdfPage = useCallback(async () => {
    const pageNumber = getValidatedPageNumber(true);
    if (!pageNumber || !base64Data || !documentKey) {
      return;
    }

    try {
      setPageEditorBusy(true);
      const pdfDocument = await PDFDocument.load(base64Data);
      const pages = pdfDocument.getPages();
      const referencePage = pages[Math.min(pageNumber - 1, pages.length - 1)];
      const { width, height } = referencePage.getSize();

      pdfDocument.insertPage(pageNumber, [width, height]);

      const nextAnnotations = insertPageIntoAnnotationState(annotationState, pageNumber);
      const nextBase64 = await pdfDocument.saveAsBase64();
      const nextPageCount = pdfDocument.getPageCount();
      const nextSelectedPage = Math.min(pageNumber + 1, nextPageCount);

      await annotationWriteQueueRef.current;
      await writeWorkingPdfCopy(documentKey, nextBase64);
      await writeAnnotationSnapshot(documentKey, nextAnnotations);
      applyPdfEditorState(
        nextBase64,
        nextAnnotations,
        nextPageCount,
        'view',
        nextSelectedPage
      );
      setShowPageEditorModal(false);
    } catch (error) {
      Alert.alert('Page Insert Failed', error.message);
    } finally {
      setPageEditorBusy(false);
    }
  }, [
    annotationState,
    applyPdfEditorState,
    base64Data,
    documentKey,
    getValidatedPageNumber,
    writeAnnotationSnapshot,
    writeWorkingPdfCopy,
  ]);

  const insertImagePdfPage = useCallback(async () => {
    const pageNumber = getValidatedPageNumber(true);
    if (!pageNumber || !base64Data || !documentKey) {
      return;
    }

    try {
      setPageEditorBusy(true);
      const imageResult = await DocumentPicker.getDocumentAsync({
        type: ['image/png', 'image/jpeg'],
        copyToCacheDirectory: true,
      });

      if (imageResult.canceled) {
        return;
      }

      const imageAsset = imageResult.assets[0];
      const mimeType = imageAsset.mimeType || '';
      const extension = imageAsset.name.split('.').pop().toLowerCase();
      const isPng = mimeType === 'image/png' || extension === 'png';
      const isJpeg =
        mimeType === 'image/jpeg' || extension === 'jpg' || extension === 'jpeg';

      if (!isPng && !isJpeg) {
        Alert.alert('Unsupported Image', 'Please choose a PNG or JPEG image.');
        return;
      }

      const pdfDocument = await PDFDocument.load(base64Data);
      const pages = pdfDocument.getPages();
      const referencePage = pages[Math.min(pageNumber - 1, pages.length - 1)];
      const { width, height } = referencePage.getSize();
      const newPage = pdfDocument.insertPage(pageNumber, [width, height]);
      const imageBase64 = await LegacyFileSystem.readAsStringAsync(imageAsset.uri, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });
      const embeddedImage = isPng
        ? await pdfDocument.embedPng(imageBase64)
        : await pdfDocument.embedJpg(imageBase64);
      const imageDimensions = embeddedImage.scale(1);
      const margin = 24;
      const scaleFactor = Math.min(
        (width - margin * 2) / imageDimensions.width,
        (height - margin * 2) / imageDimensions.height,
        1
      );
      const drawWidth = imageDimensions.width * scaleFactor;
      const drawHeight = imageDimensions.height * scaleFactor;

      newPage.drawImage(embeddedImage, {
        x: (width - drawWidth) / 2,
        y: (height - drawHeight) / 2,
        width: drawWidth,
        height: drawHeight,
      });

      const nextAnnotations = insertPageIntoAnnotationState(annotationState, pageNumber);
      const nextBase64 = await pdfDocument.saveAsBase64();
      const nextPageCount = pdfDocument.getPageCount();
      const nextSelectedPage = Math.min(pageNumber + 1, nextPageCount);

      await annotationWriteQueueRef.current;
      await writeWorkingPdfCopy(documentKey, nextBase64);
      await writeAnnotationSnapshot(documentKey, nextAnnotations);
      applyPdfEditorState(
        nextBase64,
        nextAnnotations,
        nextPageCount,
        'view',
        nextSelectedPage
      );
      setShowPageEditorModal(false);
    } catch (error) {
      Alert.alert('Image Page Failed', error.message);
    } finally {
      setPageEditorBusy(false);
    }
  }, [
    annotationState,
    applyPdfEditorState,
    base64Data,
    documentKey,
    getValidatedPageNumber,
    writeAnnotationSnapshot,
    writeWorkingPdfCopy,
  ]);

  const saveAnnotatedPdf = useCallback(async () => {
    if (documentType !== 'pdf') {
      return;
    }

    try {
      setSavingAnnotatedPdf(true);
      const { bakedPdfBase64, fileName, outputUri } = await buildEditedPdfFile(
        Paths.document.uri
      );

      const deviceSaveResult =
        Platform.OS === 'android'
          ? await savePdfToAndroidDeviceFolder(bakedPdfBase64, fileName)
          : { savedToDevice: false, reason: 'unsupported_platform' };

      const canShare = await Sharing.isAvailableAsync();
      const saveSummary = [`Saved in app storage as ${fileName}.`];

      if (deviceSaveResult.savedToDevice) {
        saveSummary.push(
          'A second copy was also saved to the folder you picked on your phone.'
        );
      } else if (Platform.OS === 'android') {
        saveSummary.push(
          'The visible device copy was not created. Next time, allow the folder picker to save directly into phone storage.'
        );
      }

      Alert.alert('Annotated PDF Saved', saveSummary.join(' '), [
        ...(canShare
          ? [
              {
                text: 'Share',
                onPress: () => Sharing.shareAsync(outputUri),
              },
            ]
          : []),
        { text: 'OK' },
      ]);
    } catch (error) {
      Alert.alert(
        'Save Failed',
        'Could not save the annotated PDF: ' + error.message
      );
    } finally {
      setSavingAnnotatedPdf(false);
    }
  }, [buildEditedPdfFile, documentType]);

  const handleWebViewMessage = useCallback(
    (event) => {
      try {
        const data = JSON.parse(event.nativeEvent.data);

        switch (data.type) {
          case 'pdfLoaded':
            setTotalPages(data.totalPages);
            sendCommand({ command: 'setMode', mode: activeTool });
            sendCommand({ command: 'setDrawColor', color: activeColor });
            sendCommand({
              command: 'setHighlightColor',
              color: activeColor + HIGHLIGHT_ALPHA_HEX,
            });
            sendCommand({ command: 'setDrawBrushWidth', px: drawBrushPx });
            sendCommand({ command: 'setMarkerBrushWidth', px: markerBrushPx });
            break;
          case 'requestTextInput':
            setTextNoteModal({ page: data.page, x: data.x, y: data.y });
            setNoteText('');
            break;
          case 'annotationSelected':
            setSelectedAnnotation(data.annotation || null);
            break;
          case 'pageSelected':
            if (Number.isInteger(data.page)) {
              setPageNumberInput(String(data.page));
            }
            break;
          case 'pageLongPressed':
            if (Number.isInteger(data.page)) {
              setPageNumberInput(String(data.page));
            }
            setShowPageEditorModal(true);
            break;
          case 'officePreviewReady':
            setOfficePreviewMeta({
              previewKind: data.previewKind || '',
              canConvertToPdf: !!data.canConvertToPdf,
              isEditable: !!data.isEditable,
            });
            setOfficePdfBusy(false);
            break;
          case 'officePreviewExportMeta':
            if (data.requestId === officePreviewExportRef.current.requestId) {
              officePreviewExportRef.current = {
                requestId: data.requestId,
                chunks: new Array(Math.max(Number(data.totalChunks) || 0, 0)).fill(''),
              };
            }
            break;
          case 'officePreviewExportChunk':
            if (
              data.requestId === officePreviewExportRef.current.requestId &&
              Number.isInteger(data.index) &&
              officePreviewExportRef.current.chunks[data.index] !== undefined
            ) {
              officePreviewExportRef.current.chunks[data.index] = data.chunk || '';
            }
            break;
          case 'officePreviewExportComplete':
            if (data.requestId === officePreviewExportRef.current.requestId) {
              const html = officePreviewExportRef.current.chunks.join('');
              officePreviewExportRef.current = {
                requestId: null,
                chunks: [],
              };

              if (!html) {
                setOfficePdfBusy(false);
                Alert.alert('PDF Conversion Failed', 'Office preview export was empty.');
                break;
              }

              finalizeOfficePreviewPdfExport(html)
                .catch((error) => {
                  Alert.alert('PDF Conversion Failed', error.message);
                })
                .finally(() => {
                  setOfficePdfBusy(false);
                });
            }
            break;
          case 'officePreviewExportError':
            if (data.requestId === officePreviewExportRef.current.requestId) {
              officePreviewExportRef.current = {
                requestId: null,
                chunks: [],
              };
              setOfficePdfBusy(false);
              Alert.alert(
                'PDF Conversion Failed',
                data.message || 'This document preview could not be exported.'
              );
            }
            break;
          case 'officePreviewCommandError':
            Alert.alert(
              'Office Edit',
              data.message || 'Select an editable area in the document first.'
            );
            break;
          case 'annotationsChanged': {
            const nextAnnotations = normalizeAnnotationState(data.annotations);
            setAnnotationState(nextAnnotations);
            persistAnnotationState(documentKey, nextAnnotations);
            setSelectedAnnotation((currentSelection) => {
              if (!currentSelection) {
                return currentSelection;
              }

              const pageAnnotations = nextAnnotations[currentSelection.page];
              if (!pageAnnotations) {
                return null;
              }

              const bucket =
                currentSelection.type === 'highlight'
                  ? pageAnnotations.highlights
                  : currentSelection.type === 'drawing'
                    ? pageAnnotations.drawings
                    : pageAnnotations.notes;

              return bucket.some(
                (annotation) => annotation.id === currentSelection.id
              )
                ? currentSelection
                : null;
            });
            break;
          }
          case 'modeChanged':
            if (data.mode !== 'view') {
              setSelectedAnnotation(null);
            }
            break;
          case 'error':
            Alert.alert('PDF Error', data.message);
            break;
        }
      } catch {}
    },
    [
      activeColor,
      activeTool,
      documentKey,
      drawBrushPx,
      markerBrushPx,
      finalizeOfficePreviewPdfExport,
      persistAnnotationState,
      sendCommand,
    ]
  );

  useEffect(() => {
    if (documentType !== 'pdf') {
      return;
    }
    sendCommand({ command: 'setDrawBrushWidth', px: drawBrushPx });
  }, [documentType, drawBrushPx, sendCommand]);

  useEffect(() => {
    if (documentType !== 'pdf') {
      return;
    }
    sendCommand({ command: 'setMarkerBrushWidth', px: markerBrushPx });
  }, [documentType, markerBrushPx, sendCommand]);

  const submitTextNote = useCallback(() => {
    if (textNoteModal && noteText.trim()) {
      sendCommand({
        command: 'addTextNote',
        page: textNoteModal.page,
        x: textNoteModal.x,
        y: textNoteModal.y,
        text: noteText.trim(),
      });
    }

    setTextNoteModal(null);
    setNoteText('');
  }, [noteText, sendCommand, textNoteModal]);

  const renderNonPdfViewer = () => {
    const extension = documentType;

    if (['txt'].includes(extension)) {
      return (
        <View style={styles.nonPdfContainer}>
          <TextFileViewer uri={documentUri} />
        </View>
      );
    }

    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'].includes(extension)) {
      return (
        <View style={styles.nonPdfContainer}>
          <WebView
            source={{ uri: documentUri }}
            style={styles.webView}
            scalesPageToFit={true}
          />
        </View>
      );
    }

    if (['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx'].includes(extension)) {
      return (
        <View style={styles.nonPdfContainer}>
          <OfficeDocViewer
            uri={documentUri}
            name={documentName}
            extension={documentType}
            base64Data={base64Data}
            onMessage={handleWebViewMessage}
            webViewRef={webViewRef}
          />
        </View>
      );
    }

    return (
      <View style={styles.unsupported}>
        <Text style={styles.unsupportedText}>
          Preview not available for .{extension} files.
        </Text>
        <Text style={styles.unsupportedHint}>
          You can still copy or share this document using the buttons above.
        </Text>
      </View>
    );
  };

  return (
    <View style={styles.container}>
      <StatusBar style="dark" backgroundColor={colors.surface} />

      <View
        style={[
          styles.header,
          {
            paddingTop: Math.max(
              insets.top,
              Platform.OS === 'android' ? 12 : 10,
            ),
          },
        ]}
      >
        <TouchableOpacity
          style={styles.headerTitleBlock}
          onPress={openRenameModal}
          disabled={!documentUri}
          activeOpacity={documentUri ? 0.65 : 1}
        >
          <Text style={styles.headerTitle} numberOfLines={1}>
            {documentName || 'Electric PDF Studio'}
          </Text>
          {documentUri ? (
            <MaterialIcons
              name="edit"
              size={18}
              color={colors.primary}
              style={styles.headerEditIcon}
            />
          ) : null}
        </TouchableOpacity>
        {totalPages > 0 ? (
          <View style={styles.pageCountPill}>
            <Text style={styles.pageCountLabel}>{totalPages}</Text>
            <Text style={styles.pageCountSuffix}>pages</Text>
          </View>
        ) : null}
      </View>

      {!documentUri && !loading && (
        <View style={styles.landing}>
          <MaterialIcons
            name="picture-as-pdf"
            size={64}
            color={colors.primary}
            style={styles.landingIcon}
          />
          <Text style={styles.landingTitle}>Document Viewer & Editor</Text>
          <Text style={styles.landingSubtitle}>
            Import PDFs, Word docs, PowerPoint, Excel, images, and text files
            from your device. View, annotate, and edit with built-in tools.
          </Text>
          <TouchableOpacity style={styles.importButton} onPress={pickDocument}>
            <View style={styles.importButtonRow}>
              <MaterialIcons name="folder-open" size={22} color={colors.onPrimary} />
              <Text style={styles.importButtonText}>Import Document</Text>
            </View>
          </TouchableOpacity>
          <Text style={styles.supportedFormats}>
            Supported: PDF, DOC, DOCX, PPT, PPTX, XLS, XLSX, TXT, Images
          </Text>
        </View>
      )}

      {loading && (
        <View style={styles.landing}>
          <ActivityIndicator size="large" color={colors.primary} />
          <Text style={styles.loadingText}>Loading document...</Text>
        </View>
      )}

      {documentUri && !loading && (
        <View style={styles.documentArea}>
          <View style={styles.actionBar}>
            <ScrollView
              horizontal
              showsHorizontalScrollIndicator={false}
              contentContainerStyle={styles.actionBarScroll}
            >
              <TouchableOpacity
                style={styles.actionIconBtn}
                onPress={pickDocument}
                accessibilityLabel="Open document"
              >
                <MaterialIcons name="folder-open" size={22} color={colors.primary} />
              </TouchableOpacity>
              <TouchableOpacity
                style={styles.actionIconBtn}
                onPress={copyDocument}
                accessibilityLabel="Copy document"
              >
                <MaterialIcons name="content-copy" size={22} color={colors.onSurface} />
              </TouchableOpacity>
              <TouchableOpacity
                style={styles.actionIconBtn}
                onPress={shareDocument}
                accessibilityLabel="Share document"
              >
                <MaterialIcons name="share" size={22} color={colors.onSurface} />
              </TouchableOpacity>
              {documentType === 'pdf' && (
                <>
                  <View style={styles.actionBarDivider} />
                  <TouchableOpacity
                    style={[
                      styles.actionIconBtn,
                      !!conversionTask && styles.actionIconBtnDisabled,
                    ]}
                    disabled={!!conversionTask}
                    onPress={() => setShowConvertModal(true)}
                    accessibilityLabel="Convert PDF"
                  >
                    <MaterialIcons
                      name="transform"
                      size={22}
                      color={colors.onSurface}
                    />
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.actionIconBtn,
                      pageEditorBusy && styles.actionIconBtnDisabled,
                    ]}
                    disabled={pageEditorBusy}
                    onPress={() => setShowPageEditorModal(true)}
                    accessibilityLabel="Edit pages"
                  >
                    <MaterialIcons name="layers" size={22} color={colors.onSurface} />
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.actionIconBtnPrimary,
                      savingAnnotatedPdf && styles.actionIconBtnDisabled,
                    ]}
                    disabled={savingAnnotatedPdf}
                    onPress={saveAnnotatedPdf}
                    accessibilityLabel="Save annotated PDF"
                  >
                    <MaterialIcons
                      name={savingAnnotatedPdf ? 'hourglass-empty' : 'save-alt'}
                      size={22}
                      color={colors.onPrimary}
                    />
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.actionIconBtnDanger,
                      !selectedAnnotation && styles.actionIconBtnDisabled,
                    ]}
                    disabled={!selectedAnnotation}
                    onPress={deleteSelectedAnnotation}
                    accessibilityLabel="Delete selected annotation"
                  >
                    <MaterialIcons name="delete-outline" size={22} color={colors.onPrimary} />
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={styles.actionIconBtn}
                    onPress={() => setShowPdfOverflowMenu(true)}
                    accessibilityLabel="More options"
                  >
                    <MaterialIcons name="more-horiz" size={24} color={colors.onSurface} />
                  </TouchableOpacity>
                </>
              )}
              {documentType !== 'pdf' && officePreviewMeta.canConvertToPdf && (
                <>
                  <View style={styles.actionBarDivider} />
                  <TouchableOpacity
                    style={[
                      styles.actionIconBtnPrimary,
                      officePdfBusy && styles.actionIconBtnDisabled,
                    ]}
                    disabled={officePdfBusy}
                    onPress={convertOfficePreviewToPdf}
                    accessibilityLabel="Convert to PDF"
                  >
                    <MaterialIcons
                      name={officePdfBusy ? 'hourglass-empty' : 'picture-as-pdf'}
                      size={22}
                      color={colors.onPrimary}
                    />
                  </TouchableOpacity>
                  {officePreviewMeta.isEditable && (
                    <TouchableOpacity
                      style={styles.actionIconBtn}
                      onPress={() => setShowToolbar((currentValue) => !currentValue)}
                      accessibilityLabel={
                        showToolbar ? 'Hide editing tools' : 'Show editing tools'
                      }
                    >
                      <MaterialIcons
                        name={showToolbar ? 'expand-less' : 'expand-more'}
                        size={22}
                        color={colors.onSurface}
                      />
                    </TouchableOpacity>
                  )}
                </>
              )}
            </ScrollView>
          </View>

          {documentType === 'pdf' && (
            <View style={styles.selectionBar}>
              <Text style={styles.selectionText}>
                {activeTool === 'view'
                  ? getSelectedAnnotationDescription(selectedAnnotation)
                  : 'Switch to Select mode to tap an annotation before deleting it.'}
              </Text>
            </View>
          )}

          {documentType !== 'pdf' && officePreviewMeta.isEditable && (
            <View style={styles.selectionBar}>
              <Text style={styles.selectionText}>
                Edit text with the toolbar: fonts, sizes, colors (including custom),
                alignment, lists, and more. Tap an image to resize it. Use Convert to
                PDF to save a copy.
              </Text>
            </View>
          )}

          {documentType === 'pdf' && showToolbar && (
            <View style={styles.toolbar}>
              <View style={styles.toolRow}>
                {TOOLS.map((tool) => (
                  <TouchableOpacity
                    key={tool.id}
                    style={[
                      styles.toolIconBtn,
                      activeTool === tool.id && styles.toolIconBtnActive,
                    ]}
                    onPress={() => selectTool(tool.id)}
                    accessibilityLabel={tool.label}
                    accessibilityHint={tool.hint}
                  >
                    <MaterialIcons
                      name={tool.icon}
                      size={24}
                      color={
                        activeTool === tool.id ? colors.onPrimary : colors.onSurface
                      }
                    />
                  </TouchableOpacity>
                ))}
              </View>

              {showColorPicker && (
                <View style={styles.colorRow}>
                  {COLORS.map((color) => (
                    <TouchableOpacity
                      key={color.value}
                      style={[
                        styles.colorBtn,
                        { backgroundColor: color.value },
                        activeColor.toUpperCase() === color.value.toUpperCase() &&
                          styles.colorBtnActive,
                      ]}
                      onPress={() => selectColor(color.value)}
                    />
                  ))}
                  <TouchableOpacity
                    style={[
                      styles.colorBtn,
                      styles.colorBtnCustom,
                      !colorMatchesPreset(activeColor) && styles.colorBtnActive,
                    ]}
                    onPress={() => openCustomColorModal('pdf')}
                    accessibilityLabel="Custom color"
                    accessibilityHint="Opens sliders and hex input for any color"
                  >
                    <MaterialIcons name="palette" size={20} color={colors.primaryDim} />
                  </TouchableOpacity>
                </View>
              )}
              {activeTool === 'draw' && (
                <View style={styles.brushSliderSection}>
                  <View style={styles.brushSliderHeader}>
                    <MaterialIcons name="brush" size={18} color={colors.primary} />
                    <Text style={styles.brushSliderLabel}>Pen thickness</Text>
                    <Text style={styles.brushSliderValue}>
                      {Math.round(drawBrushPx * 2) / 2} pt
                    </Text>
                  </View>
                  <View style={styles.brushTrackRow}>
                    <TouchableOpacity
                      onPress={() =>
                        setDrawBrushPx((v) => Math.max(2, Math.min(24, v - 0.5)))
                      }
                      hitSlop={{ top: 10, bottom: 10, left: 8, right: 8 }}
                      accessibilityLabel="Decrease pen thickness"
                    >
                      <MaterialIcons
                        name="remove-circle-outline"
                        size={28}
                        color={colors.primaryDim}
                      />
                    </TouchableOpacity>
                    <BrushWidthTrack
                      min={2}
                      max={24}
                      value={drawBrushPx}
                      onChange={setDrawBrushPx}
                      accessibilityLabel="Set pen thickness"
                      trackStyle={styles.brushTrackBg}
                      fillStyle={styles.brushTrackFill}
                    />
                    <TouchableOpacity
                      onPress={() =>
                        setDrawBrushPx((v) => Math.max(2, Math.min(24, v + 0.5)))
                      }
                      hitSlop={{ top: 10, bottom: 10, left: 8, right: 8 }}
                      accessibilityLabel="Increase pen thickness"
                    >
                      <MaterialIcons
                        name="add-circle-outline"
                        size={28}
                        color={colors.primaryDim}
                      />
                    </TouchableOpacity>
                  </View>
                </View>
              )}
              {activeTool === 'marker' && (
                <View style={styles.brushSliderSection}>
                  <View style={styles.brushSliderHeader}>
                    <MaterialIcons
                      name="format-color-fill"
                      size={18}
                      color={colors.primary}
                    />
                    <Text style={styles.brushSliderLabel}>Highlighter width</Text>
                    <Text style={styles.brushSliderValue}>
                      {Math.round(markerBrushPx * 2) / 2} pt
                    </Text>
                  </View>
                  <View style={styles.brushTrackRow}>
                    <TouchableOpacity
                      onPress={() =>
                        setMarkerBrushPx((v) => Math.max(8, Math.min(36, v - 0.5)))
                      }
                      hitSlop={{ top: 10, bottom: 10, left: 8, right: 8 }}
                      accessibilityLabel="Decrease highlighter width"
                    >
                      <MaterialIcons
                        name="remove-circle-outline"
                        size={28}
                        color={colors.primaryDim}
                      />
                    </TouchableOpacity>
                    <BrushWidthTrack
                      min={8}
                      max={36}
                      value={markerBrushPx}
                      onChange={setMarkerBrushPx}
                      accessibilityLabel="Set highlighter width"
                      trackStyle={styles.brushTrackBg}
                      fillStyle={styles.brushTrackFill}
                    />
                    <TouchableOpacity
                      onPress={() =>
                        setMarkerBrushPx((v) => Math.max(8, Math.min(36, v + 0.5)))
                      }
                      hitSlop={{ top: 10, bottom: 10, left: 8, right: 8 }}
                      accessibilityLabel="Increase highlighter width"
                    >
                      <MaterialIcons
                        name="add-circle-outline"
                        size={28}
                        color={colors.primaryDim}
                      />
                    </TouchableOpacity>
                  </View>
                </View>
              )}
            </View>
          )}

          {documentType !== 'pdf' && officePreviewMeta.isEditable && showToolbar && (
            <View style={styles.toolbar}>
              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeIconScroll}
              >
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('bold')}
                  accessibilityLabel="Bold"
                >
                  <MaterialIcons name="format-bold" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('italic')}
                  accessibilityLabel="Italic"
                >
                  <MaterialIcons name="format-italic" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('underline')}
                  accessibilityLabel="Underline"
                >
                  <MaterialIcons name="format-underlined" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('strikethrough')}
                  accessibilityLabel="Strikethrough"
                >
                  <MaterialIcons name="format-strikethrough" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('subscript')}
                  accessibilityLabel="Subscript"
                >
                  <MaterialIcons name="subscript" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('superscript')}
                  accessibilityLabel="Superscript"
                >
                  <MaterialIcons name="superscript" size={22} color={colors.onSurface} />
                </TouchableOpacity>
              </ScrollView>

              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeIconScroll}
              >
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('justifyLeft')}
                  accessibilityLabel="Align left"
                >
                  <MaterialIcons name="format-align-left" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('justifyCenter')}
                  accessibilityLabel="Align center"
                >
                  <MaterialIcons name="format-align-center" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('justifyRight')}
                  accessibilityLabel="Align right"
                >
                  <MaterialIcons name="format-align-right" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('justifyFull')}
                  accessibilityLabel="Justify"
                >
                  <MaterialIcons name="format-align-justify" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('insertUnorderedList')}
                  accessibilityLabel="Bullet list"
                >
                  <MaterialIcons name="format-list-bulleted" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('insertOrderedList')}
                  accessibilityLabel="Numbered list"
                >
                  <MaterialIcons name="format-list-numbered" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('indent')}
                  accessibilityLabel="Indent"
                >
                  <MaterialIcons name="format-indent-increase" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('outdent')}
                  accessibilityLabel="Outdent"
                >
                  <MaterialIcons name="format-indent-decrease" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('removeFormat')}
                  accessibilityLabel="Clear formatting"
                >
                  <MaterialIcons name="format-clear" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() =>
                    applyOfficeStyleCommand('hiliteColor', OFFICE_HIGHLIGHT_COLOR)
                  }
                  accessibilityLabel="Highlight background"
                >
                  <MaterialIcons name="highlight" size={22} color={colors.onSurface} />
                </TouchableOpacity>
              </ScrollView>

              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeIconScroll}
              >
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={insertImageIntoOfficePreview}
                  accessibilityLabel="Insert image"
                >
                  <MaterialIcons name="image" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('scaleImage', 0.85)}
                  accessibilityLabel="Shrink selected image"
                >
                  <MaterialIcons name="zoom-out" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('scaleImage', 1.15)}
                  accessibilityLabel="Enlarge selected image"
                >
                  <MaterialIcons name="zoom-in" size={22} color={colors.onSurface} />
                </TouchableOpacity>
              </ScrollView>

              <View style={styles.officeOptionRow}>
                {OFFICE_FONT_FAMILY_OPTIONS.map((fontOption) => (
                  <TouchableOpacity
                    key={fontOption.value}
                    style={styles.officeOptionBtn}
                    onPress={() =>
                      applyOfficeStyleCommand('setFontFamily', fontOption.value)
                    }
                  >
                    <Text style={styles.officeOptionText}>{fontOption.label}</Text>
                  </TouchableOpacity>
                ))}
              </View>

              <View style={styles.officeOptionRow}>
                {OFFICE_FONT_SIZE_OPTIONS.map((fontSizeOption) => (
                  <TouchableOpacity
                    key={fontSizeOption.label}
                    style={styles.officeOptionBtn}
                    onPress={() =>
                      applyOfficeStyleCommand(
                        fontSizeOption.action,
                        fontSizeOption.value
                      )
                    }
                  >
                    <Text style={styles.officeOptionText}>{fontSizeOption.label}</Text>
                  </TouchableOpacity>
                ))}
              </View>

              <View style={styles.colorRow}>
                {COLORS.map((color) => (
                  <TouchableOpacity
                    key={color.value}
                    style={[
                      styles.colorBtn,
                      { backgroundColor: color.value },
                      activeColor.toUpperCase() === color.value.toUpperCase() &&
                        styles.colorBtnActive,
                    ]}
                    onPress={() => applyOfficeTextColor(color.value)}
                  />
                ))}
                <TouchableOpacity
                  style={[
                    styles.colorBtn,
                    styles.colorBtnCustom,
                    !colorMatchesPreset(activeColor) && styles.colorBtnActive,
                  ]}
                  onPress={() => openCustomColorModal('office')}
                  accessibilityLabel="Custom text color"
                >
                  <MaterialIcons name="palette" size={20} color={colors.primaryDim} />
                </TouchableOpacity>
              </View>
            </View>
          )}

          {documentType === 'pdf' && viewerHtml ? (
            <WebView
              key={`${documentKey || documentUri}-${viewerInstanceId}`}
              ref={webViewRef}
              source={{ html: viewerHtml }}
              style={styles.webView}
              originWhitelist={['*']}
              javaScriptEnabled={true}
              domStorageEnabled={true}
              onMessage={handleWebViewMessage}
              startInLoadingState={true}
              renderLoading={() => (
                <ActivityIndicator
                  style={styles.loadingCenter}
                  size="large"
                  color={colors.primary}
                />
              )}
              scrollEnabled={true}
              scalesPageToFit={false}
              allowFileAccess={true}
              mixedContentMode="always"
            />
          ) : (
            renderNonPdfViewer()
          )}
        </View>
      )}

      <Modal
        visible={textNoteModal !== null}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setTextNoteModal(null)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>Add Note</Text>
            <TextInput
              style={styles.modalInput}
              placeholder="Type your note here..."
              placeholderTextColor={withAlpha(colors.onSurface, 0.45)}
              value={noteText}
              onChangeText={setNoteText}
              multiline
              autoFocus
            />
            <View style={styles.modalButtons}>
              <TouchableOpacity
                style={styles.modalBtnCancel}
                onPress={() => setTextNoteModal(null)}
              >
                <Text style={styles.modalBtnCancelText}>Cancel</Text>
              </TouchableOpacity>
              <TouchableOpacity style={styles.modalBtnSubmit} onPress={submitTextNote}>
                <Text style={styles.modalBtnSubmitText}>Add Note</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>

      <Modal
        visible={showConvertModal}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setShowConvertModal(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>Convert PDF</Text>
            <Text style={styles.pageEditorHint}>
              Converts the current edited PDF copy into another editable document
              format. Inserted or removed pages are included.
            </Text>
            <Text style={styles.convertHint}>
              Word, PowerPoint, and HTML now preserve each PDF page as a full
              page image. Text and Markdown stay text-based for easier editing.
            </Text>
            {PDF_CONVERSION_FORMATS.map((formatConfig) => (
              <TouchableOpacity
                key={formatConfig.id}
                style={styles.pageActionBtnPrimary}
                onPress={() => startPdfConversion(formatConfig.id)}
              >
                <Text style={styles.pageActionBtnPrimaryText}>
                  {formatConfig.label}
                </Text>
              </TouchableOpacity>
            ))}
            <View style={styles.modalButtons}>
              <TouchableOpacity
                style={styles.modalBtnCancel}
                onPress={() => setShowConvertModal(false)}
              >
                <Text style={styles.modalBtnCancelText}>Close</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>

      <Modal
        visible={!!conversionTask}
        transparent={true}
        animationType="fade"
        onRequestClose={() => {}}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>Converting Document</Text>
            <Text style={styles.pageEditorHint}>
              {conversionTask?.formatLabel || 'Preparing conversion...'}
            </Text>
            <ActivityIndicator
              size="large"
              color={colors.primaryDim}
              style={styles.convertLoader}
            />
            <Text style={styles.convertStatusText}>
              {conversionStatus || 'Preparing conversion...'}
            </Text>
            {conversionTask ? (
              <View style={styles.hiddenConverterWebView}>
                <WebView
                  key={conversionTask.id}
                  source={{ html: conversionTask.html }}
                  onMessage={handleConverterMessage}
                  javaScriptEnabled={true}
                  domStorageEnabled={true}
                  originWhitelist={['*']}
                  mixedContentMode="always"
                />
              </View>
            ) : null}
          </View>
        </View>
      </Modal>

      <Modal
        visible={showPageEditorModal}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setShowPageEditorModal(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>Edit PDF Pages</Text>
            <Text style={styles.pageEditorHint}>
              Tap a page to select it. Long-press any page to open this panel
              for that page instantly.
            </Text>
            <Text style={styles.pageEditorMeta}>
              Current pages: {totalPages}
            </Text>
            <View style={styles.pageSelectionCard}>
              <Text style={styles.pageSelectionLabel}>Selected page</Text>
              <Text style={styles.pageSelectionValue}>
                {Number.parseInt(pageNumberInput, 10) || 1}
              </Text>
              <Text style={styles.pageSelectionHint}>
                Remove deletes this page. Blank and image add a new page after it.
              </Text>
            </View>
            <TouchableOpacity
              style={[
                styles.pageActionBtn,
                pageEditorBusy && styles.actionBtnDisabled,
              ]}
              disabled={pageEditorBusy}
              onPress={removePdfPage}
            >
              <Text style={styles.pageActionBtnText}>Remove this page</Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={[
                styles.pageActionBtnPrimary,
                pageEditorBusy && styles.actionBtnDisabled,
              ]}
              disabled={pageEditorBusy}
              onPress={insertBlankPdfPage}
            >
              <Text style={styles.pageActionBtnPrimaryText}>
                Add blank page after
              </Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={[
                styles.pageActionBtnPrimary,
                pageEditorBusy && styles.actionBtnDisabled,
              ]}
              disabled={pageEditorBusy}
              onPress={insertImagePdfPage}
            >
              <Text style={styles.pageActionBtnPrimaryText}>
                Add image page after
              </Text>
            </TouchableOpacity>
            <View style={styles.modalButtons}>
              <TouchableOpacity
                style={styles.modalBtnCancel}
                onPress={() => setShowPageEditorModal(false)}
              >
                <Text style={styles.modalBtnCancelText}>
                  {pageEditorBusy ? 'Working...' : 'Close'}
                </Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>

      <Modal
        visible={showPdfOverflowMenu}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setShowPdfOverflowMenu(false)}
      >
        <View style={styles.overflowMenuOverlay}>
          <Pressable
            style={StyleSheet.absoluteFill}
            onPress={() => setShowPdfOverflowMenu(false)}
          />
          <View style={styles.overflowMenuCard}>
            <Text style={styles.overflowMenuTitle}>More options</Text>
            <TouchableOpacity
              style={styles.overflowMenuRow}
              onPress={() => {
                setShowToolbar((v) => !v);
                setShowPdfOverflowMenu(false);
              }}
            >
              <MaterialIcons
                name={showToolbar ? 'visibility-off' : 'visibility'}
                size={22}
                color={colors.onSurface}
              />
              <Text style={styles.overflowMenuRowText}>
                {showToolbar ? 'Hide annotation tools' : 'Show annotation tools'}
              </Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={styles.overflowMenuRow}
              onPress={() => {
                setShowPdfOverflowMenu(false);
                clearAll();
              }}
            >
              <MaterialIcons name="delete-sweep" size={22} color={ui.danger} />
              <Text style={[styles.overflowMenuRowText, styles.overflowMenuRowDanger]}>
                Clear all annotations
              </Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={styles.overflowMenuClose}
              onPress={() => setShowPdfOverflowMenu(false)}
            >
              <Text style={styles.modalBtnCancelText}>Cancel</Text>
            </TouchableOpacity>
          </View>
        </View>
      </Modal>

      <Modal
        visible={showRenameModal}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setShowRenameModal(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>Document name</Text>
            <Text style={styles.renameHint}>
              Used when saving and sharing. If you omit the extension, .{documentType || 'pdf'} is added.
            </Text>
            <TextInput
              style={styles.renameInput}
              value={renameDraft}
              onChangeText={setRenameDraft}
              placeholder="File name"
              placeholderTextColor={withAlpha(colors.onSurface, 0.45)}
              autoFocus
              selectTextOnFocus
            />
            <View style={styles.modalButtons}>
              <TouchableOpacity
                style={styles.modalBtnCancel}
                onPress={() => setShowRenameModal(false)}
              >
                <Text style={styles.modalBtnCancelText}>Cancel</Text>
              </TouchableOpacity>
              <TouchableOpacity
                style={styles.modalBtnSubmit}
                onPress={applyDocumentRename}
              >
                <Text style={styles.modalBtnSubmitText}>Save</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>

      <CustomColorModal
        visible={customColorModalVisible}
        seedHex={activeColor}
        title={customColorModalTitle}
        onApply={handleCustomColorApply}
        onClose={() => setCustomColorModalVisible(false)}
      />
    </View>
  );
}

function TextFileViewer({ uri }) {
  const [content, setContent] = useState('');
  const [loadingText, setLoadingText] = useState(true);

  React.useEffect(() => {
    (async () => {
      try {
        const file = new File(uri);
        const text = await file.text();
        setContent(text);
      } catch (error) {
        setContent('Error reading file: ' + error.message);
      }
      setLoadingText(false);
    })();
  }, [uri]);

  if (loadingText) {
    return <ActivityIndicator size="large" color={colors.primary} />;
  }

  return (
    <ScrollView style={styles.textViewerScroll}>
      <Text style={styles.textViewerContent} selectable>
        {content}
      </Text>
    </ScrollView>
  );
}

function OfficeDocViewer({
  uri,
  name,
  extension,
  base64Data,
  onMessage,
  webViewRef,
}) {
  const [error, setError] = useState(null);
  const [isLoading, setIsLoading] = useState(true);

  React.useEffect(() => {
    setError(null);
    setIsLoading(false);
  }, [base64Data, extension, name, uri]);

  if (isLoading) {
    return (
      <View style={styles.landing}>
        <ActivityIndicator size="large" color={colors.primary} />
        <Text style={styles.loadingText}>Loading document...</Text>
      </View>
    );
  }

  if (error) {
    return (
      <View style={styles.unsupported}>
        <Text style={styles.unsupportedText}>{error}</Text>
      </View>
    );
  }

  const officeHtml =
    base64Data && OFFICE_PREVIEW_TYPES.includes(extension)
      ? getOfficePreviewHtml(base64Data, {
          extension,
          fileName: name,
        })
      : `
          <!DOCTYPE html>
          <html>
          <head>
            <meta name="viewport" content="width=device-width, initial-scale=1.0" />
            <style>
              body {
                font-family: -apple-system, BlinkMacSystemFont, sans-serif;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                min-height: 100vh;
                margin: 0;
                background: #f5f5f5;
                padding: 20px;
                text-align: center;
              }
              .icon { font-size: 64px; margin-bottom: 16px; }
              h2 { color: #333; margin-bottom: 8px; font-size: 18px; }
              p { color: #666; font-size: 14px; line-height: 1.5; max-width: 320px; }
              .filename {
                color: #6200ee;
                font-weight: 600;
                word-break: break-all;
                margin: 12px 0;
                background: #ede7f6;
                padding: 8px 16px;
                border-radius: 8px;
                font-size: 13px;
              }
              .tip {
                margin-top: 20px;
                padding: 12px 16px;
                background: #e3f2fd;
                border-radius: 8px;
                color: #1565c0;
                font-size: 12px;
              }
            </style>
          </head>
          <body>
            <div class="icon">&#128196;</div>
            <h2>${name}</h2>
            <p class="filename">${name}</p>
            <p>This document opens best in its dedicated editor.</p>
            <p>Use the <strong>Share</strong> button above to open it in Microsoft Office, PowerPoint, Excel, Google Docs, or another compatible app.</p>
            <div class="tip">
              In-app preview is available for DOC, DOCX, PPTX, XLSX, and XLS files.
            </div>
          </body>
          </html>
        `;

  return (
    <WebView
      ref={webViewRef}
      source={{ html: officeHtml }}
      style={styles.webView}
      scrollEnabled={true}
      onMessage={onMessage}
      originWhitelist={['*']}
      javaScriptEnabled={true}
      domStorageEnabled={true}
      mixedContentMode="always"
    />
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: ui.shell,
  },
  header: {
    paddingBottom: 12,
    paddingHorizontal: spacing.md,
    backgroundColor: ui.shellElevated,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: colors.outlineVariant,
  },
  headerTitleBlock: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    marginRight: spacing.sm,
  },
  headerTitle: {
    color: colors.onSurface,
    fontSize: 17,
    fontWeight: '700',
    flex: 1,
    letterSpacing: -0.2,
  },
  headerEditIcon: {
    marginLeft: 6,
  },
  pageCountPill: {
    flexDirection: 'row',
    alignItems: 'baseline',
    backgroundColor: colors.surfaceContainerLow,
    paddingHorizontal: 10,
    paddingVertical: 6,
    borderRadius: radius.pill,
    borderWidth: 1,
    borderColor: withAlpha(colors.primary, 0.22),
  },
  pageCountLabel: {
    color: colors.primary,
    fontSize: 14,
    fontWeight: '800',
  },
  pageCountSuffix: {
    color: ui.textSoft,
    fontSize: 11,
    fontWeight: '600',
    marginLeft: 4,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
  },
  landing: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 32,
  },
  landingIcon: {
    marginBottom: 16,
  },
  landingTitle: {
    color: colors.onSurface,
    fontSize: 24,
    fontWeight: '700',
    marginBottom: 12,
    textAlign: 'center',
  },
  importButtonRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 10,
  },
  landingSubtitle: {
    color: ui.textMuted,
    fontSize: 14,
    textAlign: 'center',
    lineHeight: 22,
    marginBottom: 32,
    paddingHorizontal: spacing.sm,
  },
  importButton: {
    backgroundColor: colors.primary,
    paddingHorizontal: spacing.lg,
    paddingVertical: 16,
    borderRadius: radius.md,
    elevation: 4,
    shadowColor: colors.primary,
    shadowOffset: { width: 0, height: 4 },
    shadowOpacity: 0.3,
    shadowRadius: 8,
  },
  importButtonText: {
    color: colors.onPrimary,
    fontSize: 18,
    fontWeight: '600',
  },
  supportedFormats: {
    color: ui.textSoft,
    fontSize: 11,
    marginTop: 20,
    textAlign: 'center',
  },
  loadingText: {
    color: ui.textMuted,
    marginTop: 16,
    fontSize: 14,
  },
  documentArea: {
    flex: 1,
  },
  actionBar: {
    backgroundColor: ui.shellElevated,
    paddingVertical: 10,
    paddingHorizontal: spacing.sm,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: colors.outlineVariant,
  },
  actionBarScroll: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    paddingRight: spacing.md,
  },
  actionIconBtn: {
    width: 44,
    height: 44,
    borderRadius: radius.sm,
    backgroundColor: colors.surfaceContainerLowest,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.65),
  },
  actionIconBtnPrimary: {
    width: 44,
    height: 44,
    borderRadius: radius.sm,
    backgroundColor: colors.primaryDim,
    alignItems: 'center',
    justifyContent: 'center',
  },
  actionIconBtnDanger: {
    width: 44,
    height: 44,
    borderRadius: radius.sm,
    backgroundColor: ui.danger,
    alignItems: 'center',
    justifyContent: 'center',
  },
  actionIconBtnDisabled: {
    opacity: 0.45,
  },
  actionBarDivider: {
    width: StyleSheet.hairlineWidth,
    height: 26,
    backgroundColor: withAlpha(colors.outlineVariant, 0.9),
    marginHorizontal: 4,
  },
  actionBtn: {
    backgroundColor: ui.shellSoft,
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: radius.sm,
    marginHorizontal: 4,
  },
  actionBtnPrimary: {
    backgroundColor: colors.primaryDim,
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: radius.sm,
    marginHorizontal: 4,
  },
  actionBtnText: {
    color: colors.onSurface,
    fontSize: 13,
    fontWeight: '500',
  },
  actionBtnPrimaryText: {
    color: colors.onPrimary,
    fontSize: 13,
    fontWeight: '600',
  },
  actionBtnDanger: {
    backgroundColor: '#b71c1c',
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: radius.sm,
    marginHorizontal: 4,
  },
  actionBtnDangerText: {
    color: colors.onPrimary,
    fontSize: 13,
    fontWeight: '500',
  },
  actionBtnDisabled: {
    opacity: 0.5,
  },
  selectionBar: {
    backgroundColor: ui.shellSoft,
    paddingHorizontal: 14,
    paddingVertical: 10,
    borderBottomWidth: 1,
    borderBottomColor: colors.outlineVariant,
  },
  selectionText: {
    color: withAlpha(colors.onSurface, 0.72),
    fontSize: 12,
    lineHeight: 18,
  },
  toolbar: {
    backgroundColor: ui.shell,
    paddingVertical: 10,
    paddingHorizontal: spacing.sm,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: colors.outlineVariant,
  },
  toolRow: {
    flexDirection: 'row',
    justifyContent: 'center',
    alignItems: 'center',
    gap: 10,
    flexWrap: 'wrap',
  },
  toolIconBtn: {
    width: 48,
    height: 48,
    borderRadius: radius.md,
    backgroundColor: ui.shellElevated,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.55),
  },
  toolIconBtnActive: {
    backgroundColor: colors.primary,
    borderColor: colors.primaryDim,
  },
  toolBtn: {
    paddingHorizontal: spacing.sm,
    paddingVertical: 8,
    borderRadius: radius.sm,
    backgroundColor: ui.shellElevated,
  },
  toolBtnActive: {
    backgroundColor: colors.primary,
  },
  toolBtnText: {
    color: ui.textMuted,
    fontSize: 13,
    fontWeight: '500',
  },
  toolBtnTextActive: {
    color: colors.onPrimary,
  },
  colorRow: {
    flexDirection: 'row',
    justifyContent: 'center',
    marginTop: 8,
    flexWrap: 'wrap',
    gap: 8,
  },
  officeOptionRow: {
    flexDirection: 'row',
    justifyContent: 'center',
    flexWrap: 'wrap',
    marginTop: 8,
    gap: 8,
  },
  officeOptionBtn: {
    backgroundColor: ui.shellElevated,
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: radius.sm,
  },
  officeOptionText: {
    color: withAlpha(colors.onSurface, 0.88),
    fontSize: 13,
    fontWeight: '500',
  },
  colorBtn: {
    width: 28,
    height: 28,
    borderRadius: 14,
    borderWidth: 2,
    borderColor: 'transparent',
  },
  colorBtnActive: {
    borderColor: '#fff',
    borderWidth: 3,
  },
  brushSliderSection: {
    marginTop: 10,
    paddingTop: 10,
    borderTopWidth: StyleSheet.hairlineWidth,
    borderTopColor: withAlpha(colors.outlineVariant, 0.65),
  },
  brushSliderHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    marginBottom: 8,
    gap: 8,
  },
  brushSliderLabel: {
    flex: 1,
    fontSize: 12,
    fontWeight: '700',
    color: withAlpha(colors.onSurface, 0.78),
    letterSpacing: 0.3,
  },
  brushSliderValue: {
    fontSize: 12,
    fontWeight: '700',
    color: colors.primary,
    minWidth: 44,
    textAlign: 'right',
  },
  brushTrackRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 10,
  },
  brushTrackBg: {
    flex: 1,
    height: 10,
    borderRadius: 5,
    backgroundColor: withAlpha(colors.outlineVariant, 0.45),
    overflow: 'hidden',
    justifyContent: 'center',
  },
  brushTrackFill: {
    height: '100%',
    borderRadius: 5,
    backgroundColor: colors.primary,
  },
  webView: {
    flex: 1,
    backgroundColor: colors.surfaceContainerLow,
  },
  loadingCenter: {
    position: 'absolute',
    top: '50%',
    left: '50%',
    marginLeft: -20,
    marginTop: -20,
  },
  nonPdfContainer: {
    flex: 1,
    backgroundColor: colors.surfaceContainerLowest,
  },
  unsupported: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 32,
  },
  unsupportedText: {
    color: colors.onSurface,
    fontSize: 16,
    fontWeight: '600',
    marginBottom: 8,
  },
  unsupportedHint: {
    color: ui.textSoft,
    fontSize: 13,
    textAlign: 'center',
  },
  textViewerScroll: {
    flex: 1,
    padding: 16,
  },
  textViewerContent: {
    fontSize: 14,
    lineHeight: 22,
    color: colors.onSurface,
    fontFamily: Platform.OS === 'ios' ? 'Menlo' : 'monospace',
  },
  modalOverlay: {
    flex: 1,
    backgroundColor: withAlpha(colors.onSurface, 0.42),
    justifyContent: 'center',
    alignItems: 'center',
  },
  modalContent: {
    backgroundColor: colors.surfaceContainerLowest,
    borderRadius: radius.md,
    padding: spacing.md,
    width: '85%',
    maxWidth: 400,
  },
  customColorModalContent: {
    maxWidth: 440,
    maxHeight: '92%',
    paddingTop: spacing.sm,
    paddingBottom: spacing.sm,
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 12 },
        shadowOpacity: 0.18,
        shadowRadius: 24,
      },
      android: {
        elevation: 12,
      },
    }),
  },
  customColorHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: 4,
    marginBottom: 4,
  },
  customColorHeaderTitle: {
    flex: 1,
    fontSize: 18,
    fontWeight: '700',
    color: colors.onSurface,
    letterSpacing: 0.2,
    paddingRight: 8,
  },
  customColorScroll: {
    maxHeight: 520,
  },
  customColorHint: {
    fontSize: 12,
    lineHeight: 18,
    color: withAlpha(colors.onSurface, 0.7),
    marginBottom: 14,
  },
  customColorPreviewRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 14,
    marginBottom: 18,
    paddingVertical: 12,
    paddingHorizontal: 14,
    borderRadius: radius.sm,
    backgroundColor: withAlpha(colors.primary, 0.06),
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.4),
  },
  customColorPreviewCircle: {
    width: 56,
    height: 56,
    borderRadius: 28,
    borderWidth: 3,
    borderColor: withAlpha(colors.onSurface, 0.12),
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 2 },
        shadowOpacity: 0.12,
        shadowRadius: 4,
      },
      android: {
        elevation: 3,
      },
    }),
  },
  customColorPreviewMeta: {
    flex: 1,
  },
  customColorPreviewLabel: {
    fontSize: 11,
    fontWeight: '700',
    color: withAlpha(colors.onSurface, 0.55),
    textTransform: 'uppercase',
    letterSpacing: 0.8,
    marginBottom: 4,
  },
  customColorPreviewHex: {
    fontSize: 17,
    fontWeight: '700',
    fontFamily: Platform.OS === 'ios' ? 'Menlo' : 'monospace',
    color: colors.primaryDim,
  },
  customColorSectionLabel: {
    fontSize: 11,
    fontWeight: '700',
    color: withAlpha(colors.onSurface, 0.62),
    textTransform: 'uppercase',
    letterSpacing: 0.9,
    marginBottom: 8,
    marginTop: 4,
  },
  spectrumPadOuter: {
    borderRadius: 14,
    overflow: 'hidden',
    marginBottom: 6,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.5),
  },
  spectrumPad: {
    width: '100%',
    height: 200,
    borderRadius: 14,
    overflow: 'hidden',
  },
  svThumb: {
    position: 'absolute',
    width: 22,
    height: 22,
    borderRadius: 11,
    borderWidth: 3,
    borderColor: '#ffffff',
    backgroundColor: 'rgba(255,255,255,0.2)',
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 1 },
        shadowOpacity: 0.45,
        shadowRadius: 3,
      },
      android: {
        elevation: 4,
      },
    }),
  },
  hueStripOuter: {
    height: 34,
    borderRadius: 10,
    overflow: 'hidden',
    marginBottom: 8,
    position: 'relative',
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.45),
  },
  hueStripGradient: {
    ...StyleSheet.absoluteFillObject,
  },
  hueStripOverlay: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'transparent',
  },
  hueThumb: {
    position: 'absolute',
    width: 5,
    top: 0,
    bottom: 0,
    borderRadius: 3,
    borderWidth: 2,
    borderColor: '#ffffff',
    backgroundColor: 'rgba(255,255,255,0.25)',
    ...Platform.select({
      ios: {
        shadowColor: '#000',
        shadowOffset: { width: 0, height: 1 },
        shadowOpacity: 0.4,
        shadowRadius: 2,
      },
      android: {
        elevation: 3,
      },
    }),
  },
  customColorGradientHint: {
    fontSize: 11,
    lineHeight: 16,
    color: withAlpha(colors.onSurface, 0.58),
    marginBottom: 10,
  },
  gradientPresetsScroll: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: 10,
    paddingVertical: 4,
    marginBottom: 14,
    paddingRight: 8,
  },
  gradientPresetChip: {
    width: 92,
    borderRadius: 10,
    overflow: 'hidden',
    backgroundColor: ui.shellElevated,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.4),
  },
  gradientPresetFill: {
    height: 36,
    width: '100%',
  },
  gradientPresetLabel: {
    fontSize: 10,
    fontWeight: '600',
    color: withAlpha(colors.onSurface, 0.75),
    textAlign: 'center',
    paddingVertical: 6,
    paddingHorizontal: 4,
    backgroundColor: colors.surfaceContainerLowest,
  },
  customColorAdvancedToggle: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    paddingVertical: 10,
    marginBottom: 4,
  },
  customColorAdvancedToggleText: {
    flex: 1,
    fontSize: 13,
    fontWeight: '600',
    color: withAlpha(colors.onSurface, 0.82),
  },
  customColorHexInput: {
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.45),
    borderRadius: radius.sm,
    paddingHorizontal: 12,
    paddingVertical: 10,
    fontSize: 16,
    fontFamily: Platform.OS === 'ios' ? 'Menlo' : 'monospace',
    color: colors.onSurface,
    marginBottom: 8,
    backgroundColor: withAlpha(colors.surfaceContainerLow, 0.85),
  },
  customColorRgbRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 10,
    marginBottom: 12,
  },
  customColorChannelLabel: {
    width: 44,
    fontSize: 12,
    fontWeight: '600',
    color: withAlpha(colors.onSurface, 0.78),
  },
  customColorChannelValue: {
    width: 36,
    fontSize: 13,
    fontWeight: '700',
    color: colors.primary,
    textAlign: 'right',
  },
  customColorTrack: {
    flex: 1,
    height: 10,
    borderRadius: 5,
    backgroundColor: withAlpha(colors.outlineVariant, 0.45),
    overflow: 'hidden',
    justifyContent: 'center',
  },
  colorBtnCustom: {
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: ui.shellElevated,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.55),
  },
  officeIconScroll: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    paddingVertical: 4,
    marginBottom: 4,
  },
  officeIconBtn: {
    width: 44,
    height: 44,
    borderRadius: radius.sm,
    backgroundColor: ui.shellElevated,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.45),
  },
  modalTitle: {
    fontSize: 18,
    fontWeight: '700',
    color: colors.onSurface,
    marginBottom: 16,
  },
  pageEditorHint: {
    color: withAlpha(colors.onSurface, 0.78),
    fontSize: 13,
    lineHeight: 19,
    marginBottom: 8,
  },
  convertHint: {
    color: withAlpha(colors.onSurface, 0.6),
    fontSize: 12,
    lineHeight: 18,
    marginBottom: 16,
  },
  pageEditorMeta: {
    color: withAlpha(colors.onSurface, 0.6),
    fontSize: 12,
    marginBottom: 14,
  },
  pageSelectionCard: {
    borderRadius: radius.md,
    paddingHorizontal: 14,
    paddingVertical: 13,
    backgroundColor: colors.surfaceContainerLow,
    borderWidth: 1,
    borderColor: withAlpha(colors.primary, 0.18),
    marginBottom: 16,
  },
  pageSelectionLabel: {
    color: withAlpha(colors.primary, 0.82),
    fontSize: 12,
    fontWeight: '600',
    marginBottom: 4,
    textTransform: 'uppercase',
    letterSpacing: 0.5,
  },
  pageSelectionValue: {
    color: colors.onSurface,
    fontSize: 28,
    fontWeight: '700',
    marginBottom: 4,
  },
  pageSelectionHint: {
    color: withAlpha(colors.onSurface, 0.72),
    fontSize: 12,
    lineHeight: 18,
  },
  convertLoader: {
    marginVertical: 12,
  },
  convertStatusText: {
    color: withAlpha(colors.onSurface, 0.74),
    fontSize: 13,
    lineHeight: 19,
    textAlign: 'center',
  },
  hiddenConverterWebView: {
    width: 1,
    height: 1,
    opacity: 0,
    overflow: 'hidden',
  },
  pageActionBtn: {
    backgroundColor: ui.danger,
    paddingHorizontal: spacing.sm,
    paddingVertical: 12,
    borderRadius: radius.sm,
    marginBottom: 10,
  },
  pageActionBtnPrimary: {
    backgroundColor: colors.primaryDim,
    paddingHorizontal: spacing.sm,
    paddingVertical: 12,
    borderRadius: radius.sm,
    marginBottom: 10,
  },
  pageActionBtnText: {
    color: colors.onPrimary,
    fontSize: 14,
    fontWeight: '600',
    textAlign: 'center',
  },
  pageActionBtnPrimaryText: {
    color: colors.onPrimary,
    fontSize: 14,
    fontWeight: '600',
    textAlign: 'center',
  },
  modalInput: {
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.45),
    borderRadius: radius.sm,
    padding: 12,
    fontSize: 14,
    minHeight: 80,
    textAlignVertical: 'top',
    color: colors.onSurface,
    marginBottom: 16,
  },
  modalButtons: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gap: 12,
  },
  modalBtnCancel: {
    paddingHorizontal: 20,
    paddingVertical: 10,
    borderRadius: radius.sm,
  },
  modalBtnCancelText: {
    color: ui.textSoft,
    fontSize: 14,
    fontWeight: '500',
  },
  modalBtnSubmit: {
    backgroundColor: colors.primary,
    paddingHorizontal: 20,
    paddingVertical: 10,
    borderRadius: radius.sm,
  },
  modalBtnSubmitText: {
    color: colors.onPrimary,
    fontSize: 14,
    fontWeight: '600',
  },
  overflowMenuOverlay: {
    flex: 1,
    backgroundColor: withAlpha(colors.onSurface, 0.42),
    justifyContent: 'center',
    alignItems: 'center',
    padding: spacing.lg,
  },
  overflowMenuCard: {
    backgroundColor: colors.surfaceContainerLowest,
    borderRadius: radius.md,
    paddingVertical: spacing.sm,
    width: '100%',
    maxWidth: 340,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.5),
  },
  overflowMenuTitle: {
    fontSize: 13,
    fontWeight: '800',
    color: withAlpha(colors.primary, 0.95),
    letterSpacing: 0.8,
    textTransform: 'uppercase',
    paddingHorizontal: spacing.md,
    paddingBottom: spacing.xs,
    paddingTop: spacing.xs,
  },
  overflowMenuRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12,
    paddingVertical: 14,
    paddingHorizontal: spacing.md,
  },
  overflowMenuRowText: {
    flex: 1,
    fontSize: 15,
    fontWeight: '500',
    color: colors.onSurface,
  },
  overflowMenuRowDanger: {
    color: ui.danger,
    fontWeight: '600',
  },
  overflowMenuClose: {
    alignItems: 'center',
    paddingVertical: 12,
    marginTop: 4,
  },
  renameHint: {
    color: withAlpha(colors.onSurface, 0.65),
    fontSize: 12,
    lineHeight: 18,
    marginBottom: 12,
  },
  renameInput: {
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.45),
    borderRadius: radius.sm,
    padding: 12,
    fontSize: 15,
    minHeight: 48,
    color: colors.onSurface,
    marginBottom: 16,
  },
});

