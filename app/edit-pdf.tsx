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
import { Paths } from 'expo-file-system';
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
import {
  DOCUMENT_PICKER_TYPES,
  OFFICE_PREVIEW_TYPES,
} from '@/src/constants/documentPicker';
import * as PdfConversionFormats from '@/src/constants/pdfConversionFormats';
import { consumePendingEditDocument } from '@/src/navigation/pendingEditDocument';
import { recordSavedFile } from '@/src/db/savedFileHistory';

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
/** Full list for font picker modal (Word-style dropdown). */
const OFFICE_FONT_DROPDOWN_OPTIONS = [
  { label: 'Arial', value: 'Arial' },
  { label: 'Helvetica', value: 'Helvetica' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Verdana', value: 'Verdana' },
  { label: 'Tahoma', value: 'Tahoma' },
  { label: 'Trebuchet MS', value: 'Trebuchet MS' },
  { label: 'Courier New', value: 'Courier New' },
  { label: 'Consolas', value: 'Consolas' },
  { label: 'Palatino', value: 'Palatino Linotype' },
  { label: 'Garamond', value: 'Garamond' },
  { label: 'Comic Sans MS', value: 'Comic Sans MS' },
  { label: 'Impact', value: 'Impact' },
];

const OFFICE_FONT_WEIGHT_OPTIONS = [
  { label: 'Light', value: '300' },
  { label: 'Regular', value: '400' },
  { label: 'Medium', value: '500' },
  { label: 'Semibold', value: '600' },
  { label: 'Bold', value: '700' },
  { label: 'Heavy', value: '800' },
];

const OFFICE_LINE_HEIGHT_OPTIONS = [
  { label: '1', value: 1 },
  { label: '1.15', value: 1.15 },
  { label: '1.35', value: 1.35 },
  { label: '1.5', value: 1.5 },
  { label: '1.75', value: 1.75 },
  { label: '2', value: 2 },
];

const OFFICE_HEADING_OPTIONS = [
  { label: 'Normal', value: 'p' },
  { label: 'Quote', value: 'blockquote' },
  { label: 'Code', value: 'pre' },
  { label: 'Heading 1', value: 'h1' },
  { label: 'Heading 2', value: 'h2' },
  { label: 'Heading 3', value: 'h3' },
  { label: 'Heading 4', value: 'h4' },
];

const PRESENTATION_THEME_OPTIONS = [
  { label: 'Iris', value: 'default' },
  { label: 'Ocean', value: 'ocean' },
  { label: 'Sunset', value: 'sunset' },
  { label: 'Forest', value: 'forest' },
  { label: 'Midnight', value: 'midnight' },
  { label: 'Paper', value: 'paper' },
];

const PAGE_IMAGE_EXPORT_FORMATS = [
  { label: 'PNG', subtitle: 'Lossless, larger file', mimeType: 'image/png', extension: 'png' },
  { label: 'JPEG', subtitle: 'Compressed photos', mimeType: 'image/jpeg', extension: 'jpg' },
  { label: 'WebP', subtitle: 'Smaller modern format', mimeType: 'image/webp', extension: 'webp' },
];

function getTextMimeTypeForExtension(ext) {
  const e = String(ext || '').toLowerCase();
  const map = {
    txt: 'text/plain',
    md: 'text/markdown',
    html: 'text/html',
    htm: 'text/html',
    csv: 'text/csv',
    json: 'application/json',
    xml: 'application/xml',
    yaml: 'text/yaml',
    yml: 'text/yaml',
    css: 'text/css',
    js: 'text/javascript',
    ts: 'text/typescript',
    tsx: 'text/typescript',
    jsx: 'text/javascript',
    mdx: 'text/markdown',
    scss: 'text/x-scss',
    less: 'text/less',
    log: 'text/plain',
  };
  return map[e] || 'text/plain';
}

const TEXT_FILE_EXTENSIONS = [
  'txt',
  'md',
  'html',
  'htm',
  'csv',
  'json',
  'xml',
  'yaml',
  'yml',
  'css',
  'log',
  'js',
  'ts',
  'tsx',
  'jsx',
  'mdx',
  'scss',
  'less',
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

/** Move page `fromPage` to sit after page `toAfterPage` (1-based). */
function reorderPageInAnnotationState(
  rawAnnotations,
  fromPage,
  toAfterPage,
  pageCount
) {
  const norm = normalizeAnnotationState(rawAnnotations);
  const order = [];
  for (let p = 1; p <= pageCount; p += 1) {
    order.push(p);
  }
  const fromIdx = order.indexOf(fromPage);
  if (fromIdx === -1) {
    return norm;
  }
  order.splice(fromIdx, 1);
  const afterIdx = order.indexOf(toAfterPage);
  const insertAt = afterIdx === -1 ? order.length : afterIdx + 1;
  order.splice(insertAt, 0, fromPage);
  const next = {};
  order.forEach((oldPageNum, i) => {
    const key = String(i + 1);
    const src = norm[String(oldPageNum)] || norm[oldPageNum];
    next[key] = src || createEmptyPageAnnotations();
  });
  return next;
}

function shortDocLabel(name, maxLen = 22) {
  if (!name || typeof name !== 'string') {
    return 'Untitled';
  }
  const trimmed = name.trim();
  if (trimmed.length <= maxLen) {
    return trimmed;
  }
  return `${trimmed.slice(0, Math.max(1, maxLen - 1))}…`;
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

    void recordSavedFile({
      uri: targetUri,
      fileName,
      mimeType,
      directoryUri: permission.directoryUri,
      source: 'android_saf',
    });

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
  const [savingTextFile, setSavingTextFile] = useState(false);
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
  const [exportPageImageBusy, setExportPageImageBusy] = useState(false);
  const [viewerInstanceId, setViewerInstanceId] = useState(0);
  const [showRenameModal, setShowRenameModal] = useState(false);
  const [renameDraft, setRenameDraft] = useState('');
  const [customColorModalVisible, setCustomColorModalVisible] = useState(false);
  const [customColorModalTitle, setCustomColorModalTitle] =
    useState('Annotation color');
  const customColorTargetRef = useRef('pdf');
  const textFileSaveRef = useRef(() => Promise.resolve());
  const [officeFontModalVisible, setOfficeFontModalVisible] = useState(false);
  const [showPdfOverflowMenu, setShowPdfOverflowMenu] = useState(false);
  const [drawBrushPx, setDrawBrushPx] = useState(6);
  const [markerBrushPx, setMarkerBrushPx] = useState(16);

  const [secondaryUri, setSecondaryUri] = useState(null);
  const [secondaryName, setSecondaryName] = useState('');
  const [secondaryType, setSecondaryType] = useState('');
  const [secondaryKey, setSecondaryKey] = useState(null);
  const [secondaryBase64, setSecondaryBase64] = useState(null);
  const [secondaryViewerHtml, setSecondaryViewerHtml] = useState(null);
  const [secondaryAnnotationState, setSecondaryAnnotationState] = useState({});
  const [secondaryTotalPages, setSecondaryTotalPages] = useState(0);
  const [secondaryViewerInstanceId, setSecondaryViewerInstanceId] = useState(0);
  const [secondaryOfficeMeta, setSecondaryOfficeMeta] = useState({
    previewKind: '',
    canConvertToPdf: false,
    isEditable: false,
  });
  /** 1 = primary document, 2 = second document (toolbar + commands target this pane). */
  const [activePane, setActivePane] = useState(1);
  /** `row` = side-by-side, `column` = top / bottom */
  const [dualLayout, setDualLayout] = useState('row');
  const [showOpenChoiceModal, setShowOpenChoiceModal] = useState(false);
  const [showPageTransferModal, setShowPageTransferModal] = useState(false);
  const [pageTransferBusy, setPageTransferBusy] = useState(false);
  const [pageTransferSourcePane, setPageTransferSourcePane] = useState(1);
  const [pageTransferTargetPane, setPageTransferTargetPane] = useState(2);
  const [pageTransferSourcePage, setPageTransferSourcePage] = useState('1');
  const [pageTransferInsertBefore, setPageTransferInsertBefore] = useState('1');
  const [pageTransferMode, setPageTransferMode] = useState('copy');
  const [showDualPdfGuide, setShowDualPdfGuide] = useState(true);

  const insets = useSafeAreaInsets();

  const webViewRef = useRef(null);
  const webViewRef2 = useRef(null);
  /** Which pane (1 or 2) initiated image export / office PDF export (file names). */
  const exportContextPaneRef = useRef(1);
  const officeExportSlotRef = useRef(1);
  const conversionChunksRef = useRef({
    fileName: '',
    mimeType: '',
    chunks: [],
  });
  const annotationWriteQueueRef = useRef(Promise.resolve());
  const annotationDirectoryReadyRef = useRef(false);
  const pageImageExportRef = useRef({ requestId: null });
  /** `copy` = new timestamped PDF; `replace` = same base name, then open as PDF. */
  const officeSaveModeRef = useRef('copy');
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
    (
      nextBase64,
      nextAnnotations,
      nextMode = 'view',
      nextSelectedPage = 1,
      htmlOptions = {}
    ) =>
      getPdfViewerHtml(nextBase64, {
        initialAnnotations: nextAnnotations,
        initialMode: nextMode,
        initialDrawColor: activeColor,
        initialHighlightColor: activeColor + HIGHLIGHT_ALPHA_HEX,
        initialSelectedPage: nextSelectedPage,
        viewerSlot: htmlOptions.viewerSlot ?? 1,
        crossPaneDragEnabled: !!htmlOptions.crossPaneDragEnabled,
        dualLayout: htmlOptions.dualLayout === 'column' ? 'column' : 'row',
      }),
    [activeColor]
  );

  const crossPdfDropPendingRef = useRef(null);
  const [crossPdfDropBanner, setCrossPdfDropBanner] = useState(false);
  const [crossPdfDropMode, setCrossPdfDropMode] = useState('copy');

  const isDualPdfMode =
    !!secondaryUri && documentType === 'pdf' && secondaryType === 'pdf';

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

  const loadEditablePdfBase64ForPane = useCallback(
    async (pane) => {
      if (pane === 1) {
        return loadEditablePdfBase64();
      }
      if (secondaryKey) {
        const workingPdfBase64 = await loadWorkingPdfCopy(secondaryKey);
        if (workingPdfBase64) {
          return workingPdfBase64;
        }
      }
      return secondaryBase64;
    },
    [loadEditablePdfBase64, loadWorkingPdfCopy, secondaryBase64, secondaryKey]
  );

  const buildEditedPdfBase64 = useCallback(
    async (paneOverride) => {
      const pane = paneOverride ?? activePane;
      const editablePdfBase64 = await loadEditablePdfBase64ForPane(pane);

      if (!editablePdfBase64) {
        throw new Error('No editable PDF data found.');
      }

      const ann = pane === 1 ? annotationState : secondaryAnnotationState;
      return bakePdfWithAnnotations(editablePdfBase64, ann);
    },
    [activePane, annotationState, loadEditablePdfBase64ForPane, secondaryAnnotationState]
  );

  const buildEditedPdfFile = useCallback(
    async (outputDirectoryUri = Paths.cache.uri, paneOverride) => {
      const bakedPdfBase64 = await buildEditedPdfBase64(paneOverride);
      const pane = paneOverride ?? activePane;
      const exportName = pane === 1 ? documentName : secondaryName;
      const fileName = buildAnnotatedFileName(exportName);
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
    [activePane, buildEditedPdfBase64, documentName, secondaryName]
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

  const clearSecondaryDocument = useCallback(() => {
    setSecondaryUri(null);
    setSecondaryName('');
    setSecondaryType('');
    setSecondaryKey(null);
    setSecondaryBase64(null);
    setSecondaryViewerHtml(null);
    setSecondaryAnnotationState({});
    setSecondaryTotalPages(0);
    setSecondaryViewerInstanceId(0);
    setSecondaryOfficeMeta({
      previewKind: '',
      canConvertToPdf: false,
      isEditable: false,
    });
    setActivePane(1);
  }, []);

  useEffect(() => {
    if (TEXT_FILE_EXTENSIONS.includes(documentType)) {
      setOfficePreviewMeta({
        previewKind: 'text',
        canConvertToPdf: false,
        isEditable: true,
      });
    }
  }, [documentType]);

  useEffect(() => {
    if (secondaryUri && TEXT_FILE_EXTENSIONS.includes(secondaryType)) {
      setSecondaryOfficeMeta({
        previewKind: 'text',
        canConvertToPdf: false,
        isEditable: true,
      });
    }
  }, [secondaryUri, secondaryType]);

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
      const dualPdf =
        !!secondaryUri && documentType === 'pdf' && secondaryType === 'pdf';
      setViewerHtml(
        buildViewerHtml(nextBase64, nextAnnotations, nextMode, safeSelectedPage, {
          viewerSlot: 1,
          crossPaneDragEnabled: dualPdf,
          dualLayout,
        })
      );
      setViewerInstanceId((currentValue) => currentValue + 1);
    },
    [buildViewerHtml, documentType, dualLayout, secondaryType, secondaryUri]
  );

  const applySecondaryPdfEditorState = useCallback(
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
      setSecondaryBase64(nextBase64);
      setSecondaryAnnotationState(nextAnnotations);
      setActiveTool(nextMode);
      setShowColorPicker(
        nextMode === 'draw' || nextMode === 'highlight' || nextMode === 'marker'
      );
      setSecondaryTotalPages(nextPageCount);
      setPageNumberInput(String(safeSelectedPage));
      const dualPdf =
        !!secondaryUri && documentType === 'pdf' && secondaryType === 'pdf';
      setSecondaryViewerHtml(
        buildViewerHtml(nextBase64, nextAnnotations, nextMode, safeSelectedPage, {
          viewerSlot: 2,
          crossPaneDragEnabled: dualPdf,
          dualLayout,
        })
      );
      setSecondaryViewerInstanceId((currentValue) => currentValue + 1);
    },
    [buildViewerHtml, documentType, dualLayout, secondaryType, secondaryUri]
  );

  const transitionOfficeExportToPdf = useCallback(
    (pdfBase64, pdfUri, pdfName) => {
      const nextDocumentKey = `office-export-${Date.now()}`;
      setDocumentUri(pdfUri);
      setDocumentName(pdfName);
      setDocumentType('pdf');
      setDocumentKey(nextDocumentKey);
      setBase64Data(pdfBase64);
      setAnnotationState({});
      setSelectedAnnotation(null);
      setActiveTool('view');
      setActiveColor(DEFAULT_COLOR);
      setShowColorPicker(false);
      setShowToolbar(true);
      setTextNoteModal(null);
      setNoteText('');
      setTotalPages(0);
      setPageNumberInput('1');
      setViewerHtml(
        buildViewerHtml(pdfBase64, {}, 'view', 1, {
          viewerSlot: 1,
          crossPaneDragEnabled: false,
        })
      );
      setViewerInstanceId((currentValue) => currentValue + 1);
      resetOfficePreviewState();
      setShowPageEditorModal(false);
      setShowConvertModal(false);
      resetConversionState();
    },
    [buildViewerHtml, resetConversionState, resetOfficePreviewState]
  );

  const injectCommandToSlot = useCallback((slot, command) => {
    const ref = slot === 1 ? webViewRef : webViewRef2;
    if (!ref.current) {
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

    ref.current.injectJavaScript(js);
  }, []);

  const sendCommand = useCallback(
    (command) => {
      injectCommandToSlot(activePane, command);
    },
    [activePane, injectCommandToSlot]
  );

  const applyPickedAssetToSlot = useCallback(
    async (slot, file, options = {}) => {
      const { replaceAll = false } = options;
      const extension = file.name.split('.').pop().toLowerCase();

      if (slot === 2 && extension !== 'pdf') {
        Alert.alert(
          'Second window',
          'The second viewer supports PDF files only so you can arrange two PDFs side by side and move pages between them.'
        );
        return;
      }

      setLoading(true);
      try {
        if (extension === 'pdf') {
          const fileInfo = await LegacyFileSystem.getInfoAsync(file.uri, { md5: true });
          const nextDocumentKey =
            fileInfo.md5 || `${file.name}-${fileInfo.size || Date.now()}`;
          const [originalBase64, persistedAnnotations, workingPdfBase64] =
            await Promise.all([
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

          if (slot === 1 && hasSavedEdits) {
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
          } else if (slot === 2 && hasSavedEdits) {
            await annotationWriteQueueRef.current;
            activeBase64 = workingPdfBase64 || originalBase64;
          }

          if (slot === 1) {
            if (replaceAll) {
              clearSecondaryDocument();
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
            setViewerHtml(
              buildViewerHtml(activeBase64, nextAnnotations, 'view', 1, {
                viewerSlot: 1,
                crossPaneDragEnabled: false,
                dualLayout,
              })
            );
            setViewerInstanceId((currentValue) => currentValue + 1);
            setShowPageEditorModal(false);
            setShowConvertModal(false);
            setExportPageImageBusy(false);
            pageImageExportRef.current.requestId = null;
            resetConversionState();
            resetOfficePreviewState();
            setPageNumberInput('1');
            setActivePane(1);
          } else {
            setSecondaryUri(file.uri);
            setSecondaryName(file.name);
            setSecondaryType(extension);
            setSecondaryKey(nextDocumentKey);
            setSecondaryBase64(activeBase64);
            setSecondaryTotalPages(0);
            setSecondaryAnnotationState(nextAnnotations);
            setSecondaryViewerHtml(
              buildViewerHtml(activeBase64, nextAnnotations, 'view', 1, {
                viewerSlot: 2,
                crossPaneDragEnabled: true,
                dualLayout,
              })
            );
            setSecondaryViewerInstanceId((v) => v + 1);
            setActivePane(2);
            if (documentType === 'pdf' && base64Data) {
              const p = Number.parseInt(pageNumberInput, 10);
              const curPage = Number.isInteger(p) && p > 0 ? p : 1;
              setViewerHtml(
                buildViewerHtml(base64Data, annotationState, activeTool, curPage, {
                  viewerSlot: 1,
                  crossPaneDragEnabled: true,
                  dualLayout,
                })
              );
              setViewerInstanceId((v) => v + 1);
            }
          }
        } else {
          if (slot === 2) {
            return;
          }
          if (replaceAll) {
            clearSecondaryDocument();
          }
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
          setExportPageImageBusy(false);
          pageImageExportRef.current.requestId = null;
          resetConversionState();
          resetOfficePreviewState();
          setPageNumberInput('1');
          setActivePane(1);
        }
      } catch (error) {
        Alert.alert('Error', 'Failed to load document: ' + error.message);
      } finally {
        setLoading(false);
      }
    },
    [
      activeTool,
      annotationState,
      base64Data,
      buildViewerHtml,
      clearPersistedPdfEditState,
      clearSecondaryDocument,
      documentType,
      loadPersistedAnnotations,
      loadWorkingPdfCopy,
      pageNumberInput,
      promptForPdfResumeChoice,
      resetConversionState,
      resetOfficePreviewState,
      dualLayout,
    ]
  );

  const pickDocumentIntoSlot = useCallback(
    async (slot, options = {}) => {
      try {
        const result = await DocumentPicker.getDocumentAsync({
          type: DOCUMENT_PICKER_TYPES,
          copyToCacheDirectory: true,
        });

        if (result.canceled) {
          return;
        }

        await applyPickedAssetToSlot(slot, result.assets[0], options);
      } catch (error) {
        Alert.alert('Error', 'Failed to pick document: ' + error.message);
      }
    },
    [applyPickedAssetToSlot]
  );

  useEffect(() => {
    const pending = consumePendingEditDocument();
    if (!pending) {
      return;
    }
    void applyPickedAssetToSlot(
      1,
      { uri: pending.uri, name: pending.name },
      { replaceAll: true }
    );
  }, [applyPickedAssetToSlot]);

  const openDocumentPicker = useCallback(() => {
    if (!documentUri) {
      pickDocumentIntoSlot(1);
      return;
    }
    if (!secondaryUri) {
      setShowOpenChoiceModal(true);
      return;
    }
    Alert.alert(
      'Two documents open',
      'Close the second PDF, or replace everything with a new file.',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Close second',
          onPress: () => clearSecondaryDocument(),
        },
        {
          text: 'New file',
          onPress: () => {
            clearSecondaryDocument();
            pickDocumentIntoSlot(1, { replaceAll: true });
          },
        },
      ]
    );
  }, [clearSecondaryDocument, documentUri, pickDocumentIntoSlot, secondaryUri]);

  const pickDocument = openDocumentPicker;

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

      void recordSavedFile({
        uri: nextUri,
        fileName: newName,
        mimeType: documentType ? `${documentType}` : null,
        source: 'document_copy',
      });

      Alert.alert('Success', 'Document copied as: ' + newName);
    } catch (error) {
      Alert.alert('Error', 'Failed to copy: ' + error.message);
    }
  }, [documentName, documentType, documentUri]);

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

  const runOfficePreviewExport = useCallback(() => {
    const canExport =
      activePane === 1
        ? officePreviewMeta.canConvertToPdf
        : secondaryOfficeMeta.canConvertToPdf;
    if (!canExport || officePdfBusy) {
      return;
    }

    officeExportSlotRef.current = activePane;

    const requestId = Date.now().toString();
    officePreviewExportRef.current = {
      requestId,
      chunks: [],
    };
    setOfficePdfBusy(true);

    sendCommand({
      command: 'exportPreview',
      requestId,
      title: activePane === 1 ? documentName : secondaryName,
    });
  }, [
    activePane,
    documentName,
    officePdfBusy,
    officePreviewMeta.canConvertToPdf,
    secondaryName,
    secondaryOfficeMeta.canConvertToPdf,
    sendCommand,
  ]);

  const promptSaveOfficeDocument = useCallback(() => {
    const canExport =
      activePane === 1
        ? officePreviewMeta.canConvertToPdf
        : secondaryOfficeMeta.canConvertToPdf;
    if (!canExport || officePdfBusy) {
      return;
    }

    Alert.alert(
      'Save edited document',
      'Edits are exported as a PDF. Native Word, Excel, and PowerPoint formats cannot be written back from this preview.',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Save as copy',
          onPress: () => {
            officeSaveModeRef.current = 'copy';
            runOfficePreviewExport();
          },
        },
        {
          text: 'Replace with PDF',
          style: 'destructive',
          onPress: () => {
            Alert.alert(
              'Replace with PDF?',
              'The Office file will close and the viewer will open a PDF with the same base name. You can share it like any saved PDF.',
              [
                { text: 'Cancel', style: 'cancel' },
                {
                  text: 'Continue',
                  onPress: () => {
                    officeSaveModeRef.current = 'replace';
                    runOfficePreviewExport();
                  },
                },
              ]
            );
          },
        },
      ]
    );
  }, [
    activePane,
    officePdfBusy,
    officePreviewMeta.canConvertToPdf,
    runOfficePreviewExport,
    secondaryOfficeMeta.canConvertToPdf,
  ]);

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
      const activePdfType = activePane === 1 ? documentType : secondaryType;
      if (activePdfType !== 'pdf') {
        return;
      }

      const formatConfig = PdfConversionFormats.PDF_CONVERSION_FORMATS.find(
        (conversionFormat) => conversionFormat.id === formatId
      );

      if (!formatConfig) {
        return;
      }

      try {
        setShowConvertModal(false);
        setConversionStatus('Preparing edited PDF...');

        const bakedPdfBase64 = await buildEditedPdfBase64();
        const exportBaseName = activePane === 1 ? documentName : secondaryName;
        const outputFileName = buildConvertedFileName(
          exportBaseName,
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
    [
      activePane,
      buildEditedPdfBase64,
      documentName,
      documentType,
      resetConversionState,
      secondaryName,
      secondaryType,
    ]
  );

  const finalizeOfficePreviewPdfExport = useCallback(
    async (html, mode = 'copy') => {
      const exportLabel =
        officeExportSlotRef.current === 2 ? secondaryName : documentName;

      const printResult = await Print.printToFileAsync({
        html,
      });
      const convertedPdfBase64 = await LegacyFileSystem.readAsStringAsync(
        printResult.uri,
        {
          encoding: LegacyFileSystem.EncodingType.Base64,
        }
      );

      if (mode === 'replace') {
        const stripped =
          exportLabel.replace(/\.[^/.]+$/, '') || exportLabel || 'document';
        const pdfFileName = `${sanitizeFileSegment(stripped)}.pdf`;
        const outputUri = `${Paths.document.uri}${pdfFileName}`;

        await LegacyFileSystem.writeAsStringAsync(outputUri, convertedPdfBase64, {
          encoding: LegacyFileSystem.EncodingType.Base64,
        });

        void recordSavedFile({
          uri: outputUri,
          fileName: pdfFileName,
          mimeType: 'application/pdf',
          source: 'office_replace_pdf',
        });

        const deviceSaveResult =
          Platform.OS === 'android'
            ? await savePdfToAndroidDeviceFolder(convertedPdfBase64, pdfFileName)
            : { savedToDevice: false, reason: 'unsupported_platform' };

        transitionOfficeExportToPdf(convertedPdfBase64, outputUri, pdfFileName);

        const canShare = await Sharing.isAvailableAsync();
        const saveSummary = [
          `Saved as ${pdfFileName}. The viewer is now showing this PDF.`,
        ];

        if (deviceSaveResult.savedToDevice) {
          saveSummary.push(
            'A copy was also saved to the folder you picked on your phone.'
          );
        } else if (Platform.OS === 'android') {
          saveSummary.push(
            'To save a copy in Downloads, allow the folder picker when prompted.'
          );
        }

        Alert.alert('Document saved', saveSummary.join(' '), [
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
        return;
      }

      const convertedFileName = buildConvertedFileName(exportLabel, 'pdf');
      const outputUri = `${Paths.document.uri}${convertedFileName}`;

      await LegacyFileSystem.writeAsStringAsync(outputUri, convertedPdfBase64, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });

      void recordSavedFile({
        uri: outputUri,
        fileName: convertedFileName,
        mimeType: 'application/pdf',
        source: 'office_export_pdf',
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
    [documentName, secondaryName, transitionOfficeExportToPdf]
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

            void recordSavedFile({
              uri: outputUri,
              fileName,
              mimeType,
              source: 'converter_export',
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
      const tp = activePane === 1 ? totalPages : secondaryTotalPages;
      const upperBound = allowInsertAfterLast ? tp : tp;

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
    [activePane, pageNumberInput, secondaryTotalPages, totalPages]
  );

  const removePdfPageAt = useCallback(
    async (pane, pageNumber) => {
      const isPrimary = pane === 1;
      const pdfBase64 = isPrimary ? base64Data : secondaryBase64;
      const pdfKey = isPrimary ? documentKey : secondaryKey;
      const pdfAnnotations = isPrimary ? annotationState : secondaryAnnotationState;
      const tp = isPrimary ? totalPages : secondaryTotalPages;

      if (!pageNumber || !pdfBase64 || !pdfKey) {
        return;
      }

      if (tp <= 1) {
        Alert.alert('Cannot Remove Page', 'A PDF must keep at least one page.');
        return;
      }

      try {
        setPageEditorBusy(true);
        const pdfDocument = await PDFDocument.load(pdfBase64);
        pdfDocument.removePage(pageNumber - 1);

        const nextAnnotations = removePageFromAnnotationState(pdfAnnotations, pageNumber);
        const nextBase64 = await pdfDocument.saveAsBase64();
        const nextPageCount = pdfDocument.getPageCount();
        const nextSelectedPage = Math.max(1, Math.min(pageNumber, nextPageCount));

        await annotationWriteQueueRef.current;
        await writeWorkingPdfCopy(pdfKey, nextBase64);
        await writeAnnotationSnapshot(pdfKey, nextAnnotations);
        if (isPrimary) {
          applyPdfEditorState(
            nextBase64,
            nextAnnotations,
            nextPageCount,
            'view',
            nextSelectedPage
          );
        } else {
          applySecondaryPdfEditorState(
            nextBase64,
            nextAnnotations,
            nextPageCount,
            'view',
            nextSelectedPage
          );
        }
        setShowPageEditorModal(false);
      } catch (error) {
        Alert.alert('Page Update Failed', error.message);
      } finally {
        setPageEditorBusy(false);
      }
    },
    [
      annotationState,
      applyPdfEditorState,
      applySecondaryPdfEditorState,
      base64Data,
      documentKey,
      secondaryAnnotationState,
      secondaryBase64,
      secondaryKey,
      secondaryTotalPages,
      totalPages,
      writeAnnotationSnapshot,
      writeWorkingPdfCopy,
    ]
  );

  const removePdfPage = useCallback(async () => {
    const pageNumber = getValidatedPageNumber(false);
    if (!pageNumber) {
      return;
    }
    await removePdfPageAt(activePane, pageNumber);
  }, [activePane, getValidatedPageNumber, removePdfPageAt]);

  const reorderPdfPageInPane = useCallback(
    async (slot, fromPage, toAfterPage) => {
      const isPrimary = slot === 1;
      const pdfBase64 = isPrimary ? base64Data : secondaryBase64;
      const pdfKey = isPrimary ? documentKey : secondaryKey;
      const pdfAnnotations = isPrimary ? annotationState : secondaryAnnotationState;
      const tp = isPrimary ? totalPages : secondaryTotalPages;

      if (
        !pdfBase64 ||
        !pdfKey ||
        !Number.isInteger(fromPage) ||
        !Number.isInteger(toAfterPage) ||
        fromPage < 1 ||
        toAfterPage < 1 ||
        fromPage > tp ||
        toAfterPage > tp
      ) {
        return;
      }
      if (fromPage === toAfterPage || tp <= 1) {
        return;
      }

      try {
        setPageEditorBusy(true);
        const pdfDocument = await PDFDocument.load(pdfBase64);
        const fromIdx = fromPage - 1;
        let insertAt = toAfterPage;
        const [copied] = await pdfDocument.copyPages(pdfDocument, [fromIdx]);
        pdfDocument.removePage(fromIdx);
        if (fromIdx < insertAt) {
          insertAt -= 1;
        }
        pdfDocument.insertPage(insertAt, copied);

        const nextAnnotations = reorderPageInAnnotationState(
          pdfAnnotations,
          fromPage,
          toAfterPage,
          tp
        );
        const nextBase64 = await pdfDocument.saveAsBase64();
        const nextPageCount = pdfDocument.getPageCount();

        const order = [];
        for (let p = 1; p <= tp; p += 1) {
          order.push(p);
        }
        const fi = order.indexOf(fromPage);
        order.splice(fi, 1);
        const ai = order.indexOf(toAfterPage);
        order.splice(ai === -1 ? order.length : ai + 1, 0, fromPage);
        const newSelectedPage = order.indexOf(fromPage) + 1;

        await annotationWriteQueueRef.current;
        await writeWorkingPdfCopy(pdfKey, nextBase64);
        await writeAnnotationSnapshot(pdfKey, nextAnnotations);
        if (isPrimary) {
          applyPdfEditorState(
            nextBase64,
            nextAnnotations,
            nextPageCount,
            'view',
            newSelectedPage
          );
        } else {
          applySecondaryPdfEditorState(
            nextBase64,
            nextAnnotations,
            nextPageCount,
            'view',
            newSelectedPage
          );
        }
      } catch (error) {
        Alert.alert('Reorder Failed', error.message);
      } finally {
        setPageEditorBusy(false);
      }
    },
    [
      annotationState,
      applyPdfEditorState,
      applySecondaryPdfEditorState,
      base64Data,
      documentKey,
      secondaryAnnotationState,
      secondaryBase64,
      secondaryKey,
      secondaryTotalPages,
      totalPages,
      writeAnnotationSnapshot,
      writeWorkingPdfCopy,
    ]
  );

  const insertBlankPdfPage = useCallback(async () => {
    const pageNumber = getValidatedPageNumber(true);
    const isPrimary = activePane === 1;
    const pdfBase64 = isPrimary ? base64Data : secondaryBase64;
    const pdfKey = isPrimary ? documentKey : secondaryKey;
    const pdfAnnotations = isPrimary ? annotationState : secondaryAnnotationState;

    if (!pageNumber || !pdfBase64 || !pdfKey) {
      return;
    }

    try {
      setPageEditorBusy(true);
      const pdfDocument = await PDFDocument.load(pdfBase64);
      const pages = pdfDocument.getPages();
      const referencePage = pages[Math.min(pageNumber - 1, pages.length - 1)];
      const { width, height } = referencePage.getSize();

      pdfDocument.insertPage(pageNumber, [width, height]);

      const nextAnnotations = insertPageIntoAnnotationState(pdfAnnotations, pageNumber);
      const nextBase64 = await pdfDocument.saveAsBase64();
      const nextPageCount = pdfDocument.getPageCount();
      const nextSelectedPage = Math.min(pageNumber + 1, nextPageCount);

      await annotationWriteQueueRef.current;
      await writeWorkingPdfCopy(pdfKey, nextBase64);
      await writeAnnotationSnapshot(pdfKey, nextAnnotations);
      if (isPrimary) {
        applyPdfEditorState(
          nextBase64,
          nextAnnotations,
          nextPageCount,
          'view',
          nextSelectedPage
        );
      } else {
        applySecondaryPdfEditorState(
          nextBase64,
          nextAnnotations,
          nextPageCount,
          'view',
          nextSelectedPage
        );
      }
      setShowPageEditorModal(false);
    } catch (error) {
      Alert.alert('Page Insert Failed', error.message);
    } finally {
      setPageEditorBusy(false);
    }
  }, [
    activePane,
    annotationState,
    applyPdfEditorState,
    applySecondaryPdfEditorState,
    base64Data,
    documentKey,
    getValidatedPageNumber,
    secondaryAnnotationState,
    secondaryBase64,
    secondaryKey,
    writeAnnotationSnapshot,
    writeWorkingPdfCopy,
  ]);

  const insertImagePdfPage = useCallback(async () => {
    const pageNumber = getValidatedPageNumber(true);
    const isPrimary = activePane === 1;
    const pdfBase64 = isPrimary ? base64Data : secondaryBase64;
    const pdfKey = isPrimary ? documentKey : secondaryKey;
    const pdfAnnotations = isPrimary ? annotationState : secondaryAnnotationState;

    if (!pageNumber || !pdfBase64 || !pdfKey) {
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

      const pdfDocument = await PDFDocument.load(pdfBase64);
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

      const nextAnnotations = insertPageIntoAnnotationState(pdfAnnotations, pageNumber);
      const nextBase64 = await pdfDocument.saveAsBase64();
      const nextPageCount = pdfDocument.getPageCount();
      const nextSelectedPage = Math.min(pageNumber + 1, nextPageCount);

      await annotationWriteQueueRef.current;
      await writeWorkingPdfCopy(pdfKey, nextBase64);
      await writeAnnotationSnapshot(pdfKey, nextAnnotations);
      if (isPrimary) {
        applyPdfEditorState(
          nextBase64,
          nextAnnotations,
          nextPageCount,
          'view',
          nextSelectedPage
        );
      } else {
        applySecondaryPdfEditorState(
          nextBase64,
          nextAnnotations,
          nextPageCount,
          'view',
          nextSelectedPage
        );
      }
      setShowPageEditorModal(false);
    } catch (error) {
      Alert.alert('Image Page Failed', error.message);
    } finally {
      setPageEditorBusy(false);
    }
  }, [
    activePane,
    annotationState,
    applyPdfEditorState,
    applySecondaryPdfEditorState,
    base64Data,
    documentKey,
    getValidatedPageNumber,
    secondaryAnnotationState,
    secondaryBase64,
    secondaryKey,
    writeAnnotationSnapshot,
    writeWorkingPdfCopy,
  ]);

  const performPdfPageTransfer = useCallback(
    async ({
      srcPane,
      tgtPane,
      srcPage,
      insertBefore,
      mode,
      closeModal = true,
    }) => {
      if (!secondaryUri || documentType !== 'pdf' || secondaryType !== 'pdf') {
        Alert.alert(
          'Page transfer',
          'Open two PDF files (use Import → Add second document) to move or copy pages between them.'
        );
        return;
      }

      if (srcPane === tgtPane) {
        Alert.alert('Page transfer', 'Pick two different documents.');
        return;
      }

      const tgtCount = tgtPane === 1 ? totalPages : secondaryTotalPages;

      if (!Number.isInteger(srcPage) || srcPage < 1) {
        Alert.alert('Invalid', 'Enter a valid source page number.');
        return;
      }

      if (
        !Number.isInteger(insertBefore) ||
        insertBefore < 1 ||
        insertBefore > tgtCount + 1
      ) {
        Alert.alert(
          'Invalid',
          `Insert position must be between 1 and ${tgtCount + 1} (before that page, or last+1 to append).`
        );
        return;
      }

      const srcBase64 = srcPane === 1 ? base64Data : secondaryBase64;
      const tgtBase64 = tgtPane === 1 ? base64Data : secondaryBase64;
      const srcKey = srcPane === 1 ? documentKey : secondaryKey;
      const tgtKey = tgtPane === 1 ? documentKey : secondaryKey;
      let srcAnn = srcPane === 1 ? annotationState : secondaryAnnotationState;
      let tgtAnn = tgtPane === 1 ? annotationState : secondaryAnnotationState;

      if (!srcBase64 || !tgtBase64 || !srcKey || !tgtKey) {
        return;
      }

      try {
        setPageTransferBusy(true);
        const srcPdf = await PDFDocument.load(srcBase64);
        const tgtPdf = await PDFDocument.load(tgtBase64);

        if (srcPage > srcPdf.getPageCount()) {
          Alert.alert('Invalid', 'Source page does not exist.');
          return;
        }

        const srcIndex = srcPage - 1;
        const insertIndex = insertBefore - 1;
        const [copied] = await tgtPdf.copyPages(srcPdf, [srcIndex]);
        tgtPdf.insertPage(insertIndex, copied);

        let nextSrcBase64 = srcBase64;
        if (mode === 'move') {
          if (srcPdf.getPageCount() <= 1) {
            Alert.alert('Cannot move', 'The source PDF must keep at least one page.');
            return;
          }
          srcPdf.removePage(srcIndex);
          nextSrcBase64 = await srcPdf.saveAsBase64();
          srcAnn = removePageFromAnnotationState(srcAnn, srcPage);
        }

        const nextTgtBase64 = await tgtPdf.saveAsBase64();
        const tgtInsertAfter = insertBefore - 1;
        tgtAnn = insertPageIntoAnnotationState(tgtAnn, tgtInsertAfter);

        const nextTgtCount = tgtPdf.getPageCount();

        await annotationWriteQueueRef.current;

        if (mode === 'move') {
          const nextSrcCount = srcPdf.getPageCount();
          if (srcPane === 1) {
            await writeWorkingPdfCopy(srcKey, nextSrcBase64);
            await writeAnnotationSnapshot(srcKey, srcAnn);
            applyPdfEditorState(
              nextSrcBase64,
              srcAnn,
              nextSrcCount,
              'view',
              Math.min(srcPage, nextSrcCount)
            );
          } else {
            await writeWorkingPdfCopy(srcKey, nextSrcBase64);
            await writeAnnotationSnapshot(srcKey, srcAnn);
            applySecondaryPdfEditorState(
              nextSrcBase64,
              srcAnn,
              nextSrcCount,
              'view',
              Math.min(srcPage, nextSrcCount)
            );
          }
        }

        if (tgtPane === 1) {
          await writeWorkingPdfCopy(tgtKey, nextTgtBase64);
          await writeAnnotationSnapshot(tgtKey, tgtAnn);
          applyPdfEditorState(
            nextTgtBase64,
            tgtAnn,
            nextTgtCount,
            'view',
            Math.min(insertBefore, nextTgtCount)
          );
        } else {
          await writeWorkingPdfCopy(tgtKey, nextTgtBase64);
          await writeAnnotationSnapshot(tgtKey, tgtAnn);
          applySecondaryPdfEditorState(
            nextTgtBase64,
            tgtAnn,
            nextTgtCount,
            'view',
            Math.min(insertBefore, nextTgtCount)
          );
        }

        if (closeModal) {
          setShowPageTransferModal(false);
        }
        setCrossPdfDropBanner(false);
        crossPdfDropPendingRef.current = null;
        Alert.alert(
          'Done',
          mode === 'move'
            ? 'Page moved. Annotations were adjusted where possible; review both PDFs.'
            : 'Page copied into the target PDF. Annotations on the new page start empty.'
        );
      } catch (error) {
        Alert.alert('Transfer failed', error?.message || 'Could not transfer page.');
      } finally {
        setPageTransferBusy(false);
      }
    },
    [
      annotationState,
      applyPdfEditorState,
      applySecondaryPdfEditorState,
      base64Data,
      documentKey,
      documentType,
      secondaryAnnotationState,
      secondaryBase64,
      secondaryKey,
      secondaryTotalPages,
      secondaryType,
      secondaryUri,
      totalPages,
      writeAnnotationSnapshot,
      writeWorkingPdfCopy,
    ]
  );

  const executePdfPageTransfer = useCallback(async () => {
    await performPdfPageTransfer({
      srcPane: pageTransferSourcePane,
      tgtPane: pageTransferTargetPane,
      srcPage: Number.parseInt(pageTransferSourcePage, 10),
      insertBefore: Number.parseInt(pageTransferInsertBefore, 10),
      mode: pageTransferMode,
      closeModal: true,
    });
  }, [
    pageTransferInsertBefore,
    pageTransferMode,
    pageTransferSourcePage,
    pageTransferSourcePane,
    pageTransferTargetPane,
    performPdfPageTransfer,
  ]);

  const cancelCrossPdfDrop = useCallback(() => {
    crossPdfDropPendingRef.current = null;
    setCrossPdfDropBanner(false);
  }, []);

  const appendCrossPdfPageToOtherEnd = useCallback(() => {
    const p = crossPdfDropPendingRef.current;
    if (!p || !secondaryUri) {
      return;
    }
    const tgtPane = p.sourcePane === 1 ? 2 : 1;
    const tgtCount = tgtPane === 1 ? totalPages : secondaryTotalPages;
    const mode = p.mode === 'move' ? 'move' : 'copy';
    crossPdfDropPendingRef.current = null;
    setCrossPdfDropBanner(false);
    void performPdfPageTransfer({
      srcPane: p.sourcePane,
      tgtPane,
      srcPage: p.sourcePage,
      insertBefore: tgtCount + 1,
      mode,
      closeModal: false,
    });
  }, [performPdfPageTransfer, secondaryUri, secondaryTotalPages, totalPages]);

  const exportPdfPageAsImage = useCallback(
    (format) => {
      const activePdf = activePane === 1 ? documentType : secondaryType;
      if (activePdf !== 'pdf' || exportPageImageBusy || pageEditorBusy) {
        return;
      }
      const pageNumber = getValidatedPageNumber(false);
      if (!pageNumber) {
        return;
      }
      exportContextPaneRef.current = activePane;
      const requestId = `${Date.now()}-${Math.random().toString(36).slice(2, 11)}`;
      pageImageExportRef.current.requestId = requestId;
      setExportPageImageBusy(true);
      const isLossy =
        format.mimeType === 'image/jpeg' || format.mimeType === 'image/webp';
      sendCommand({
        command: 'exportPageImage',
        requestId,
        page: pageNumber,
        mimeType: format.mimeType,
        ...(isLossy ? { quality: 0.92 } : {}),
      });
    },
    [
      activePane,
      documentType,
      exportPageImageBusy,
      getValidatedPageNumber,
      pageEditorBusy,
      secondaryType,
      sendCommand,
    ]
  );

  const runSaveAnnotatedPdfForPane = useCallback(
    async (pane) => {
      const pdfType = pane === 1 ? documentType : secondaryType;
      if (pdfType !== 'pdf') {
        return;
      }

      try {
        setSavingAnnotatedPdf(true);
        const { bakedPdfBase64, fileName, outputUri } = await buildEditedPdfFile(
          Paths.document.uri,
          pane
        );

        void recordSavedFile({
          uri: outputUri,
          fileName,
          mimeType: 'application/pdf',
          source: 'annotated_pdf',
        });

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
    },
    [buildEditedPdfFile, documentType, secondaryType]
  );

  const saveAnnotatedPdf = useCallback(async () => {
    const activePdfType = activePane === 1 ? documentType : secondaryType;
    if (activePdfType !== 'pdf') {
      return;
    }

    const dualPdf =
      secondaryUri && documentType === 'pdf' && secondaryType === 'pdf';
    if (dualPdf) {
      const label1 = documentName?.trim() || 'First PDF';
      const label2 = secondaryName?.trim() || 'Second PDF';
      Alert.alert('Save annotated PDF', 'Which document should be saved?', [
        { text: 'Cancel', style: 'cancel' },
        {
          text: label1,
          onPress: () => {
            void runSaveAnnotatedPdfForPane(1);
          },
        },
        {
          text: label2,
          onPress: () => {
            void runSaveAnnotatedPdfForPane(2);
          },
        },
      ]);
      return;
    }

    await runSaveAnnotatedPdfForPane(activePane);
  }, [
    activePane,
    documentName,
    documentType,
    runSaveAnnotatedPdfForPane,
    secondaryName,
    secondaryType,
    secondaryUri,
  ]);

  const saveTextDocumentToDevice = useCallback(async () => {
    if (!documentUri || !TEXT_FILE_EXTENSIONS.includes(documentType)) {
      return;
    }

    const run = textFileSaveRef.current;
    if (typeof run !== 'function') {
      Alert.alert('Save', 'Editor is still loading. Try again in a moment.');
      return;
    }

    try {
      setSavingTextFile(true);
      const ok = await run();
      if (!ok) {
        return;
      }

      void recordSavedFile({
        uri: documentUri,
        fileName: documentName || `edited.${documentType}`,
        mimeType: getTextMimeTypeForExtension(documentType),
        source: 'text_document',
      });

      const fileBase64 = await LegacyFileSystem.readAsStringAsync(documentUri, {
        encoding: LegacyFileSystem.EncodingType.Base64,
      });
      const mimeType = getTextMimeTypeForExtension(documentType);
      const deviceSaveResult =
        Platform.OS === 'android'
          ? await saveBase64ToAndroidDeviceFolder(
              fileBase64,
              documentName || `edited.${documentType}`,
              mimeType
            )
          : { savedToDevice: false, reason: 'unsupported_platform' };

      const canShare = await Sharing.isAvailableAsync();
      const saveSummary = ['Your changes were written to this document.'];

      if (deviceSaveResult.savedToDevice) {
        saveSummary.push(
          'A second copy was also saved to the folder you picked on your phone.'
        );
      } else if (Platform.OS === 'android') {
        saveSummary.push(
          'To save a copy in Downloads or another folder, allow access when the system folder picker appears.'
        );
      }

      Alert.alert('Saved', saveSummary.join(' '), [
        ...(canShare
          ? [
              {
                text: 'Share',
                onPress: () => Sharing.shareAsync(documentUri),
              },
            ]
          : []),
        { text: 'OK' },
      ]);
    } catch (error) {
      Alert.alert(
        'Save Failed',
        error && error.message ? error.message : 'Could not save this file.'
      );
    } finally {
      setSavingTextFile(false);
    }
  }, [documentName, documentType, documentUri]);

  const processViewerMessage = useCallback(
    (data, slot) => {
      const isPrimary = slot === 1;

      switch (data.type) {
        case 'pdfLoaded':
          if (isPrimary) {
            setTotalPages(data.totalPages);
          } else {
            setSecondaryTotalPages(data.totalPages);
          }
          injectCommandToSlot(slot, { command: 'setMode', mode: activeTool });
          injectCommandToSlot(slot, { command: 'setDrawColor', color: activeColor });
          injectCommandToSlot(slot, {
            command: 'setHighlightColor',
            color: activeColor + HIGHLIGHT_ALPHA_HEX,
          });
          injectCommandToSlot(slot, { command: 'setDrawBrushWidth', px: drawBrushPx });
          injectCommandToSlot(slot, { command: 'setMarkerBrushWidth', px: markerBrushPx });
          break;
        case 'requestTextInput':
          setActivePane(slot);
          setTextNoteModal({ page: data.page, x: data.x, y: data.y });
          setNoteText('');
          break;
        case 'annotationSelected':
          setSelectedAnnotation(data.annotation || null);
          break;
        case 'crossPdfPickSource':
          if (
            data.slot &&
            Number.isInteger(data.page) &&
            secondaryUri &&
            documentType === 'pdf' &&
            secondaryType === 'pdf'
          ) {
            crossPdfDropPendingRef.current = {
              sourcePane: data.slot,
              sourcePage: data.page,
              mode: crossPdfDropMode === 'move' ? 'move' : 'copy',
            };
            setCrossPdfDropBanner(true);
            setActivePane(data.slot);
          }
          break;
        case 'pageReorderDrag':
          if (
            data.slot &&
            Number.isInteger(data.fromPage) &&
            Number.isInteger(data.toAfterPage)
          ) {
            void reorderPdfPageInPane(data.slot, data.fromPage, data.toAfterPage);
          }
          break;
        case 'pageMenuAction':
          if (!data.slot || !Number.isInteger(data.page)) {
            break;
          }
          setActivePane(data.slot);
          if (data.action === 'editPages') {
            setPageNumberInput(String(data.page));
            setShowPageEditorModal(true);
          } else if (data.action === 'removePage') {
            Alert.alert(
              'Remove page',
              `Remove page ${data.page} from this document?`,
              [
                { text: 'Cancel', style: 'cancel' },
                {
                  text: 'Remove',
                  style: 'destructive',
                  onPress: () => {
                    void removePdfPageAt(data.slot, data.page);
                  },
                },
              ]
            );
          }
          break;
        case 'pageSelected':
          setActivePane(slot);
          if (Number.isInteger(data.page)) {
            setPageNumberInput(String(data.page));
          }
          if (crossPdfDropPendingRef.current) {
            const pending = crossPdfDropPendingRef.current;
            if (slot === pending.sourcePane) {
              crossPdfDropPendingRef.current = null;
              setCrossPdfDropBanner(false);
            } else {
              const insertAfterPage = data.page;
              const insertBefore = insertAfterPage + 1;
              const mode = pending.mode === 'move' ? 'move' : 'copy';
              crossPdfDropPendingRef.current = null;
              setCrossPdfDropBanner(false);
              void performPdfPageTransfer({
                srcPane: pending.sourcePane,
                tgtPane: slot,
                srcPage: pending.sourcePage,
                insertBefore,
                mode,
                closeModal: false,
              });
            }
          }
          break;
        case 'pageLongPressed':
          break;
        case 'pageExportImageReady':
          if (data.requestId !== pageImageExportRef.current.requestId) {
            break;
          }
          pageImageExportRef.current.requestId = null;
          setExportPageImageBusy(false);
          (async () => {
            try {
              const page = Number(data.page) || 1;
              const mime = data.mimeType || 'image/png';
              const ext =
                mime === 'image/jpeg' || mime === 'image/jpg'
                  ? 'jpg'
                  : mime === 'image/webp'
                    ? 'webp'
                    : 'png';
              const srcName =
                exportContextPaneRef.current === 2 ? secondaryName : documentName;
              const rawBase = (srcName || 'document').replace(/\.[^/.]+$/, '');
              const safeBase =
                rawBase.replace(/[^\w\-]+/g, '_').slice(0, 72) || 'document';
                const fileName = `${safeBase}_page_${page}.${ext}`;
                const outputUri = `${Paths.cache.uri}${fileName}`;
                await LegacyFileSystem.writeAsStringAsync(outputUri, data.base64, {
                  encoding: LegacyFileSystem.EncodingType.Base64,
                });

                const imageMimeType =
                  ext === 'jpg'
                    ? 'image/jpeg'
                    : ext === 'webp'
                      ? 'image/webp'
                      : 'image/png';

                void recordSavedFile({
                  uri: outputUri,
                  fileName,
                  mimeType: imageMimeType,
                  source: 'pdf_page_image',
                });

                const deviceSaveResult =
                  Platform.OS === 'android'
                    ? await saveBase64ToAndroidDeviceFolder(
                        data.base64,
                        fileName,
                        imageMimeType
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
                    'To save a copy in Downloads or another folder, allow access when the folder picker appears.'
                  );
                }

                Alert.alert('Page image ready', saveSummary.join(' '), [
                  ...(canShare
                    ? [
                        {
                          text: 'Share',
                          onPress: () => {
                            Sharing.shareAsync(outputUri).catch(() => {});
                          },
                        },
                      ]
                    : []),
                  { text: 'OK' },
                ]);
              } catch (error) {
                Alert.alert(
                  'Save failed',
                  error && error.message ? error.message : 'Could not save the image.'
                );
              }
            })();
            break;
        case 'pageExportImageError':
          if (data.requestId !== pageImageExportRef.current.requestId) {
            break;
          }
          pageImageExportRef.current.requestId = null;
          setExportPageImageBusy(false);
          Alert.alert(
            'Export failed',
            data.message || 'Could not export this page as an image.'
          );
          break;
        case 'officePreviewReady':
          if (isPrimary) {
            setOfficePreviewMeta({
              previewKind: data.previewKind || '',
              canConvertToPdf: !!data.canConvertToPdf,
              isEditable: !!data.isEditable,
            });
          } else {
            setSecondaryOfficeMeta({
              previewKind: data.previewKind || '',
              canConvertToPdf: !!data.canConvertToPdf,
              isEditable: !!data.isEditable,
            });
          }
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

            finalizeOfficePreviewPdfExport(html, officeSaveModeRef.current)
              .catch((error) => {
                Alert.alert('PDF Conversion Failed', error.message);
              })
              .finally(() => {
                officeSaveModeRef.current = 'copy';
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
          if (isPrimary) {
            setAnnotationState(nextAnnotations);
            persistAnnotationState(documentKey, nextAnnotations);
          } else {
            setSecondaryAnnotationState(nextAnnotations);
            persistAnnotationState(secondaryKey, nextAnnotations);
          }
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

            return bucket.some((annotation) => annotation.id === currentSelection.id)
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
        default:
          break;
      }
    },
    [
      activeColor,
      activeTool,
      crossPdfDropMode,
      documentKey,
      documentName,
      documentType,
      drawBrushPx,
      finalizeOfficePreviewPdfExport,
      injectCommandToSlot,
      markerBrushPx,
      performPdfPageTransfer,
      persistAnnotationState,
      removePdfPageAt,
      reorderPdfPageInPane,
      secondaryKey,
      secondaryName,
      secondaryType,
      secondaryUri,
    ]
  );

  const handleWebViewMessagePrimary = useCallback(
    (event) => {
      try {
        const data = JSON.parse(event.nativeEvent.data);
        processViewerMessage(data, 1);
      } catch {}
    },
    [processViewerMessage]
  );

  const handleWebViewMessageSecondary = useCallback(
    (event) => {
      try {
        const data = JSON.parse(event.nativeEvent.data);
        processViewerMessage(data, 2);
      } catch {}
    },
    [processViewerMessage]
  );

  const dualLayoutHtmlSyncRef = useRef(true);
  useEffect(() => {
    if (dualLayoutHtmlSyncRef.current) {
      dualLayoutHtmlSyncRef.current = false;
      return;
    }
    if (documentType !== 'pdf' || !base64Data) {
      return;
    }
    const dualPdf = !!secondaryUri && secondaryType === 'pdf';
    const cur = Math.max(1, Number.parseInt(pageNumberInput, 10) || 1);
    setViewerHtml(
      buildViewerHtml(base64Data, annotationState, activeTool, cur, {
        viewerSlot: 1,
        crossPaneDragEnabled: dualPdf,
        dualLayout,
      })
    );
    setViewerInstanceId((v) => v + 1);
    if (secondaryUri && secondaryBase64 && secondaryType === 'pdf') {
      const cur2 = Math.max(
        1,
        Math.min(Math.max(secondaryTotalPages, 1), Number.parseInt(pageNumberInput, 10) || 1)
      );
      setSecondaryViewerHtml(
        buildViewerHtml(secondaryBase64, secondaryAnnotationState, activeTool, cur2, {
          viewerSlot: 2,
          crossPaneDragEnabled: dualPdf,
          dualLayout,
        })
      );
      setSecondaryViewerInstanceId((v) => v + 1);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps -- only refresh viewer HTML when split layout toggles
  }, [dualLayout]);

  const activeDocumentType = activePane === 1 ? documentType : secondaryType;
  const activeOfficeMeta =
    activePane === 1 ? officePreviewMeta : secondaryOfficeMeta;
  const activeTotalPages = activePane === 1 ? totalPages : secondaryTotalPages;

  useEffect(() => {
    if (crossPdfDropPendingRef.current) {
      crossPdfDropPendingRef.current.mode =
        crossPdfDropMode === 'move' ? 'move' : 'copy';
    }
  }, [crossPdfDropMode]);

  useEffect(() => {
    if (activeDocumentType !== 'pdf') {
      return;
    }
    sendCommand({ command: 'setDrawBrushWidth', px: drawBrushPx });
  }, [activeDocumentType, drawBrushPx, sendCommand]);

  useEffect(() => {
    if (activeDocumentType !== 'pdf') {
      return;
    }
    sendCommand({ command: 'setMarkerBrushWidth', px: markerBrushPx });
  }, [activeDocumentType, markerBrushPx, sendCommand]);

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

    if (TEXT_FILE_EXTENSIONS.includes(extension)) {
      return (
        <View style={styles.nonPdfContainer}>
          <EditableTextFileViewer
            uri={documentUri}
            extension={extension}
            onRegisterSave={(saveFn) => {
              textFileSaveRef.current = saveFn;
            }}
          />
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
            onMessage={handleWebViewMessagePrimary}
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

  const renderDocumentPane = (pane) => {
    const isPrimary = pane === 1;
    const pdfHtml = isPrimary ? viewerHtml : secondaryViewerHtml;
    const pdfKey = isPrimary ? documentKey : secondaryKey;
    const pdfUri = isPrimary ? documentUri : secondaryUri;
    const pdfInstanceId = isPrimary ? viewerInstanceId : secondaryViewerInstanceId;
    const webRef = isPrimary ? webViewRef : webViewRef2;
    const onMsg = isPrimary ? handleWebViewMessagePrimary : handleWebViewMessageSecondary;
    const type = isPrimary ? documentType : secondaryType;

    if (isPrimary) {
      if (type === 'pdf' && pdfHtml) {
        return (
          <WebView
            key={`${pdfKey || pdfUri}-${pdfInstanceId}`}
            ref={webRef}
            source={{ html: pdfHtml }}
            style={styles.webView}
            originWhitelist={['*']}
            javaScriptEnabled={true}
            domStorageEnabled={true}
            onMessage={onMsg}
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
        );
      }
      return renderNonPdfViewer();
    }

    if (type === 'pdf' && pdfHtml) {
      return (
        <WebView
          key={`${pdfKey || pdfUri}-${pdfInstanceId}`}
          ref={webRef}
          source={{ html: pdfHtml }}
          style={styles.webView}
          originWhitelist={['*']}
          javaScriptEnabled={true}
          domStorageEnabled={true}
          onMessage={onMsg}
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
      );
    }

    return (
      <View style={styles.unsupported}>
        <Text style={styles.unsupportedText}>Second viewer</Text>
        <Text style={styles.unsupportedHint}>Loading PDF…</Text>
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
        {activeTotalPages > 0 ? (
          <View style={styles.pageCountPill}>
            <Text style={styles.pageCountLabel}>{activeTotalPages}</Text>
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
            Supported: PDF, Office, TXT, MD, HTML, JSON, code files, Images
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
              {secondaryUri ? (
                isDualPdfMode ? (
                  <>
                    <View style={styles.layoutToggleGroup}>
                      <TouchableOpacity
                        style={[
                          styles.layoutToggleBtn,
                          dualLayout === 'row' && styles.layoutToggleBtnActive,
                        ]}
                        onPress={() => setDualLayout('row')}
                        accessibilityLabel="Split view: PDFs side by side"
                      >
                        <MaterialIcons
                          name="view-column"
                          size={18}
                          color={dualLayout === 'row' ? colors.primaryDim : colors.onSurface}
                        />
                        <Text
                          style={[
                            styles.layoutToggleBtnLabel,
                            dualLayout === 'row' && styles.layoutToggleBtnLabelActive,
                          ]}
                          numberOfLines={1}
                        >
                          Split
                        </Text>
                      </TouchableOpacity>
                      <TouchableOpacity
                        style={[
                          styles.layoutToggleBtn,
                          dualLayout === 'column' && styles.layoutToggleBtnActive,
                        ]}
                        onPress={() => setDualLayout('column')}
                        accessibilityLabel="Stacked view: PDFs top and bottom"
                      >
                        <MaterialIcons
                          name="view-stream"
                          size={18}
                          color={
                            dualLayout === 'column' ? colors.primaryDim : colors.onSurface
                          }
                        />
                        <Text
                          style={[
                            styles.layoutToggleBtnLabel,
                            dualLayout === 'column' && styles.layoutToggleBtnLabelActive,
                          ]}
                          numberOfLines={1}
                        >
                          Stack
                        </Text>
                      </TouchableOpacity>
                    </View>
                    <TouchableOpacity
                      style={styles.mergePagesBarBtn}
                      onPress={() => setShowPageTransferModal(true)}
                      accessibilityLabel="Merge pages between the two PDFs"
                      accessibilityHint="Opens a form to copy or move a page into the other document"
                    >
                      <MaterialIcons name="post-add" size={22} color={colors.primaryDim} />
                    </TouchableOpacity>
                  </>
                ) : (
                  <>
                    <TouchableOpacity
                      style={[
                        styles.actionIconBtn,
                        dualLayout === 'row' && styles.actionIconBtnActivePane,
                      ]}
                      onPress={() => setDualLayout('row')}
                      accessibilityLabel="Side by side layout"
                    >
                      <MaterialIcons
                        name="view-column"
                        size={22}
                        color={dualLayout === 'row' ? colors.primary : colors.onSurface}
                      />
                    </TouchableOpacity>
                    <TouchableOpacity
                      style={[
                        styles.actionIconBtn,
                        dualLayout === 'column' && styles.actionIconBtnActivePane,
                      ]}
                      onPress={() => setDualLayout('column')}
                      accessibilityLabel="Top and bottom layout"
                    >
                      <MaterialIcons
                        name="view-stream"
                        size={22}
                        color={dualLayout === 'column' ? colors.primary : colors.onSurface}
                      />
                    </TouchableOpacity>
                  </>
                )
              ) : null}
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
              {activeDocumentType === 'pdf' && (
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
              {activeDocumentType !== 'pdf' &&
                (TEXT_FILE_EXTENSIONS.includes(activeDocumentType) ||
                  activeOfficeMeta.canConvertToPdf ||
                  (activeOfficeMeta.isEditable &&
                    activeOfficeMeta.previewKind === 'text')) && (
                  <>
                    <View style={styles.actionBarDivider} />
                    {TEXT_FILE_EXTENSIONS.includes(activeDocumentType) && (
                      <TouchableOpacity
                        style={[
                          styles.actionIconBtnPrimary,
                          savingTextFile && styles.actionIconBtnDisabled,
                        ]}
                        disabled={savingTextFile}
                        onPress={saveTextDocumentToDevice}
                        accessibilityLabel="Save file to device"
                      >
                        <MaterialIcons
                          name={savingTextFile ? 'hourglass-empty' : 'save-alt'}
                          size={22}
                          color={colors.onPrimary}
                        />
                      </TouchableOpacity>
                    )}
                    {activeOfficeMeta.canConvertToPdf && (
                      <TouchableOpacity
                        style={[
                          styles.actionIconBtnPrimary,
                          officePdfBusy && styles.actionIconBtnDisabled,
                        ]}
                        disabled={officePdfBusy}
                        onPress={promptSaveOfficeDocument}
                        accessibilityLabel="Save edited document as PDF"
                      >
                        <MaterialIcons
                          name={officePdfBusy ? 'hourglass-empty' : 'save-alt'}
                          size={22}
                          color={colors.onPrimary}
                        />
                      </TouchableOpacity>
                    )}
                    {activeOfficeMeta.isEditable && (
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

          {isDualPdfMode && showDualPdfGuide && !crossPdfDropBanner ? (
            <View style={styles.dualPdfGuideCard}>
              <View style={styles.dualPdfGuideHeader}>
                <MaterialIcons name="menu-book" size={22} color={colors.primaryDim} />
                <Text style={styles.dualPdfGuideTitle}>Two PDFs — quick tips</Text>
                <TouchableOpacity
                  onPress={() => setShowDualPdfGuide(false)}
                  hitSlop={12}
                  accessibilityLabel="Dismiss tips"
                >
                  <MaterialIcons name="close" size={20} color={colors.onSurface} />
                </TouchableOpacity>
              </View>
              <Text style={styles.dualPdfGuideLine}>
                <Text style={styles.dualPdfGuideBold}>Reorder: </Text>
                double-tap a page, then drag onto another page in the same PDF.
              </Text>
              <Text style={styles.dualPdfGuideLine}>
                <Text style={styles.dualPdfGuideBold}>Add to the other PDF: </Text>
                drag toward the outer screen edge, or tap ⇄ then tap where it should go.
              </Text>
              <Text style={styles.dualPdfGuideLine}>
                <Text style={styles.dualPdfGuideBold}>Page menu ⋮: </Text>
                change page order or remove a page — or use Merge pages in the bar above.
              </Text>
            </View>
          ) : null}

          {crossPdfDropBanner && secondaryUri ? (
            <View style={styles.crossPdfBannerCard}>
              <View style={styles.crossPdfBannerTopRow}>
                <View style={styles.crossPdfBannerIconWrap}>
                  <MaterialIcons name="swap-horiz" size={22} color={colors.primaryDim} />
                </View>
                <View style={styles.crossPdfBannerTextBlock}>
                  <Text style={styles.crossPdfBannerKicker}>Cross-document</Text>
                  <Text style={styles.crossPdfBannerTitle}>Send this page to the other PDF</Text>
                  <Text style={styles.crossPdfBannerSub}>
                    Same PDF: drop on another page to reorder. Other PDF: drag to the far edge or
                    finish here. Tap ⇄ on a page for a quick target. Use ⋮ for page list or remove.
                  </Text>
                </View>
              </View>
              <Text style={styles.crossPdfBannerStepLabel}>How to insert</Text>
              <View style={styles.crossPdfBannerModes}>
                <TouchableOpacity
                  style={[
                    styles.crossPdfModeChip,
                    crossPdfDropMode === 'copy' && styles.crossPdfModeChipOn,
                  ]}
                  onPress={() => setCrossPdfDropMode('copy')}
                  accessibilityRole="button"
                  accessibilityState={{ selected: crossPdfDropMode === 'copy' }}
                >
                  <MaterialIcons
                    name="content-copy"
                    size={16}
                    color={
                      crossPdfDropMode === 'copy' ? colors.primaryDim : colors.onSurface
                    }
                  />
                  <Text
                    style={[
                      styles.crossPdfModeChipText,
                      crossPdfDropMode === 'copy' && styles.crossPdfModeChipTextOn,
                    ]}
                  >
                    Copy
                  </Text>
                </TouchableOpacity>
                <TouchableOpacity
                  style={[
                    styles.crossPdfModeChip,
                    crossPdfDropMode === 'move' && styles.crossPdfModeChipOn,
                  ]}
                  onPress={() => setCrossPdfDropMode('move')}
                  accessibilityRole="button"
                  accessibilityState={{ selected: crossPdfDropMode === 'move' }}
                >
                  <MaterialIcons
                    name="drive-file-move"
                    size={16}
                    color={
                      crossPdfDropMode === 'move' ? colors.primaryDim : colors.onSurface
                    }
                  />
                  <Text
                    style={[
                      styles.crossPdfModeChipText,
                      crossPdfDropMode === 'move' && styles.crossPdfModeChipTextOn,
                    ]}
                  >
                    Move
                  </Text>
                </TouchableOpacity>
              </View>
              <View style={styles.crossPdfBannerActions}>
                <TouchableOpacity
                  style={styles.crossPdfBannerBtnPrimary}
                  onPress={appendCrossPdfPageToOtherEnd}
                  disabled={pageTransferBusy}
                >
                  <MaterialIcons name="vertical-align-bottom" size={18} color={colors.onPrimary} />
                  <Text style={styles.crossPdfBannerBtnPrimaryText}>Append at end</Text>
                </TouchableOpacity>
                <TouchableOpacity style={styles.crossPdfBannerBtnSecondary} onPress={cancelCrossPdfDrop}>
                  <Text style={styles.crossPdfBannerBtnSecondaryText}>Cancel</Text>
                </TouchableOpacity>
              </View>
            </View>
          ) : null}

          {activeDocumentType === 'pdf' && (
            <View style={styles.selectionBar}>
              <Text style={styles.selectionText}>
                {activeTool === 'view'
                  ? getSelectedAnnotationDescription(selectedAnnotation)
                  : 'Switch to Select mode to tap an annotation before deleting it.'}
              </Text>
            </View>
          )}

          {activeDocumentType !== 'pdf' && activeOfficeMeta.isEditable && (
            <View style={styles.selectionBar}>
              <Text style={styles.selectionText}>
                {activeOfficeMeta.previewKind === 'text'
                  ? 'Plain-text editing: tap the save icon to write your changes and optionally save a copy to a folder on your device (Android). Office files use the toolbar below; tap Save to export as PDF (copy or replace with PDF).'
                  : 'Edit with the toolbar: fonts (dropdown), weight, line height, headings, undo, colors, alignment, lists, and more. Tap an image to resize. Tap Save to export your edits as a PDF — choose a new copy or replace the open file with a PDF.'}
              </Text>
            </View>
          )}

          {activeDocumentType === 'pdf' && showToolbar && (
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

          {documentType !== 'pdf' &&
            officePreviewMeta.isEditable &&
            showToolbar &&
            officePreviewMeta.previewKind === 'text' && (
              <View style={styles.toolbar}>
                <Text style={styles.textFileToolbarHint}>
                  Plain text — use the save button in the action bar. Monospace / code
                  files: pinch-zoom the editor if needed.
                </Text>
              </View>
            )}

          {documentType !== 'pdf' &&
            officePreviewMeta.isEditable &&
            showToolbar &&
            officePreviewMeta.previewKind !== 'text' && (
            <View style={styles.toolbar}>
              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeIconScroll}
              >
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('undo')}
                  accessibilityLabel="Undo"
                >
                  <MaterialIcons name="undo" size={22} color={colors.onSurface} />
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.officeIconBtn}
                  onPress={() => applyOfficeStyleCommand('redo')}
                  accessibilityLabel="Redo"
                >
                  <MaterialIcons name="redo" size={22} color={colors.onSurface} />
                </TouchableOpacity>
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

              {activeOfficeMeta.previewKind === 'presentation' && (
                <>
                  <Text style={styles.officeSubheadingLabel}>Deck theme</Text>
                  <ScrollView
                    horizontal
                    showsHorizontalScrollIndicator={false}
                    contentContainerStyle={styles.officeOptionRow}
                  >
                    {PRESENTATION_THEME_OPTIONS.map((opt) => (
                      <TouchableOpacity
                        key={opt.value}
                        style={styles.officeOptionBtn}
                        onPress={() =>
                          applyOfficeStyleCommand('setPresentationTheme', opt.value)
                        }
                      >
                        <Text style={styles.officeOptionText}>{opt.label}</Text>
                      </TouchableOpacity>
                    ))}
                  </ScrollView>
                </>
              )}

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

              <TouchableOpacity
                style={styles.officeFontDropdownBtn}
                onPress={() => setOfficeFontModalVisible(true)}
                accessibilityLabel="Choose font family"
              >
                <MaterialIcons name="font-download" size={20} color={colors.primaryDim} />
                <Text style={styles.officeFontDropdownBtnText}>Font family</Text>
                <MaterialIcons
                  name="arrow-drop-down"
                  size={22}
                  color={colors.onSurface}
                />
              </TouchableOpacity>

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

              <Text style={styles.officeSubheadingLabel}>Font weight</Text>
              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeOptionRow}
              >
                {OFFICE_FONT_WEIGHT_OPTIONS.map((opt) => (
                  <TouchableOpacity
                    key={opt.value}
                    style={styles.officeOptionBtn}
                    onPress={() => applyOfficeStyleCommand('setFontWeight', opt.value)}
                  >
                    <Text style={styles.officeOptionText}>{opt.label}</Text>
                  </TouchableOpacity>
                ))}
              </ScrollView>

              <Text style={styles.officeSubheadingLabel}>Line height</Text>
              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeOptionRow}
              >
                {OFFICE_LINE_HEIGHT_OPTIONS.map((opt) => (
                  <TouchableOpacity
                    key={String(opt.value)}
                    style={styles.officeOptionBtn}
                    onPress={() => applyOfficeStyleCommand('setLineHeight', opt.value)}
                  >
                    <Text style={styles.officeOptionText}>{opt.label}</Text>
                  </TouchableOpacity>
                ))}
              </ScrollView>

              <Text style={styles.officeSubheadingLabel}>Paragraph / heading</Text>
              <ScrollView
                horizontal
                showsHorizontalScrollIndicator={false}
                contentContainerStyle={styles.officeOptionRow}
              >
                {OFFICE_HEADING_OPTIONS.map((opt) => (
                  <TouchableOpacity
                    key={opt.value}
                    style={styles.officeOptionBtn}
                    onPress={() => applyOfficeStyleCommand('setHeading', opt.value)}
                  >
                    <Text style={styles.officeOptionText}>{opt.label}</Text>
                  </TouchableOpacity>
                ))}
              </ScrollView>

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

          {secondaryUri ? (
            <View
              style={[
                styles.dualViewerOuter,
                { flexDirection: dualLayout === 'row' ? 'row' : 'column' },
              ]}
            >
              <View
                style={[
                  styles.dualPaneShell,
                  activePane === 1 && styles.dualPaneShellActive,
                ]}
              >
                <TouchableOpacity
                  style={styles.dualPaneTab}
                  onPress={() => setActivePane(1)}
                  activeOpacity={0.85}
                >
                  <MaterialIcons
                    name="description"
                    size={16}
                    color={activePane === 1 ? colors.primary : colors.onSurface}
                  />
                  <Text
                    style={styles.dualPaneTabText}
                    numberOfLines={1}
                  >
                    {documentName || 'Document 1'}
                  </Text>
                </TouchableOpacity>
                <View style={styles.dualPaneBody}>{renderDocumentPane(1)}</View>
              </View>
              <View
                style={[
                  styles.dualPaneShell,
                  activePane === 2 && styles.dualPaneShellActive,
                ]}
              >
                <TouchableOpacity
                  style={styles.dualPaneTab}
                  onPress={() => setActivePane(2)}
                  activeOpacity={0.85}
                >
                  <MaterialIcons
                    name="description"
                    size={16}
                    color={activePane === 2 ? colors.primary : colors.onSurface}
                  />
                  <Text
                    style={styles.dualPaneTabText}
                    numberOfLines={1}
                  >
                    {secondaryName || 'Document 2'}
                  </Text>
                </TouchableOpacity>
                <View style={styles.dualPaneBody}>{renderDocumentPane(2)}</View>
              </View>
            </View>
          ) : documentType === 'pdf' && viewerHtml ? (
            <WebView
              key={`${documentKey || documentUri}-${viewerInstanceId}`}
              ref={webViewRef}
              source={{ html: viewerHtml }}
              style={styles.webView}
              originWhitelist={['*']}
              javaScriptEnabled={true}
              domStorageEnabled={true}
              onMessage={handleWebViewMessagePrimary}
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
            {PdfConversionFormats.PDF_CONVERSION_FORMATS.map((formatConfig) => (
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
        visible={showOpenChoiceModal}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setShowOpenChoiceModal(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>Open document</Text>
            <Text style={styles.pageEditorHint}>
              Replace what you are viewing with another file, or open a second PDF alongside it.
              Two PDFs can be shown side by side or stacked; you can copy, move, or reorder pages
              between them.
            </Text>
            <TouchableOpacity
              style={styles.pageActionBtnPrimary}
              onPress={() => {
                setShowOpenChoiceModal(false);
                pickDocumentIntoSlot(1, { replaceAll: true });
              }}
            >
              <Text style={styles.pageActionBtnPrimaryText}>Replace with new file</Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={styles.pageActionBtnPrimary}
              onPress={() => {
                setShowOpenChoiceModal(false);
                pickDocumentIntoSlot(2);
              }}
            >
              <Text style={styles.pageActionBtnPrimaryText}>Add second PDF</Text>
            </TouchableOpacity>
            <View style={styles.modalButtons}>
              <TouchableOpacity
                style={styles.modalBtnCancel}
                onPress={() => setShowOpenChoiceModal(false)}
              >
                <Text style={styles.modalBtnCancelText}>Cancel</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>

      <Modal
        visible={showPageTransferModal}
        transparent={true}
        animationType="fade"
        onRequestClose={() => !pageTransferBusy && setShowPageTransferModal(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={[styles.modalContent, styles.pageTransferModalCard]}>
            <ScrollView
              keyboardShouldPersistTaps="handled"
              showsVerticalScrollIndicator={false}
            >
              <Text style={styles.modalTitle}>Merge pages</Text>
              <Text style={styles.pageEditorHint}>
                Copy or move one page from either PDF into the other. Page content is merged;
                ink and highlights on the new page start fresh. All other pages stay as they are.
              </Text>

              <Text style={styles.pageSelectionLabel}>From document</Text>
              <View style={styles.pageTransferRow}>
                <TouchableOpacity
                  style={[
                    styles.pageTransferChip,
                    pageTransferSourcePane === 1 && styles.pageTransferChipOn,
                  ]}
                  onPress={() => setPageTransferSourcePane(1)}
                >
                  <Text
                    style={[
                      styles.pageTransferChipText,
                      pageTransferSourcePane === 1 && styles.pageTransferChipTextOn,
                    ]}
                    numberOfLines={1}
                  >
                    1 — {shortDocLabel(documentName || 'First')}
                  </Text>
                </TouchableOpacity>
                <TouchableOpacity
                  style={[
                    styles.pageTransferChip,
                    pageTransferSourcePane === 2 && styles.pageTransferChipOn,
                  ]}
                  onPress={() => setPageTransferSourcePane(2)}
                >
                  <Text
                    style={[
                      styles.pageTransferChipText,
                      pageTransferSourcePane === 2 && styles.pageTransferChipTextOn,
                    ]}
                    numberOfLines={1}
                  >
                    2 — {shortDocLabel(secondaryName || 'Second')}
                  </Text>
                </TouchableOpacity>
              </View>

              <Text style={styles.pageSelectionLabel}>Into document</Text>
              <View style={styles.pageTransferRow}>
                <TouchableOpacity
                  style={[
                    styles.pageTransferChip,
                    pageTransferTargetPane === 1 && styles.pageTransferChipOn,
                  ]}
                  onPress={() => setPageTransferTargetPane(1)}
                >
                  <Text
                    style={[
                      styles.pageTransferChipText,
                      pageTransferTargetPane === 1 && styles.pageTransferChipTextOn,
                    ]}
                    numberOfLines={1}
                  >
                    1 — {shortDocLabel(documentName || 'First')}
                  </Text>
                </TouchableOpacity>
                <TouchableOpacity
                  style={[
                    styles.pageTransferChip,
                    pageTransferTargetPane === 2 && styles.pageTransferChipOn,
                  ]}
                  onPress={() => setPageTransferTargetPane(2)}
                >
                  <Text
                    style={[
                      styles.pageTransferChipText,
                      pageTransferTargetPane === 2 && styles.pageTransferChipTextOn,
                    ]}
                    numberOfLines={1}
                  >
                    2 — {shortDocLabel(secondaryName || 'Second')}
                  </Text>
                </TouchableOpacity>
              </View>

              <Text style={styles.pageSelectionLabel}>Source page #</Text>
              <TextInput
                style={styles.modalInput}
                value={pageTransferSourcePage}
                onChangeText={setPageTransferSourcePage}
                keyboardType="number-pad"
                placeholder="1"
                placeholderTextColor={withAlpha(colors.onSurface, 0.45)}
              />

              <Text style={styles.pageSelectionLabel}>
                Insert before page # (in target document)
              </Text>
              <Text style={styles.pageEditorHint}>
                Use 1 to insert at the beginning. Use one more than the last page number to append
                at the end.
              </Text>
              <TextInput
                style={styles.modalInput}
                value={pageTransferInsertBefore}
                onChangeText={setPageTransferInsertBefore}
                keyboardType="number-pad"
                placeholder="1"
                placeholderTextColor={withAlpha(colors.onSurface, 0.45)}
              />

              <Text style={styles.pageSelectionLabel}>Mode</Text>
              <View style={styles.pageTransferRow}>
                <TouchableOpacity
                  style={[
                    styles.pageTransferChip,
                    pageTransferMode === 'copy' && styles.pageTransferChipOn,
                  ]}
                  onPress={() => setPageTransferMode('copy')}
                >
                  <Text
                    style={[
                      styles.pageTransferChipText,
                      pageTransferMode === 'copy' && styles.pageTransferChipTextOn,
                    ]}
                  >
                    Copy
                  </Text>
                </TouchableOpacity>
                <TouchableOpacity
                  style={[
                    styles.pageTransferChip,
                    pageTransferMode === 'move' && styles.pageTransferChipOn,
                  ]}
                  onPress={() => setPageTransferMode('move')}
                >
                  <Text
                    style={[
                      styles.pageTransferChipText,
                      pageTransferMode === 'move' && styles.pageTransferChipTextOn,
                    ]}
                  >
                    Move (remove from source)
                  </Text>
                </TouchableOpacity>
              </View>
            </ScrollView>

            <View style={styles.modalButtons}>
              <TouchableOpacity
                style={styles.modalBtnCancel}
                disabled={pageTransferBusy}
                onPress={() => setShowPageTransferModal(false)}
              >
                <Text style={styles.modalBtnCancelText}>Cancel</Text>
              </TouchableOpacity>
              <TouchableOpacity
                style={[
                  styles.modalBtnSubmit,
                  pageTransferBusy && styles.actionBtnDisabled,
                ]}
                disabled={pageTransferBusy}
                onPress={executePdfPageTransfer}
              >
                <Text style={styles.modalBtnSubmitText}>
                  {pageTransferBusy ? 'Working…' : 'Merge pages'}
                </Text>
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
          <View style={[styles.modalContent, styles.pageEditorModalCard]}>
            <ScrollView
              style={styles.pageEditorModalScroll}
              contentContainerStyle={styles.pageEditorModalScrollContent}
              keyboardShouldPersistTaps="handled"
              showsVerticalScrollIndicator={false}
            >
              <Text style={styles.modalTitle}>Edit PDF Pages</Text>
              <Text style={styles.pageEditorHint}>
                Tap a page to select it. Long-press any page to open this panel
                for that page instantly.
              </Text>
              <Text style={styles.pageEditorMeta}>
                Current pages: {activeTotalPages}
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
              <Text style={styles.pageExportSectionTitle}>Download page as image</Text>
              <Text style={styles.pageExportSectionHint}>
                Exports the selected page including highlights, drawings, and notes.
              </Text>
              {PAGE_IMAGE_EXPORT_FORMATS.map((fmt) => (
                <TouchableOpacity
                  key={fmt.mimeType}
                  style={[
                    styles.pageActionBtnPrimary,
                    (pageEditorBusy || exportPageImageBusy) && styles.actionBtnDisabled,
                  ]}
                  disabled={pageEditorBusy || exportPageImageBusy}
                  onPress={() => exportPdfPageAsImage(fmt)}
                >
                  <View style={styles.pageExportFormatRow}>
                    <Text style={styles.pageActionBtnPrimaryText}>
                      {fmt.label}
                      {exportPageImageBusy ? '…' : ''}
                    </Text>
                    <Text style={styles.pageExportFormatSubtitle}>{fmt.subtitle}</Text>
                  </View>
                </TouchableOpacity>
              ))}
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
                    {pageEditorBusy || exportPageImageBusy ? 'Working...' : 'Close'}
                  </Text>
                </TouchableOpacity>
              </View>
            </ScrollView>
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
            {secondaryUri && documentType === 'pdf' && secondaryType === 'pdf' ? (
              <TouchableOpacity
                style={styles.overflowMenuRow}
                onPress={() => {
                  setShowPdfOverflowMenu(false);
                  setShowPageTransferModal(true);
                }}
              >
                <MaterialIcons name="swap-horiz" size={22} color={colors.onSurface} />
                <Text style={styles.overflowMenuRowText}>
                  Move or copy page between PDFs
                </Text>
              </TouchableOpacity>
            ) : null}
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

      <Modal
        visible={officeFontModalVisible}
        transparent={true}
        animationType="fade"
        onRequestClose={() => setOfficeFontModalVisible(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={[styles.modalContent, styles.officeFontModalCard]}>
            <Text style={styles.modalTitle}>Font family</Text>
            <ScrollView
              style={styles.officeFontModalScroll}
              keyboardShouldPersistTaps="handled"
            >
              {OFFICE_FONT_DROPDOWN_OPTIONS.map((opt) => (
                <TouchableOpacity
                  key={opt.value}
                  style={styles.officeFontModalRow}
                  onPress={() => {
                    applyOfficeStyleCommand('setFontFamily', opt.value);
                    setOfficeFontModalVisible(false);
                  }}
                >
                  <Text style={styles.officeFontModalRowText}>{opt.label}</Text>
                  <Text style={styles.officeFontModalRowMeta}>{opt.value}</Text>
                </TouchableOpacity>
              ))}
            </ScrollView>
            <TouchableOpacity
              style={styles.officeFontModalClose}
              onPress={() => setOfficeFontModalVisible(false)}
            >
              <Text style={styles.modalBtnCancelText}>Close</Text>
            </TouchableOpacity>
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

function EditableTextFileViewer({ uri, extension, onRegisterSave }) {
  const [content, setContent] = useState('');
  const [loadingText, setLoadingText] = useState(true);
  const originalRef = useRef('');

  const save = useCallback(async () => {
    try {
      await LegacyFileSystem.writeAsStringAsync(uri, content, {
        encoding: LegacyFileSystem.EncodingType.UTF8,
      });
      originalRef.current = content;
      return true;
    } catch (error) {
      Alert.alert('Save failed', String(error?.message || error));
      return false;
    }
  }, [uri, content]);

  useEffect(() => {
    onRegisterSave?.(save);
  }, [onRegisterSave, save]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoadingText(true);
      try {
        const text = await LegacyFileSystem.readAsStringAsync(uri, {
          encoding: LegacyFileSystem.EncodingType.UTF8,
        });
        if (!cancelled) {
          setContent(text);
          originalRef.current = text;
        }
      } catch (error) {
        if (!cancelled) {
          setContent('Error reading file: ' + error.message);
        }
      } finally {
        if (!cancelled) {
          setLoadingText(false);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [uri]);

  if (loadingText) {
    return (
      <View style={styles.textEditorLoading}>
        <ActivityIndicator size="large" color={colors.primary} />
      </View>
    );
  }

  return (
    <View style={styles.textEditorWrapper}>
      <TextInput
        style={styles.textEditorInput}
        multiline
        value={content}
        onChangeText={setContent}
        textAlignVertical="top"
        autoCapitalize="none"
        autoCorrect={false}
        spellCheck={false}
        scrollEnabled
        placeholder="File is empty. Start typing…"
        placeholderTextColor={withAlpha(colors.onSurface, 0.38)}
      />
    </View>
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
  layoutToggleGroup: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    marginRight: 2,
  },
  layoutToggleBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 5,
    paddingHorizontal: 10,
    paddingVertical: 7,
    borderRadius: radius.sm,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.85),
    backgroundColor: colors.surfaceContainerLowest,
    maxWidth: 120,
  },
  layoutToggleBtnActive: {
    borderColor: colors.primaryDim,
    backgroundColor: withAlpha(colors.primary, 0.12),
  },
  layoutToggleBtnLabel: {
    fontSize: 12,
    fontWeight: '600',
    color: withAlpha(colors.onSurface, 0.85),
  },
  layoutToggleBtnLabelActive: {
    color: colors.primaryDim,
  },
  mergePagesBarBtn: {
    width: 44,
    height: 44,
    borderRadius: radius.sm,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1.5,
    borderColor: colors.primaryDim,
    backgroundColor: withAlpha(colors.primary, 0.1),
  },
  dualPdfGuideCard: {
    marginHorizontal: spacing.sm,
    marginBottom: spacing.sm,
    padding: spacing.sm,
    borderRadius: radius.md,
    backgroundColor: colors.surfaceContainerLowest,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.65),
    ...Platform.select({
      ios: {
        shadowColor: colors.onSurface,
        shadowOffset: { width: 0, height: 1 },
        shadowOpacity: 0.06,
        shadowRadius: 3,
      },
      android: { elevation: 2 },
    }),
  },
  dualPdfGuideHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    marginBottom: 8,
  },
  dualPdfGuideTitle: {
    flex: 1,
    fontSize: 14,
    fontWeight: '700',
    color: colors.onSurface,
  },
  dualPdfGuideLine: {
    fontSize: 12,
    lineHeight: 18,
    color: withAlpha(colors.onSurface, 0.78),
    marginBottom: 6,
  },
  dualPdfGuideBold: {
    fontWeight: '700',
    color: colors.onSurface,
  },
  crossPdfBannerCard: {
    marginHorizontal: spacing.sm,
    marginBottom: spacing.sm,
    padding: spacing.sm,
    borderRadius: radius.md,
    backgroundColor: withAlpha(colors.primary, 0.07),
    borderWidth: 1,
    borderColor: withAlpha(colors.primaryDim, 0.35),
    ...Platform.select({
      ios: {
        shadowColor: colors.onSurface,
        shadowOffset: { width: 0, height: 2 },
        shadowOpacity: 0.08,
        shadowRadius: 4,
      },
      android: { elevation: 3 },
    }),
  },
  crossPdfBannerTopRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: 10,
    marginBottom: 10,
  },
  crossPdfBannerIconWrap: {
    width: 40,
    height: 40,
    borderRadius: 20,
    backgroundColor: withAlpha(colors.primary, 0.12),
    alignItems: 'center',
    justifyContent: 'center',
  },
  crossPdfBannerTextBlock: {
    flex: 1,
    minWidth: 140,
  },
  crossPdfBannerKicker: {
    fontSize: 10,
    fontWeight: '700',
    letterSpacing: 0.6,
    textTransform: 'uppercase',
    color: colors.primaryDim,
    marginBottom: 2,
  },
  crossPdfBannerTitle: {
    fontSize: 14,
    fontWeight: '700',
    color: colors.onSurface,
    marginBottom: 4,
  },
  crossPdfBannerSub: {
    fontSize: 12,
    lineHeight: 17,
    color: withAlpha(colors.onSurface, 0.72),
  },
  crossPdfBannerStepLabel: {
    fontSize: 11,
    fontWeight: '700',
    color: withAlpha(colors.onSurface, 0.88),
    marginBottom: 6,
  },
  crossPdfBannerModes: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: 8,
    marginBottom: 10,
  },
  crossPdfModeChip: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: radius.sm,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.85),
    backgroundColor: colors.surfaceContainerLowest,
  },
  crossPdfModeChipOn: {
    borderColor: colors.primaryDim,
    backgroundColor: withAlpha(colors.primary, 0.16),
  },
  crossPdfModeChipText: {
    fontSize: 13,
    fontWeight: '700',
    color: colors.onSurface,
  },
  crossPdfModeChipTextOn: {
    color: colors.primaryDim,
  },
  crossPdfBannerActions: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: 8,
  },
  crossPdfBannerBtnPrimary: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingHorizontal: 14,
    paddingVertical: 10,
    borderRadius: radius.sm,
    backgroundColor: colors.primaryDim,
  },
  crossPdfBannerBtnPrimaryText: {
    fontSize: 13,
    fontWeight: '700',
    color: colors.onPrimary,
  },
  crossPdfBannerBtnSecondary: {
    paddingHorizontal: 14,
    paddingVertical: 10,
    borderRadius: radius.sm,
    backgroundColor: colors.surfaceContainerLowest,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.75),
  },
  crossPdfBannerBtnSecondaryText: {
    fontSize: 13,
    fontWeight: '600',
    color: colors.primaryDim,
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
  dualViewerOuter: {
    flex: 1,
    minHeight: 0,
  },
  dualPaneShell: {
    flex: 1,
    minHeight: 120,
    minWidth: 0,
    borderWidth: StyleSheet.hairlineWidth,
    borderColor: withAlpha(colors.outlineVariant, 0.65),
  },
  dualPaneShellActive: {
    borderColor: colors.primaryDim,
    borderWidth: 2,
  },
  dualPaneTab: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingHorizontal: 10,
    paddingVertical: 8,
    backgroundColor: ui.shellElevated,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: colors.outlineVariant,
  },
  dualPaneTabText: {
    flex: 1,
    fontSize: 12,
    fontWeight: '600',
    color: colors.onSurface,
  },
  dualPaneBody: {
    flex: 1,
    minHeight: 0,
  },
  actionIconBtnActivePane: {
    borderColor: colors.primaryDim,
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
  textEditorWrapper: {
    flex: 1,
    minHeight: 120,
  },
  textEditorLoading: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 24,
  },
  textEditorInput: {
    flex: 1,
    margin: 12,
    padding: 14,
    borderRadius: radius.sm,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.55),
    backgroundColor: colors.surfaceContainerLowest,
    fontSize: 14,
    lineHeight: 22,
    color: colors.onSurface,
    fontFamily: Platform.OS === 'ios' ? 'Menlo' : 'monospace',
  },
  textFileToolbarHint: {
    fontSize: 12,
    lineHeight: 18,
    color: withAlpha(colors.onSurface, 0.72),
    paddingHorizontal: 4,
    marginBottom: 4,
  },
  officeFontDropdownBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 8,
    marginBottom: 10,
    paddingVertical: 10,
    paddingHorizontal: 14,
    borderRadius: radius.sm,
    backgroundColor: ui.shellElevated,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.5),
  },
  officeFontDropdownBtnText: {
    flex: 1,
    fontSize: 14,
    fontWeight: '700',
    color: colors.primaryDim,
  },
  officeSubheadingLabel: {
    fontSize: 11,
    fontWeight: '700',
    color: withAlpha(colors.onSurface, 0.58),
    textTransform: 'uppercase',
    letterSpacing: 0.7,
    marginBottom: 6,
    marginTop: 4,
  },
  officeFontModalCard: {
    maxWidth: 400,
    maxHeight: '72%',
  },
  officeFontModalScroll: {
    maxHeight: 360,
  },
  officeFontModalRow: {
    paddingVertical: 12,
    paddingHorizontal: 4,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: withAlpha(colors.outlineVariant, 0.5),
  },
  officeFontModalRowText: {
    fontSize: 16,
    fontWeight: '600',
    color: colors.onSurface,
  },
  officeFontModalRowMeta: {
    fontSize: 11,
    color: withAlpha(colors.onSurface, 0.5),
    marginTop: 4,
    fontFamily: Platform.OS === 'ios' ? 'Menlo' : 'monospace',
  },
  officeFontModalClose: {
    marginTop: 12,
    alignItems: 'center',
    paddingVertical: 8,
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
  pageEditorModalCard: {
    maxHeight: '88%',
  },
  pageTransferModalCard: {
    maxHeight: '88%',
    width: '92%',
    maxWidth: 420,
  },
  pageTransferRow: {
    flexDirection: 'row',
    gap: 8,
    marginBottom: 12,
    flexWrap: 'wrap',
  },
  pageTransferChip: {
    flex: 1,
    minWidth: 120,
    paddingVertical: 10,
    paddingHorizontal: 8,
    borderRadius: radius.sm,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.85),
    backgroundColor: colors.surfaceContainerLowest,
  },
  pageTransferChipOn: {
    borderColor: colors.primaryDim,
    backgroundColor: withAlpha(colors.primary, 0.12),
  },
  pageTransferChipText: {
    fontSize: 12,
    color: colors.onSurface,
    textAlign: 'center',
  },
  pageTransferChipTextOn: {
    fontWeight: '600',
    color: colors.primaryDim,
  },
  pageEditorModalScroll: {
    flexGrow: 0,
  },
  pageEditorModalScrollContent: {
    paddingBottom: 4,
  },
  pageExportSectionTitle: {
    color: colors.onSurface,
    fontSize: 14,
    fontWeight: '700',
    marginBottom: 4,
    marginTop: 4,
  },
  pageExportSectionHint: {
    color: withAlpha(colors.onSurface, 0.72),
    fontSize: 12,
    lineHeight: 18,
    marginBottom: 12,
  },
  pageExportFormatRow: {
    alignItems: 'center',
  },
  pageExportFormatSubtitle: {
    color: withAlpha(colors.onPrimary, 0.88),
    fontSize: 11,
    marginTop: 4,
    textAlign: 'center',
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

