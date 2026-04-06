// @ts-nocheck
import MaterialIcons from '@expo/vector-icons/MaterialIcons';
import * as DocumentPicker from 'expo-document-picker';
import { Paths } from 'expo-file-system';
import * as LegacyFileSystem from 'expo-file-system/legacy';
import { Stack, useRouter } from 'expo-router';
import * as Sharing from 'expo-sharing';
import { useCallback, useRef, useState } from 'react';
import {
  ActivityIndicator,
  Alert,
  Modal,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  TouchableOpacity,
  View,
} from 'react-native';
import { WebView } from 'react-native-webview';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { CONVERSION_READ_MORE_SECTIONS } from '@/src/constants/conversionReadMore';
import { DOCUMENT_PICKER_TYPES } from '@/src/constants/documentPicker';
import { PDF_CONVERSION_FORMATS } from '@/src/constants/pdfConversionFormats';
import { recordSavedFile } from '@/src/db/savedFileHistory';
import { setPendingEditDocument } from '@/src/navigation/pendingEditDocument';
import { electricCuratorTheme, withAlpha } from '@/src/theme/electric-curator';
import { buildConvertedFileName } from '@/src/utils/convertOutputFileName';
import { getPdfConverterHtml } from '@/src/utils/pdfConverterHtml';

const { colors, spacing, radius, typography } = electricCuratorTheme;

const OFFICE_EXTS = new Set(['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']);
const IMAGE_EXTS = new Set([
  'jpg',
  'jpeg',
  'png',
  'gif',
  'webp',
  'bmp',
  'heic',
  'heif',
  'tif',
  'tiff',
]);

const PDF_TO_OFFICE_FORMATS = PDF_CONVERSION_FORMATS.filter((f) =>
  ['docx', 'pptx'].includes(f.id)
);
const PDF_TO_TEXT_WEB_FORMATS = PDF_CONVERSION_FORMATS.filter((f) =>
  ['txt', 'md', 'html'].includes(f.id)
);

function extensionOf(name: string) {
  const parts = name.split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : '';
}

function pickedCategory(ext: string) {
  if (ext === 'pdf') {
    return 'pdf';
  }
  if (OFFICE_EXTS.has(ext)) {
    return 'office';
  }
  if (IMAGE_EXTS.has(ext)) {
    return 'image';
  }
  return 'other';
}

export default function ConvertPdfPage() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const [picked, setPicked] = useState(null);
  const [pdfBase64, setPdfBase64] = useState(null);
  const [picking, setPicking] = useState(false);
  const [readMoreOpen, setReadMoreOpen] = useState(false);
  const [conversionTask, setConversionTask] = useState(null);
  const [conversionStatus, setConversionStatus] = useState('');
  const conversionChunksRef = useRef({
    fileName: '',
    mimeType: '',
    chunks: [],
  });

  const resetConversionState = useCallback(() => {
    conversionChunksRef.current = {
      fileName: '',
      mimeType: '',
      chunks: [],
    };
    setConversionTask(null);
    setConversionStatus('');
  }, []);

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
              source: 'convert_pdf_screen',
            });

            const canShare = await Sharing.isAvailableAsync();
            resetConversionState();

            Alert.alert('Document converted', `Saved in app storage as ${fileName}.`, [
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
              'Conversion failed',
              data.message ||
                'The document could not be converted. Check your connection and try again.'
            );
            break;
        }
      } catch (error) {
        resetConversionState();
        Alert.alert('Conversion failed', error.message);
      }
    },
    [resetConversionState]
  );

  const pickDocument = useCallback(async () => {
    try {
      setPicking(true);
      const result = await DocumentPicker.getDocumentAsync({
        type: DOCUMENT_PICKER_TYPES,
        copyToCacheDirectory: true,
      });
      if (result.canceled) {
        return;
      }
      const file = result.assets[0];
      const ext = extensionOf(file.name);
      setPicked({
        uri: file.uri,
        name: file.name,
        ext,
      });
      if (ext === 'pdf') {
        const b64 = await LegacyFileSystem.readAsStringAsync(file.uri, {
          encoding: LegacyFileSystem.EncodingType.Base64,
        });
        setPdfBase64(b64);
      } else {
        setPdfBase64(null);
      }
    } catch (e) {
      Alert.alert('Error', e.message || 'Could not open the file.');
    } finally {
      setPicking(false);
    }
  }, []);

  const startPdfToFormat = useCallback(
    async (formatId) => {
      if (!pdfBase64 || !picked) {
        return;
      }
      const formatConfig = PDF_CONVERSION_FORMATS.find((f) => f.id === formatId);
      if (!formatConfig) {
        return;
      }
      try {
        setConversionStatus('Preparing…');
        const outputFileName = buildConvertedFileName(picked.name, formatConfig.extension);
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
          html: getPdfConverterHtml(pdfBase64, {
            format: formatId,
            fileName: outputFileName,
            mimeType: formatConfig.mimeType,
          }),
        });
      } catch (error) {
        resetConversionState();
        Alert.alert('Conversion failed', error.message);
      }
    },
    [pdfBase64, picked, resetConversionState]
  );

  const openInEditor = useCallback(() => {
    if (!picked) {
      return;
    }
    setPendingEditDocument({ uri: picked.uri, name: picked.name });
    router.push('/edit-pdf');
  }, [picked, router]);

  const openImagesToPdf = useCallback(() => {
    router.push('/convert-images');
  }, [router]);

  const category = picked ? pickedCategory(picked.ext) : null;
  const isPdf = category === 'pdf';
  const isOffice = category === 'office';
  const isImage = category === 'image';

  return (
    <>
      <Stack.Screen
        options={{
          title: 'PDF Converter',
          headerShown: true,
          headerStyle: { backgroundColor: colors.surfaceContainerLow },
          headerTintColor: colors.onSurface,
          headerTitleStyle: { fontWeight: '700' },
          headerBackTitle: 'Back',
        }}
      />
      <ScrollView
        style={styles.scroll}
        contentContainerStyle={[
          styles.scrollContent,
          { paddingBottom: insets.bottom + spacing.xl },
        ]}
        keyboardShouldPersistTaps="handled"
      >
        <View style={styles.readerShell}>
          <View style={styles.readerBookRule} />
          <Text style={styles.docKicker}>Read-only guide</Text>
          <Text style={styles.docTitle}>PDF Converter</Text>
          <Text style={styles.docBody}>
            This page explains how conversion works. It is not the PDF editor — there are no
            annotation tools or page editing here. Read the summary below, then use Actions to pick
            a file. To draw, merge pages, or export Office files to PDF, open Edit document after
            you choose a file.
          </Text>
        </View>

        <Text style={styles.actionsSectionLabel}>Actions</Text>

        <TouchableOpacity
          style={styles.primaryBtn}
          onPress={pickDocument}
          disabled={picking}
          accessibilityLabel="Open document"
        >
          {picking ? (
            <ActivityIndicator color={colors.onPrimary} />
          ) : (
            <>
              <MaterialIcons name="folder-open" size={22} color={colors.onPrimary} />
              <Text style={styles.primaryBtnText}>Open document</Text>
            </>
          )}
        </TouchableOpacity>

        <Text style={styles.hint}>
          Same file types as the editor picker: PDF, Office, images, text, Markdown, HTML, and more.
        </Text>

        {picked ? (
          <View style={styles.card}>
            <View style={styles.cardHeader}>
              <MaterialIcons name="insert-drive-file" size={24} color={colors.primaryDim} />
              <View style={styles.cardHeaderText}>
                <Text style={styles.cardTitle} numberOfLines={2}>
                  {picked.name}
                </Text>
                <Text style={styles.cardMeta}>
                  {isPdf && 'PDF — choose an output format below.'}
                  {isOffice &&
                    'Office file — export to PDF and other options are in the editor.'}
                  {isImage && 'Image — build a PDF from pictures, or open in the editor.'}
                  {!isPdf && !isOffice && !isImage && 'Open in the editor to save or print to PDF.'}
                </Text>
              </View>
            </View>

            {isPdf ? (
              <>
                <View style={styles.section}>
                  <Text style={styles.sectionTitle}>Microsoft-style</Text>
                  <Text style={styles.sectionHint}>PDF → Word or PowerPoint</Text>
                  <View style={styles.formatGrid}>
                    {PDF_TO_OFFICE_FORMATS.map((fmt) => (
                      <TouchableOpacity
                        key={fmt.id}
                        style={styles.formatChip}
                        onPress={() => startPdfToFormat(fmt.id)}
                        disabled={!!conversionTask}
                      >
                        <Text style={styles.formatChipText}>{fmt.label}</Text>
                      </TouchableOpacity>
                    ))}
                  </View>
                </View>
                <View style={styles.section}>
                  <Text style={styles.sectionTitle}>Text & web</Text>
                  <Text style={styles.sectionHint}>PDF → plain text, Markdown, or HTML</Text>
                  <View style={styles.formatGrid}>
                    {PDF_TO_TEXT_WEB_FORMATS.map((fmt) => (
                      <TouchableOpacity
                        key={fmt.id}
                        style={styles.formatChip}
                        onPress={() => startPdfToFormat(fmt.id)}
                        disabled={!!conversionTask}
                      >
                        <Text style={styles.formatChipText}>{fmt.label}</Text>
                      </TouchableOpacity>
                    ))}
                  </View>
                </View>
                <Text style={styles.convertFootnote}>
                  Editing, annotations, and page tools are not on this screen — use Edit document
                  when you need them.
                </Text>
              </>
            ) : null}

            {isOffice ? (
              <View style={styles.infoPanel}>
                <MaterialIcons name="info-outline" size={20} color={colors.primaryDim} />
                <Text style={styles.infoPanelText}>
                  Word, Excel, and PowerPoint → PDF: open the editor, then Save and export as PDF.
                  To change content or add notes, use Edit document.
                </Text>
              </View>
            ) : null}

            {isImage ? (
              <View style={styles.infoPanel}>
                <MaterialIcons name="photo-library" size={20} color={colors.primaryDim} />
                <Text style={styles.infoPanelText}>
                  For several photos in one PDF, use Images to PDF. One image can also open in the
                  editor for a single-page PDF.
                </Text>
              </View>
            ) : null}

            {!isPdf && !isOffice && !isImage ? (
              <View style={styles.infoPanel}>
                <MaterialIcons name="description" size={20} color={colors.primaryDim} />
                <Text style={styles.infoPanelText}>
                  This file type is converted or exported to PDF from the editor (Save / print to
                  PDF). Tap Edit document to continue.
                </Text>
              </View>
            ) : null}

            {isImage ? (
              <TouchableOpacity style={styles.secondaryBtn} onPress={openImagesToPdf}>
                <MaterialIcons name="collections" size={20} color={colors.primaryDim} />
                <Text style={styles.secondaryBtnText}>Images to PDF</Text>
                <MaterialIcons
                  name="chevron-right"
                  size={22}
                  color={withAlpha(colors.onSurface, 0.45)}
                />
              </TouchableOpacity>
            ) : null}

            <TouchableOpacity
              style={[styles.editDocBtn, isPdf && styles.editDocBtnMuted]}
              onPress={openInEditor}
              accessibilityLabel="Edit document"
              accessibilityHint="Opens the full PDF editor with this file"
            >
              <MaterialIcons name="edit" size={22} color={colors.primaryDim} />
              <View style={styles.editDocBtnTextCol}>
                <Text style={styles.editDocBtnTitle}>Edit document</Text>
                <Text style={styles.editDocBtnSub}>
                  {isPdf
                    ? 'Annotations, page order, merge, Office export…'
                    : 'Save as PDF, edit content, annotate…'}
                </Text>
              </View>
              <MaterialIcons
                name="chevron-right"
                size={22}
                color={withAlpha(colors.onSurface, 0.45)}
              />
            </TouchableOpacity>
          </View>
        ) : null}

        <Pressable
          style={styles.readMoreHeader}
          onPress={() => setReadMoreOpen((v) => !v)}
          accessibilityRole="button"
          accessibilityState={{ expanded: readMoreOpen }}
        >
          <MaterialIcons
            name={readMoreOpen ? 'expand-less' : 'expand-more'}
            size={22}
            color={colors.primaryDim}
          />
          <Text style={styles.readMoreTitle}>Read more — all conversion paths</Text>
        </Pressable>

        {readMoreOpen ? (
          <View style={styles.readerArticle}>
            <Text style={styles.readerArticleLabel}>Documentation</Text>
            <View style={styles.readMoreBody}>
              {CONVERSION_READ_MORE_SECTIONS.map((section) => (
                <View key={section.title} style={styles.readMoreSection}>
                  <Text style={styles.readMoreSectionTitle}>{section.title}</Text>
                  {section.lines.map((line, lineIndex) => (
                    <Text key={`${section.title}-${lineIndex}`} style={styles.readMorePara}>
                      {line}
                    </Text>
                  ))}
                </View>
              ))}
              <Text style={styles.readMoreParaMuted}>
                The full editor (opened via Edit document) is where you merge PDFs, reorder pages,
                annotate, and export Office files to PDF.
              </Text>
            </View>
          </View>
        ) : null}
      </ScrollView>

      <Modal visible={!!conversionTask} transparent animationType="fade" onRequestClose={() => {}}>
        <View style={styles.modalOverlay}>
          <View style={styles.modalCard}>
            <Text style={styles.modalTitle}>Converting</Text>
            <Text style={styles.modalHint}>{conversionTask?.formatLabel || 'Preparing…'}</Text>
            <ActivityIndicator size="large" color={colors.primaryDim} style={styles.loader} />
            <Text style={styles.modalStatus}>{conversionStatus || 'Preparing conversion…'}</Text>
            {conversionTask ? (
              <View style={styles.hiddenWebView}>
                <WebView
                  key={conversionTask.id}
                  source={{ html: conversionTask.html }}
                  onMessage={handleConverterMessage}
                  javaScriptEnabled
                  domStorageEnabled
                  originWhitelist={['*']}
                  mixedContentMode="always"
                />
              </View>
            ) : null}
          </View>
        </View>
      </Modal>
    </>
  );
}

const styles = StyleSheet.create({
  scroll: {
    flex: 1,
    backgroundColor: colors.surface,
  },
  scrollContent: {
    paddingHorizontal: spacing.md,
    paddingTop: spacing.md,
    gap: spacing.md,
  },
  readerShell: {
    borderRadius: radius.md,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.75),
    backgroundColor: colors.surfaceContainerLow,
    padding: spacing.md,
    gap: spacing.sm,
  },
  readerBookRule: {
    alignSelf: 'flex-start',
    width: 40,
    height: 3,
    borderRadius: 2,
    backgroundColor: withAlpha(colors.primaryDim, 0.45),
    marginBottom: spacing.xs,
  },
  docKicker: {
    fontSize: 11,
    fontWeight: '700',
    letterSpacing: 1.2,
    color: withAlpha(colors.onSurface, 0.55),
    textTransform: 'uppercase',
  },
  docTitle: {
    fontSize: 20,
    fontWeight: '700',
    color: colors.onSurface,
    lineHeight: 26,
  },
  docBody: {
    fontSize: 14,
    lineHeight: 22,
    color: withAlpha(colors.onSurface, 0.82),
  },
  actionsSectionLabel: {
    fontSize: 12,
    fontWeight: '700',
    letterSpacing: 0.6,
    color: withAlpha(colors.onSurface, 0.55),
    textTransform: 'uppercase',
    marginBottom: -4,
  },
  primaryBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 10,
    backgroundColor: colors.primaryDim,
    paddingVertical: 14,
    borderRadius: radius.md,
    minHeight: 52,
  },
  primaryBtnText: {
    ...typography.titleSm,
    color: colors.onPrimary,
    fontWeight: '700',
  },
  hint: {
    ...typography.bodyMd,
    fontSize: 12,
    color: withAlpha(colors.onSurface, 0.62),
    lineHeight: 18,
  },
  card: {
    backgroundColor: colors.surfaceContainerLowest,
    borderRadius: radius.md,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.65),
    padding: spacing.md,
    gap: spacing.md,
  },
  cardHeader: {
    flexDirection: 'row',
    gap: spacing.sm,
    alignItems: 'flex-start',
  },
  cardHeaderText: {
    flex: 1,
    minWidth: 0,
  },
  cardTitle: {
    ...typography.titleSm,
    color: colors.onSurface,
    fontWeight: '700',
  },
  cardMeta: {
    ...typography.bodyMd,
    fontSize: 12,
    color: withAlpha(colors.onSurface, 0.7),
    marginTop: 4,
  },
  section: {
    gap: spacing.sm,
  },
  sectionTitle: {
    ...typography.labelMd,
    color: colors.onSurface,
    fontWeight: '700',
  },
  sectionHint: {
    ...typography.bodyMd,
    fontSize: 12,
    color: withAlpha(colors.onSurface, 0.62),
    marginBottom: spacing.xs,
    marginTop: -4,
  },
  formatGrid: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: 8,
  },
  formatChip: {
    paddingHorizontal: 12,
    paddingVertical: 10,
    borderRadius: radius.sm,
    backgroundColor: withAlpha(colors.primary, 0.12),
    borderWidth: 1,
    borderColor: withAlpha(colors.primaryDim, 0.35),
  },
  formatChipText: {
    fontSize: 13,
    fontWeight: '600',
    color: colors.primaryDim,
  },
  convertFootnote: {
    ...typography.bodyMd,
    fontSize: 12,
    lineHeight: 18,
    color: withAlpha(colors.onSurface, 0.58),
    fontStyle: 'italic',
  },
  infoPanel: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: 10,
    padding: spacing.sm,
    borderRadius: radius.sm,
    backgroundColor: withAlpha(colors.primary, 0.07),
    borderWidth: 1,
    borderColor: withAlpha(colors.primaryDim, 0.22),
  },
  infoPanelText: {
    flex: 1,
    fontSize: 13,
    lineHeight: 20,
    color: withAlpha(colors.onSurface, 0.85),
  },
  secondaryBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    paddingVertical: 12,
    paddingHorizontal: 12,
    borderRadius: radius.sm,
    backgroundColor: colors.surfaceContainerLow,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.8),
  },
  secondaryBtnText: {
    flex: 1,
    fontSize: 14,
    fontWeight: '600',
    color: colors.primaryDim,
  },
  editDocBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 10,
    paddingVertical: 14,
    paddingHorizontal: 14,
    borderRadius: radius.md,
    backgroundColor: withAlpha(colors.primary, 0.1),
    borderWidth: 1.5,
    borderColor: colors.primaryDim,
  },
  editDocBtnMuted: {
    backgroundColor: colors.surfaceContainerLow,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.85),
  },
  editDocBtnTextCol: {
    flex: 1,
    minWidth: 0,
  },
  editDocBtnTitle: {
    fontSize: 15,
    fontWeight: '700',
    color: colors.primaryDim,
  },
  editDocBtnSub: {
    fontSize: 12,
    lineHeight: 17,
    color: withAlpha(colors.onSurface, 0.68),
    marginTop: 2,
  },
  readMoreHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6,
    paddingVertical: spacing.sm,
  },
  readMoreTitle: {
    ...typography.labelMd,
    color: colors.primaryDim,
    fontWeight: '700',
  },
  readerArticle: {
    borderRadius: radius.md,
    borderWidth: 1,
    borderColor: withAlpha(colors.outlineVariant, 0.6),
    backgroundColor: withAlpha(colors.surfaceContainerLowest, 0.98),
    padding: spacing.md,
    gap: spacing.sm,
  },
  readerArticleLabel: {
    fontSize: 11,
    fontWeight: '700',
    letterSpacing: 1,
    color: withAlpha(colors.primaryDim, 0.95),
    textTransform: 'uppercase',
  },
  readMoreBody: {
    paddingBottom: spacing.xs,
    gap: spacing.md,
  },
  readMoreSection: {
    gap: spacing.xs,
  },
  readMoreSectionTitle: {
    fontSize: 13,
    fontWeight: '700',
    color: colors.onSurface,
    marginBottom: 4,
  },
  readMorePara: {
    ...typography.bodyMd,
    fontSize: 13,
    lineHeight: 20,
    color: withAlpha(colors.onSurface, 0.85),
  },
  readMoreParaMuted: {
    ...typography.bodyMd,
    fontSize: 12,
    lineHeight: 18,
    color: withAlpha(colors.onSurface, 0.58),
    fontStyle: 'italic',
  },
  modalOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0,0,0,0.45)',
    alignItems: 'center',
    justifyContent: 'center',
    padding: spacing.lg,
  },
  modalCard: {
    width: '100%',
    maxWidth: 400,
    backgroundColor: colors.surfaceContainerLowest,
    borderRadius: radius.md,
    padding: spacing.lg,
  },
  modalTitle: {
    fontSize: 18,
    fontWeight: '700',
    color: colors.onSurface,
    marginBottom: 6,
  },
  modalHint: {
    fontSize: 14,
    color: withAlpha(colors.onSurface, 0.72),
    marginBottom: 12,
  },
  loader: {
    marginVertical: 12,
  },
  modalStatus: {
    fontSize: 13,
    textAlign: 'center',
    color: withAlpha(colors.onSurface, 0.74),
    lineHeight: 19,
  },
  hiddenWebView: {
    width: 1,
    height: 1,
    opacity: 0,
    overflow: 'hidden',
  },
});
