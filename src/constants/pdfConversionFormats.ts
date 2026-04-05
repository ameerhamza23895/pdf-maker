export const PDF_CONVERSION_FORMATS = [
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
] as const;

export type PdfConversionFormatId = (typeof PDF_CONVERSION_FORMATS)[number]['id'];
