/** Long-form copy for the PDF Converter screen — expandable “Read more” only. */

export type ConversionReadMoreSection = {
  title: string;
  lines: string[];
};

export const CONVERSION_READ_MORE_SECTIONS: ConversionReadMoreSection[] = [
  {
    title: 'On this screen (you selected a PDF)',
    lines: [
      'PDF → Word (.docx), PowerPoint (.pptx), plain text (.txt), Markdown (.md), or HTML (.html). Tap the format you need after opening the file.',
      'Conversion runs in the background; when it finishes you can share the new file or keep working.',
    ],
  },
  {
    title: 'Office & text → PDF (uses the editor)',
    lines: [
      'Word → PDF, Excel → PDF, PowerPoint → PDF: tap Edit document, then use Save and export as PDF (copy or replace).',
      'Plain text, Markdown, HTML, and similar: open with Edit document to review or change content, then save or export to PDF from there.',
    ],
  },
  {
    title: 'Cross-format (two steps)',
    lines: [
      'Word ↔ PowerPoint: there is no single built-in step. Practical path: export the source to PDF in the editor, open that PDF here, then convert to .docx or .pptx.',
      'PPT → Word: PowerPoint → PDF in the editor, then PDF → Word on this screen.',
      'Word → PPT: Word → PDF in the editor, then PDF → PowerPoint on this screen.',
      'Excel ↔ Word / PPT: same pattern — PDF is the bridge, or edit in the target format after exporting.',
    ],
  },
  {
    title: 'Images',
    lines: [
      'Photos → PDF: use the Images to PDF tool on the home screen to build a multi-page PDF from pictures.',
    ],
  },
];
