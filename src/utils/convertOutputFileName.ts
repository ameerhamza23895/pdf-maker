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

export function buildConvertedFileName(documentName: string, extension: string) {
  return `${sanitizeFileSegment(documentName)}_converted_${buildTimestampToken()}.${extension}`;
}
