/** MIME types for the system file picker; includes a catch-all type so JSON and code files show on all devices. */
export const DOCUMENT_PICKER_TYPES = [
  'application/pdf',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'text/plain',
  'text/markdown',
  'text/html',
  'text/css',
  'text/javascript',
  'application/javascript',
  'application/json',
  'application/xml',
  'text/xml',
  'text/yaml',
  'application/x-yaml',
  'application/octet-stream',
  'image/*',
  '*/*',
] as const;

export const OFFICE_PREVIEW_TYPES = ['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx'] as const;
