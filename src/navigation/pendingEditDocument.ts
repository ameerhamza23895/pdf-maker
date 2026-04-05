export type PendingEditDocument = {
  uri: string;
  name: string;
};

let pending: PendingEditDocument | null = null;

export function setPendingEditDocument(doc: PendingEditDocument) {
  pending = doc;
}

export function consumePendingEditDocument(): PendingEditDocument | null {
  const next = pending;
  pending = null;
  return next;
}
