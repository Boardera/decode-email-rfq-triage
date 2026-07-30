// Shared filename sanitizer for every generated report and artifact, so
// the files on disk and the attachment names listed in the reply email
// can never disagree.
export function sanitizeName(name: string): string {
  return name.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 60) || 'project';
}
