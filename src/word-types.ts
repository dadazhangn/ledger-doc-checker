export type WordFieldKind = 'projectName' | 'softwareVersion' | 'documentVersion';
export interface WordSegment {
  id: string;
  location: string;
  text: string;
  notices: string[];
}
export interface WordSource {
  id: string;
  name: string;
  sha256: string;
  format: 'doc' | 'docx';
  segments: WordSegment[];
  notices: string[];
}
export interface WordBatch {
  documents: WordSource[];
  files: { id: string; name: string; sha256: string }[];
  errors: { id: string; name: string; message: string }[];
}
export interface WordSelection {
  documentId: string;
  category: WordFieldKind;
  segmentId: string;
  value: string;
  subject: string;
  stage: string;
}
export interface WordRecord extends WordSelection {
  file: string;
  sha256: string;
  location: string;
  notices: string[];
}
export interface WordDifferenceGroup {
  category: WordFieldKind;
  subject: string;
  records: WordRecord[];
}
export interface WordReport {
  schemaVersion: 1;
  engineVersion: string;
  mode: 'word-comparison';
  judgment: 'human-review';
  selectionHash: string;
  files: WordBatch['files'];
  errors: WordBatch['errors'];
  records: WordRecord[];
  differences: WordDifferenceGroup[];
  uncompared: WordRecord[];
  coverage: { documentId: string; file: string; selected: WordFieldKind[]; unselected: WordFieldKind[] }[];
}
