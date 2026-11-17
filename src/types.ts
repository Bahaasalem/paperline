import type { Workbook } from './editors/sheet/model';
import type { Deck } from './editors/slides/model';

export type DocKind = 'doc' | 'sheet' | 'slides';

export interface DocPageSetup {
  orientation: 'portrait' | 'landscape';
  margin: 'narrow' | 'normal' | 'wide';
}

export interface DocData {
  html: string;
  page: DocPageSetup;
}

export type OpenFile =
  | { id: number; kind: 'doc'; name: string; data: DocData }
  | { id: number; kind: 'sheet'; name: string; data: Workbook }
  | { id: number; kind: 'slides'; name: string; data: Deck };

export const KIND_ACCENT: Record<DocKind, string> = {
  doc: 'var(--doc)',
  sheet: 'var(--sheet)',
  slides: 'var(--slides)',
};
