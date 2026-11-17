import type { Workbook } from './editors/sheet/model';
import type { Deck } from './editors/slides/model';
import type { PdfDoc } from './editors/pdf/model';

export type DocKind = 'doc' | 'sheet' | 'slides' | 'pdf';

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
  | { id: number; kind: 'slides'; name: string; data: Deck }
  | { id: number; kind: 'pdf'; name: string; data: PdfDoc };

export const KIND_ACCENT: Record<DocKind, string> = {
  doc: 'var(--doc)',
  sheet: 'var(--sheet)',
  slides: 'var(--slides)',
  pdf: 'var(--pdf)',
};
