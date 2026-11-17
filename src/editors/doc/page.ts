import type { DocPageSetup } from '../../types';

export const PAGE = { w: 794, h: 1123 }; // A4 at 96 dpi
export const MARGIN_PX: Record<DocPageSetup['margin'], number> = { narrow: 48, normal: 96, wide: 144 };
