import { useSyncExternalStore } from 'react';

/**
 * Fonts on the user's device.
 *
 * Two sources, merged:
 * 1. A quick probe of common fonts (Windows, macOS, Linux, Android, iOS, Arabic families):
 *    a font counts as installed when text drawn with it measures differently from every
 *    generic fallback. Needs no permission and runs once, on first use.
 * 2. The Local Font Access API (`queryLocalFonts`, Chrome/Edge desktop): the full list of
 *    every installed family. It needs the user's permission, so it runs only after a click,
 *    or by itself when permission was granted before. The list is remembered.
 */

/** Web fonts that ship with Paperline, available on every device. */
export const BUNDLED_FONTS = ['Inter', 'IBM Plex Sans Arabic'];

const CANDIDATES = [
  // Windows
  'Arial', 'Arial Black', 'Bahnschrift', 'Calibri', 'Calibri Light', 'Cambria', 'Cambria Math', 'Candara', 'Cascadia Code',
  'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Ebrima', 'Franklin Gothic Medium', 'Gabriola',
  'Gadugi', 'Georgia', 'Impact', 'Ink Free', 'Javanese Text', 'Leelawadee UI', 'Lucida Console', 'Lucida Sans Unicode',
  'Malgun Gothic', 'Microsoft Himalaya', 'Microsoft JhengHei', 'Microsoft New Tai Lue', 'Microsoft PhagsPa',
  'Microsoft Sans Serif', 'Microsoft Tai Le', 'Microsoft YaHei', 'Microsoft Yi Baiti', 'MingLiU-ExtB', 'Mongolian Baiti',
  'MS Gothic', 'MV Boli', 'Myanmar Text', 'Nirmala UI', 'Palatino Linotype', 'Segoe Print', 'Segoe Script', 'Segoe UI',
  'Segoe UI Emoji', 'Segoe UI Historic', 'Segoe UI Symbol', 'SimSun', 'Sitka Text', 'Sylfaen', 'Symbol', 'Tahoma',
  'Times New Roman', 'Trebuchet MS', 'Verdana', 'Webdings', 'Wingdings', 'Yu Gothic', 'Aptos', 'Aptos Display',
  'Book Antiqua', 'Bookman Old Style', 'Century Gothic', 'Garamond', 'Gill Sans MT', 'Lucida Bright', 'Lucida Handwriting',
  'Monotype Corsiva', 'Rockwell', 'Tw Cen MT', 'Agency FB', 'Baskerville Old Face', 'Bodoni MT', 'Brush Script MT',
  'Copperplate Gothic Bold', 'Edwardian Script ITC', 'Elephant', 'Felix Titling', 'Footlight MT Light', 'Forte',
  'Harrington', 'Lucida Calligraphy', 'Old English Text MT', 'Perpetua', 'Script MT Bold', 'Stencil',
  // Arabic (Windows / Office / macOS / Google)
  'Arabic Typesetting', 'Sakkal Majalla', 'Simplified Arabic', 'Simplified Arabic Fixed', 'Traditional Arabic',
  'Andalus', 'Aldhabi', 'Urdu Typesetting', 'Arial Unicode MS', 'Al Bayan', 'Al Nile', 'Al Tarikh', 'Baghdad',
  'Damascus', 'DecoType Naskh', 'Diwan Kufi', 'Diwan Thuluth', 'Farah', 'Farisi', 'Geeza Pro', 'KufiStandardGK',
  'Mishafi', 'Mishafi Gold', 'Muna', 'Nadeem', 'Noto Nastaliq Urdu', 'Noto Sans Arabic', 'Noto Naskh Arabic',
  'Noto Kufi Arabic', 'Sana', 'Waseem', 'Amiri', 'Cairo', 'Tajawal', 'Almarai', 'Changa', 'El Messiri', 'Harmattan',
  'Lateef', 'Scheherazade New', 'Reem Kufi', 'Aref Ruqaa', 'Markazi Text', 'Mada', 'Lalezar', 'Katibeh', 'Rakkas',
  'Kufam', 'Readex Pro', 'Vazirmatn', 'Dubai', 'Dubai Light', 'Dubai Medium', 'Frutiger LT Arabic', 'GE SS Two',
  'Hacen Tunisia', 'DIN Next LT Arabic', 'Droid Arabic Kufi', 'Droid Arabic Naskh', 'Janna LT', 'Hiragino Sans',
  // macOS / iOS
  'American Typewriter', 'Andale Mono', 'Apple Chancery', 'Apple Color Emoji', 'Apple SD Gothic Neo', 'Avenir',
  'Avenir Next', 'Avenir Next Condensed', 'Baskerville', 'Big Caslon', 'Bradley Hand', 'Chalkboard', 'Chalkboard SE',
  'Chalkduster', 'Charter', 'Cochin', 'Copperplate', 'Courier', 'Didot', 'Futura', 'Geneva', 'Gill Sans', 'Helvetica',
  'Helvetica Neue', 'Herculanum', 'Hoefler Text', 'Iowan Old Style', 'Lucida Grande', 'Luminari', 'Marker Felt',
  'Menlo', 'Monaco', 'Noteworthy', 'Optima', 'Palatino', 'Papyrus', 'Phosphate', 'PT Mono', 'PT Sans', 'PT Serif',
  'Rockwell', 'San Francisco', 'SF Pro', 'SF Pro Display', 'SF Pro Text', 'SF Mono', 'New York', 'Savoye LET',
  'SignPainter', 'Skia', 'Snell Roundhand', 'Superclarendon', 'Times', 'Trattatello', 'Zapfino', 'Kohinoor Devanagari',
  'PingFang SC', 'Hiragino Mincho ProN', 'STHeiti',
  // Linux / Android / ChromeOS / common downloads
  'Cantarell', 'DejaVu Sans', 'DejaVu Sans Mono', 'DejaVu Serif', 'Droid Sans', 'Droid Sans Mono', 'Droid Serif',
  'Fira Code', 'Fira Mono', 'Fira Sans', 'FreeMono', 'FreeSans', 'FreeSerif', 'Liberation Mono', 'Liberation Sans',
  'Liberation Serif', 'Noto Color Emoji', 'Noto Mono', 'Noto Sans', 'Noto Serif', 'Noto Sans Mono', 'Open Sans',
  'Roboto', 'Roboto Condensed', 'Roboto Mono', 'Roboto Slab', 'Source Code Pro', 'Source Sans Pro', 'Source Sans 3',
  'Source Serif Pro', 'Ubuntu', 'Ubuntu Mono', 'Ubuntu Condensed', 'Carlito', 'Caladea', 'Comfortaa', 'Lato',
  'Montserrat', 'Merriweather', 'Nunito', 'Oswald', 'Poppins', 'Raleway', 'Playfair Display', 'JetBrains Mono',
  'IBM Plex Sans', 'IBM Plex Serif', 'IBM Plex Mono', 'Google Sans', 'Product Sans', 'Josefin Sans', 'Quicksand',
  'Work Sans', 'Rubik', 'Manrope', 'DM Sans', 'Karla', 'Libre Baskerville', 'Libre Franklin', 'EB Garamond',
  'Crimson Text', 'Cormorant Garamond', 'Bitter', 'Arvo', 'Pacifico', 'Dancing Script', 'Great Vibes', 'Lobster',
  'Caveat', 'Indie Flower', 'Shadows Into Light', 'Permanent Marker', 'Bebas Neue', 'Anton', 'Abril Fatface',
];

const GENERIC = ['monospace', 'sans-serif', 'serif'];
const SAMPLE = 'mmmmmmmmmmlli1WQ@#أبجد هوز حطي';

function probe(): string[] {
  const ctx = document.createElement('canvas').getContext('2d');
  if (!ctx) return [];
  const base = GENERIC.map((g) => {
    ctx.font = `72px ${g}`;
    const m = ctx.measureText(SAMPLE);
    return [m.width, m.actualBoundingBoxAscent, m.actualBoundingBoxDescent].join('|');
  });
  const found: string[] = [];
  for (const f of CANDIDATES) {
    if (BUNDLED_FONTS.includes(f)) continue;
    // A missing font falls back to each generic and measures the same as all three.
    const differs = GENERIC.some((g, i) => {
      ctx.font = `72px "${f}", ${g}`;
      const m = ctx.measureText(SAMPLE);
      return [m.width, m.actualBoundingBoxAscent, m.actualBoundingBoxDescent].join('|') !== base[i];
    });
    if (differs) found.push(f);
  }
  return found;
}

// ---- store ----
interface State {
  device: string[];
  /** True once every installed family came from the Local Font Access API. */
  complete: boolean;
  recent: string[];
}

const STORE = 'paperline.fonts';
const RECENT = 'paperline.fonts.recent';

function load<T>(key: string, fallback: T): T {
  try {
    const v = localStorage.getItem(key);
    return v ? (JSON.parse(v) as T) : fallback;
  } catch {
    return fallback;
  }
}
function keep(key: string, v: unknown) {
  try {
    localStorage.setItem(key, JSON.stringify(v));
  } catch {
    /* storage unavailable */
  }
}

let state: State = { device: [], complete: false, recent: load<string[]>(RECENT, []) };
let started = false;
const subs = new Set<() => void>();
const set = (s: Partial<State>) => {
  state = { ...state, ...s };
  subs.forEach((f) => f());
};

const sortFonts = (a: string[]) => [...new Set(a)].sort((x, y) => x.localeCompare(y));

type LocalFontData = { family: string };
const api = () => (window as unknown as { queryLocalFonts?: () => Promise<LocalFontData[]> }).queryLocalFonts;

export const canListAllFonts = () => typeof window !== 'undefined' && !!api() && window.self === window.top;

/** Ask for the full list of installed fonts (shows the browser's permission prompt the first time). */
export async function listAllFonts(): Promise<boolean> {
  const q = api();
  if (!q) return false;
  try {
    const fonts = await q();
    const families = sortFonts(fonts.map((f) => f.family).filter((f) => f && !f.startsWith('.')));
    if (!families.length) return false;
    keep(STORE, families);
    set({ device: sortFonts([...families, ...state.device]), complete: true });
    return true;
  } catch {
    return false;
  }
}

function start() {
  if (started || typeof document === 'undefined') return;
  started = true;
  const saved = load<string[]>(STORE, []);
  // Probe after first paint so the UI never waits on it.
  setTimeout(() => {
    set({ device: sortFonts([...saved, ...probe()]), complete: saved.length > 0 });
    // Permission granted earlier: refresh the full list silently.
    const perms = navigator.permissions as unknown as { query?: (d: { name: string }) => Promise<{ state: string }> } | undefined;
    perms?.query?.({ name: 'local-fonts' }).then((p) => { if (p.state === 'granted') listAllFonts(); }).catch(() => null);
  }, 0);
}

export function rememberFont(f: string) {
  const recent = [f, ...state.recent.filter((x) => x !== f)].slice(0, 6);
  keep(RECENT, recent);
  set({ recent });
}

export function useDeviceFonts(): State {
  start();
  return useSyncExternalStore(
    (f) => {
      subs.add(f);
      return () => subs.delete(f);
    },
    () => state,
  );
}

/** CSS font-family value for a family name, with sensible fallbacks (incl. Arabic). */
export const familyCss = (f: string) => `"${f.replace(/"/g, '')}", "IBM Plex Sans Arabic", sans-serif`;
