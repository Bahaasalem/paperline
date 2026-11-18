/**
 * The few things that differ between the browser, the desktop apps (Tauri) and the
 * phone apps (Capacitor): saving a file, printing, a confirm box, and files the
 * system asks us to open ("Open with Paperline", double-click).
 */

type Shell = 'web' | 'desktop' | 'mobile';

const w = window as unknown as {
  __TAURI_INTERNALS__?: unknown;
  Capacitor?: { isNativePlatform?: () => boolean; getPlatform?: () => string };
};

export const shell: Shell = w.__TAURI_INTERNALS__ ? 'desktop' : w.Capacitor?.isNativePlatform?.() ? 'mobile' : 'web';
const isMac = /Mac/i.test(navigator.platform || navigator.userAgent);

const tauriCore = () => import('@tauri-apps/api/core');

/** Save through the system's own dialog (desktop), share sheet (phones) or a download (web). Returns false if cancelled. */
export async function nativeSave(blob: Blob, filename: string): Promise<boolean | null> {
  if (shell === 'desktop') {
    const { save } = await import('@tauri-apps/plugin-dialog');
    const ext = filename.split('.').pop() ?? '';
    const path = await save({ defaultPath: filename, filters: ext ? [{ name: ext.toUpperCase(), extensions: [ext] }] : [] });
    if (!path) return false;
    const { invoke } = await tauriCore();
    await invoke('write_file', new Uint8Array(await blob.arrayBuffer()), { headers: { path: encodeURIComponent(path) } });
    return true;
  }
  if (shell === 'mobile') {
    await shareFile(blob, filename);
    return true;
  }
  return null; // let the browser code handle it
}

/** Phones: write to the app's cache and open the share sheet (Save to Files, Drive, Mail, Print…). */
async function shareFile(blob: Blob, filename: string) {
  const [{ Filesystem, Directory }, { Share }] = await Promise.all([import('@capacitor/filesystem'), import('@capacitor/share')]);
  const data = await blobToBase64(blob);
  const res = await Filesystem.writeFile({ path: filename, data, directory: Directory.Cache });
  try {
    await Share.share({ title: filename, files: [res.uri] });
  } catch (e) {
    // closing the share sheet is not an error
    if (!/cancel/i.test(String((e as Error)?.message ?? e))) throw e;
  }
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(',')[1] ?? '');
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}

/**
 * Where window.print() is not available (macOS and phone web views), print a PDF instead:
 * on the Mac it opens in Preview, on phones in the share sheet, which has Print.
 */
export const printsViaPdf = shell === 'mobile' || (shell === 'desktop' && isMac);

export async function printPdfBlob(blob: Blob, filename: string) {
  if (shell === 'desktop') {
    const { invoke } = await tauriCore();
    await invoke('open_for_print', new Uint8Array(await blob.arrayBuffer()), { headers: { name: encodeURIComponent(filename) } });
    return;
  }
  await shareFile(blob, filename);
}

/** A yes/no question. Uses the system dialog in the desktop apps. */
export async function ask(message: string): Promise<boolean> {
  if (shell === 'desktop') {
    const { confirm } = await import('@tauri-apps/plugin-dialog');
    return confirm(message, { title: 'Paperline', kind: 'warning' });
  }
  return window.confirm(message);
}

/**
 * Files the system asked us to open: at launch (double-click, "Open with") and later,
 * while the app runs. Calls onFile for each one.
 */
/** Desktop: web links open in the system browser instead of replacing the app. */
if (shell === 'desktop') {
  document.addEventListener('click', (e) => {
    const a = (e.target as HTMLElement | null)?.closest?.('a[href]') as HTMLAnchorElement | null;
    if (!a || a.hasAttribute('download') || !/^(https?:|mailto:)/i.test(a.href)) return;
    e.preventDefault();
    import('@tauri-apps/plugin-opener').then(({ openUrl }) => openUrl(a.href)).catch(console.error);
  }, true);
}

export async function listenForOpenedFiles(onFile: (f: File) => void) {
  if (shell !== 'desktop') return () => {};
  const { invoke } = await tauriCore();
  const { listen } = await import('@tauri-apps/api/event');
  const load = async (paths: string[]) => {
    for (const path of paths) {
      try {
        const buf = await invoke<ArrayBuffer>('read_file', { path });
        const name = path.split(/[\\/]/).pop() || 'file';
        onFile(new File([buf], name));
      } catch (e) {
        console.error('open failed', path, e);
      }
    }
  };
  const un = await listen<string[]>('open-files', (e) => load(e.payload));
  await load(await invoke<string[]>('take_opened_files'));
  return un;
}
