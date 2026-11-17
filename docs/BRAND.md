# Paperline brand

Paperline is one calm, friendly place for every office file. The brand should feel
like good stationery: clean paper, confident ink, one bright line.

## The mark

A sheet of paper with a folded corner, two lines of text, and an amber **signature
line** that runs off the page. Paper + line. It sits on a rounded tile with the brand
gradient.

- Source: `src/ui/Logo.tsx` (`Logo`, `Wordmark`, `AppIcon`) and `public/favicon.svg`.
- Keep clear space of at least a quarter of the tile around it.
- Do not recolor the mark, add effects, or rotate it. On dark backgrounds use it as is.

## Wordmark

`paper` in ink, `line` in the brand gradient, always lowercase, Inter Bold with tight
letter spacing. The wordmark is always written left to right, also in the Arabic interface.

## Color

| Token | Hex | Use |
| --- | --- | --- |
| `--brand1` Indigo | `#5b4cf0` | Brand gradient start, links, greeting |
| `--brand2` Lagoon | `#0ea5a4` | Brand gradient end |
| `--spark` Amber | `#f5a524` | The signature line; tiny highlights only |
| `--ink` | `#1d2131` | Text, primary dark buttons |
| `--bg` Paper | `#f4f5fa` | App background |

Every app has its own color, used for its icon, tab, ribbon accent and buttons:

| App | Token | Hex |
| --- | --- | --- |
| Document | `--doc` | `#4f46e5` |
| Spreadsheet | `--sheet` | `#0e9f6e` |
| Presentation | `--slides` | `#e0436b` |
| PDF | `--pdf` | `#dc6a1f` |

Inside an editor, `--accent` is set to that app's color, so components only use `--accent`.

## App icons

One family: the app's color tile, the same folded sheet, and one simple mark
(lines, grid, chart, signature). New apps (notebook, database, diagram…) follow the same
recipe with a new color and mark.

## Type

- Interface: **Inter** for Latin, **IBM Plex Sans Arabic** for Arabic.
- Documents use the fonts on the user's device, plus the two above.
- Signatures: Dancing Script, Great Vibes and Aref Ruqaa.

## Voice

Plain, warm and short. Say what a button does ("Fill form", "Split into files"), not what
it is technically. Arabic copy is written, not translated word for word.

## Shape and motion

- Corners: 9px for controls, 12–14px for menus and small cards, 18–20px for large cards.
- Soft, low shadows; cards lift 2–3px on hover.
- Motion is short (150–250ms) and never blocks work.
