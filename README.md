# Outlook Sticky Notes

A Chrome extension that lets you pin private notes and action items directly to the top of any Outlook email thread — so the context you need is always right there when you open the conversation.

**[Install from the Chrome Web Store](#)** <!-- update link once approved -->

---

## Features

- **Pin multiple notes per thread** — add as many notes as you need to any email conversation
- **Drag-to-reorder** — rearrange notes by dragging, or use the keyboard (Arrow Up/Down)
- **Inline editing** — click any note to edit it in place
- **Undo deletions** — accidentally deleted a note? Hit undo before it auto-dismisses (7 seconds)
- **Email list badges** — a small icon appears on any thread in your inbox that has notes attached
- **Pop-out window support** — notes sync between the main Outlook window and any pop-out email windows
- **Keyboard accessible** — full keyboard navigation throughout the panel
- **500-character limit per note** with a storage warning when you're running low

## Privacy

All notes are saved locally using Chrome's built-in `storage` API. Nothing is sent to any server — your notes never leave your device.

**Permissions used:** `storage` only. No access to your tabs, browsing history, or email content.

[Privacy Policy](https://github.com/CoolpantsMcBadass/outlook-sticky-notes-ext/blob/main/privacy-policy.md)

## Supported Outlook Versions

| Version | URL |
|---|---|
| Outlook on the web | `outlook.cloud.microsoft` |
| Outlook 365 | `outlook.office.com` |
| Outlook.com | `outlook.live.com` |

## How It Works

1. Open any email thread in Outlook on the web
2. A sticky note icon appears in the top-right of the reading pane — click it to expand your notes panel
3. Use the **+** button to add a new note
4. Drag notes to reorder them, or click the trash icon to delete
5. Collapse the panel back to the icon when done — the icon shows your note count at a glance

## Installing from Source

1. Clone or download this repository
2. Open Chrome and go to `chrome://extensions`
3. Enable **Developer mode** (toggle in the top right)
4. Click **Load unpacked** and select this folder
5. Open Outlook on the web — the sticky note icon will appear on any email thread

## File Structure

```
outlook-sticky-notes-ext/
├── manifest.json       # Chrome MV3 manifest
├── content.js          # All extension logic (injected into Outlook)
├── styles.css          # Panel and badge styles
├── icons/
│   ├── postit.png      # Pixel art post-it icon (panel collapsed state)
│   ├── icon16.png      # Extension icons for Chrome UI
│   ├── icon32.png
│   ├── icon48.png
│   └── icon128.png
└── privacy-policy.md
```

## Technical Notes

- Built with **Manifest V3** — no background service worker needed
- Uses `match_origin_as_fallback: true` to support Outlook pop-out windows, which render email content in opaque-origin iframes
- Notes are keyed by decoded conversation ID from the URL, with automatic migration from legacy URL-encoded keys
- Email list badges are rebadged dynamically via `MutationObserver` to handle Outlook's virtual scroll