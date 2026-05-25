# Changelog

## [1.0.3] - 2026-05-25

### Fixed
- Notes now load correctly in "Popout only" mode (View > Reading pane > Popout only). Previously the pop-out window's URL is `about:blank` and the main window's URL is just the inbox (no message ID), so `getKey()` fell through to the subject-text fallback and created a mismatched key. Now reads the selected conversation's `data-convid` from the opener's email list DOM, which matches the key used in normal reading-pane mode.
- `isPopout()` now recognizes "Popout only" mode windows (`about:blank` with a same-origin opener), not just the standard pop-out triggered by the email's pop-out icon.

## [1.0.2] - 2026-05-13

### Fixed
- Badges now reliably reappear after Outlook re-renders list rows on hover, scroll, or selection changes. Previously badges could disappear and only return after clicking an email.
- Badge scans are now idempotent — rows that already have a badge are skipped, preventing flicker during frequent updates.
- Fixed stale notes from a previous email thread occasionally remaining visible when switching to a new thread. Outlook navigates via `history.pushState()` which doesn't always trigger a DOM mutation; navigation is now detected via the Navigation API directly.
- Fixed badges not appearing on the list row of the currently open email thread. The reading pane also contains conversation elements with matching IDs; searches are now scoped to the mail list container first to avoid shadowing.
- Fixed injected panel being silently dropped when the user navigates quickly while a panel injection is mid-flight.
- Fixed list observer not reconnecting when Outlook replaces the mail list container element during virtual scroll.
- Badge scans now run immediately when the list observer is first attached, so already-rendered emails are badged without waiting for a list mutation.

## [1.0.0] - 2026-05-10

### Released
- Initial Chrome Web Store submission.
- Sticky notes panel injected into Outlook email threads (outlook.cloud.microsoft, outlook.office.com, outlook.live.com).
- Add, edit, delete, and drag-to-reorder notes per thread.
- Notes keyed by conversation ID — survive page reloads and persist across sessions.
- Collapsed post-it icon shows note count; auto-collapses when last note is deleted.
- Email list badges: small post-it icon on threads that have notes.
- Undo bar for accidental deletes (7-second window).
- 500-character limit with live counter.
- Compose view detection — panel hidden when drafting.
- Pop-out window support.
- Storage usage warning at 80% capacity.
- Keyboard shortcuts: Ctrl/Cmd+Enter to save, Escape to cancel, Arrow keys to reorder.
- Drag-and-drop note reordering.
